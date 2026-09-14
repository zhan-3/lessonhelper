"""Non-destructive inspection of an explicitly supplied CDP browser.

This module is deliberately small: it is the harness-neutral seam used by
adapters and tests.  It never launches a browser, searches for endpoints, or
calls ``Browser.close``.  A connection made here is always borrowed.
"""

from __future__ import annotations

import hashlib
import logging
import re
import threading
import time
from collections.abc import Callable
from dataclasses import dataclass, field
from typing import Any, ClassVar
from urllib.parse import urlsplit

from .browser_redaction import TraceRedactor

_ENDPOINT = re.compile(r"^(https?://|ws://|wss://)[^\s]+$", re.IGNORECASE)


def _synchronized(method):
    def wrapped(self, *args, **kwargs):
        with self._lock:
            return method(self, *args, **kwargs)
    return wrapped


@dataclass(frozen=True)
class ObserverResult:
    """Stable, adapter-friendly result envelope."""

    status: str
    summary: str
    data: dict[str, Any] | None = None
    warnings: tuple[str, ...] = ()
    next_actions: tuple[str, ...] = ()

    def to_dict(self) -> dict[str, Any]:
        return {
            "status": self.status,
            "summary": self.summary,
            "data": self.data,
            "warnings": list(self.warnings),
            "next_actions": list(self.next_actions),
        }


@dataclass
class _Trace:
    identity: str
    started_targets: set[str]
    events: list[dict[str, Any]] = field(default_factory=list)
    checkpoint: int = 0
    listeners: list[tuple[Any, str, Callable[..., Any]]] = field(default_factory=list)
    dropped_events: int = 0
    redaction_error: str = ""
    missing_evidence: set[str] = field(default_factory=set)
    redactor: TraceRedactor = field(default_factory=TraceRedactor, repr=False)
    started_at: float = field(default_factory=time.monotonic)
    operation_boundary_sequence: int = 0
    operation_boundary_ms: float = 0.0


class BorrowedBrowserObserver:
    """Attach to one already-running browser for read-only target inspection."""

    _owners: ClassVar[dict[str, BorrowedBrowserObserver]] = {}
    _owners_lock = threading.RLock()
    MAX_TARGETS = 100

    def __init__(self, *, session_id: str = "default", max_targets: int = MAX_TARGETS, max_runtime_seconds: float = 600) -> None:
        if not session_id or len(session_id) > 128:
            raise ValueError("session_id must be a non-empty bounded value")
        if not 1 <= max_targets <= self.MAX_TARGETS:
            raise ValueError("max_targets must be between 1 and 100")
        if not 0 < max_runtime_seconds <= 3600:
            raise ValueError("max_runtime_seconds must be greater than zero and at most 3600")
        self.session_id = session_id
        self.max_targets = max_targets
        self.max_runtime_seconds = max_runtime_seconds
        self._lock = threading.RLock()
        self._playwright: Any | None = None
        self._browser: Any | None = None
        self._endpoint = ""
        self._trace: _Trace | None = None
        self._trace_counter = 0
        self.max_events = 500
        self._target_ids: dict[int, str] = {}
        self._loader_counter = 0
        self._request_loaders: dict[int, str] = {}
        self._frame_loaders: dict[str, str] = {}
        self._last_terminal_result: ObserverResult | None = None

    @property
    def connected(self) -> bool:
        try:
            return self._browser is not None and self._browser.is_connected()
        except Exception:
            return False

    @_synchronized
    def connect(self, endpoint: str) -> ObserverResult:
        """Connect only to *endpoint*; repeated calls are idempotent."""
        endpoint = str(endpoint or "").strip()
        if not endpoint or not _ENDPOINT.match(endpoint):
            raise ValueError("an explicit HTTP or WebSocket CDP endpoint is required")
        if (urlsplit(endpoint).hostname or "").lower() not in {"127.0.0.1", "localhost", "::1"}:
            raise ValueError("the CDP endpoint must be loopback-only")
        with self._owners_lock:
            owner = self._owners.get(self.session_id)
            if owner is not None and owner is not self:
                return ObserverResult(
                    "already_connected", "borrowed browser is already connected",
                    owner._connection_data(),
                )
            if self.connected:
                if endpoint != self._endpoint:
                    return ObserverResult(
                        "already_connected", "borrowed browser is already connected",
                        {"connection": "borrowed", "endpoint_changed": True},
                        ("disconnect before connecting to another endpoint",),
                    )
                return ObserverResult("already_connected", "borrowed browser is already connected", self._connection_data())
            # Reserve the session before the potentially slow CDP handshake;
            # another thread cannot create a second borrowed attachment.
            self._owners[self.session_id] = self
        if self.connected:
            if endpoint != self._endpoint:
                return ObserverResult(
                    "already_connected", "borrowed browser is already connected",
                    {"connection": "borrowed", "endpoint_changed": True},
                    ("disconnect before connecting to another endpoint",),
                )
            return ObserverResult("already_connected", "borrowed browser is already connected", self._connection_data())

        if self._playwright is not None:
            stale_playwright, self._playwright = self._playwright, None
            self._browser = None
            stale_playwright.stop()

        from playwright.sync_api import sync_playwright

        playwright = sync_playwright().start()
        try:
            browser = playwright.chromium.connect_over_cdp(endpoint)
        except Exception:
            with self._owners_lock:
                if self._owners.get(self.session_id) is self:
                    del self._owners[self.session_id]
            # Stopping Playwright only tears down this client connection; no
            # browser close is attempted when attachment fails.
            playwright.stop()
            return ObserverResult(
                "failed", "unable to connect to the supplied CDP endpoint",
                warnings=("verify the browser is running and the endpoint is reachable",),
                next_actions=("check the explicit endpoint, then retry once",),
            )
        self._playwright, self._browser, self._endpoint = playwright, browser, endpoint
        with self._owners_lock:
            self._owners[self.session_id] = self
        return ObserverResult("connected", "connected to borrowed browser", self._connection_data())

    @_synchronized
    def start_observation(self) -> ObserverResult:
        if not self.connected:
            return ObserverResult("failed", "cannot observe without a borrowed browser connection", next_actions=("connect first",))
        if self._trace is not None:
            return ObserverResult("already_observing", "observation is already active", {"trace_id": self._trace.identity})
        self._last_terminal_result = None
        self._request_loaders.clear()
        self._frame_loaders.clear()
        self._loader_counter = 0
        self._trace_counter += 1
        trace = _Trace(f"trace-{self._trace_counter}", self._target_keys())
        self._trace = trace
        for context in self._browser.contexts:
            self._listen(context, "request", self._on_request)
            self._listen(context, "response", self._on_response)
            self._listen(context, "requestfailed", self._on_request_failed)
            self._listen(context, "page", self._on_page)
            for page in context.pages:
                self._listen_page(page)
        return ObserverResult("observing", "observation started before external browser activity", {"trace_id": trace.identity, "target_count": len(trace.started_targets)})

    @_synchronized
    def mark_operation_boundary(self) -> ObserverResult:
        if self._trace is None:
            return ObserverResult("failed", "no active observation", next_actions=("start observation first",))
        self.pump_events()
        trace = self._trace
        if trace.operation_boundary_sequence:
            return ObserverResult("already_marked", "operation boundary is already marked", {"trace_id": trace.identity, "operation_boundary_sequence": trace.operation_boundary_sequence})
        trace.operation_boundary_sequence = len(trace.events) + 1
        trace.operation_boundary_ms = round((time.monotonic() - trace.started_at) * 1000, 3)
        return ObserverResult("marked", "operation boundary marked before external activity", {"trace_id": trace.identity, "operation_boundary_sequence": trace.operation_boundary_sequence, "operation_boundary_ms": trace.operation_boundary_ms})

    @_synchronized
    def pump_events(self) -> None:
        if self._trace is None or not self.connected:
            return
        for context in self._browser.contexts:
            pages = context.pages
            if pages:
                try:
                    pages[0].wait_for_timeout(1)
                except Exception as error:
                    self._trace.missing_evidence.add("event_pump")
                    logging.getLogger(__name__).debug("event pump ignored: %s", error)

    @_synchronized
    def enforce_budgets(self) -> ObserverResult | None:
        if self._trace is not None and time.monotonic() - self._trace.started_at > self.max_runtime_seconds:
            self._trace.missing_evidence.add("runtime_budget_exhausted")
            return self.stop_observation()
        return None

    @_synchronized
    def checkpoint(self) -> ObserverResult:
        if self._trace is None:
            return self._last_terminal_result or ObserverResult("failed", "no active observation", next_actions=("start observation first",))
        if time.monotonic() - self._trace.started_at > self.max_runtime_seconds:
            self._trace.missing_evidence.add("runtime_budget_exhausted")
            return self.stop_observation()
        return self._trace_delta(final=False)

    @_synchronized
    def cancel_observation(self) -> ObserverResult:
        if self._trace is None:
            return ObserverResult("cancelled", "no active observation", {"connection": "borrowed"})
        stopped = self.stop_observation()
        result = ObserverResult("cancelled", "observation cancelled locally", stopped.data, stopped.warnings, ("the borrowed browser remains open",))
        self._last_terminal_result = result
        return result

    @_synchronized
    def stop_observation(self) -> ObserverResult:
        if self._trace is None:
            return self._last_terminal_result or ObserverResult("stopped", "no active observation", {"delta": {"events": [], "target_changes": []}})
        result = self._trace_delta(final=True)
        trace = self._trace
        for target, event_name, callback in trace.listeners:
            try:
                target.remove_listener(event_name, callback)
            except Exception as error:
                # A target may have detached concurrently; the trace is still
                # discarded and the borrowed browser remains untouched.
                logging.getLogger(__name__).debug("listener cleanup ignored: %s", error)
        trace.redactor.close()
        self._trace = None
        self._last_terminal_result = result
        return result

    def _listen(self, target: Any, event_name: str, callback: Callable[..., Any]) -> None:
        target.on(event_name, callback)
        if self._trace is not None:
            self._trace.listeners.append((target, event_name, callback))

    def _redaction_failed(self, error: Exception) -> None:
        if self._trace is not None:
            self._trace.redaction_error = str(error)[:160]

    def _record(self, event: dict[str, Any]) -> None:
        trace = self._trace
        if trace is None:
            return
        if len(trace.events) >= self.max_events:
            trace.dropped_events += 1
            return
        if trace.redaction_error:
            return
        try:
            safe_event = trace.redactor.redact_value(event)
        except (RuntimeError, TypeError, ValueError) as error:
            trace.redaction_error = str(error)[:160]
            return
        trace.events.append({"sequence": len(trace.events) + 1, "elapsed_ms": round((time.monotonic() - trace.started_at) * 1000, 3), **safe_event})

    def _loader_identity(self, request: Any) -> str:
        existing = self._request_loaders.get(id(request))
        if existing:
            return existing
        redirected = getattr(request, "redirected_from", None)
        if redirected is not None and id(redirected) in self._request_loaders:
            identity = self._request_loaders[id(redirected)]
        else:
            context = self._request_context(request)
            frame_identity = str(context["frame_identity"])
            if str(getattr(request, "resource_type", "")) == "document" or frame_identity not in self._frame_loaders:
                self._loader_counter += 1
                identity = hashlib.sha256(f"{self.session_id}:loader:{self._loader_counter}".encode()).hexdigest()[:16]
                self._frame_loaders[frame_identity] = identity
            else:
                identity = self._frame_loaders[frame_identity]
        self._request_loaders[id(request)] = identity
        return identity

    def _request_context(self, request: Any) -> dict[str, Any]:
        frame = getattr(request, "frame", None)
        if frame is None:
            target_identity = "unknown"
            frame_identity = "unknown"
        elif getattr(frame, "parent_frame", None) is None:
            page = getattr(frame, "page", None)
            target_identity = self._target_identity_for(page, "page") if page is not None else self._target_identity_for(frame, "frame")
            frame_identity = self._target_identity_for(frame, "frame")
        else:
            target_identity = self._target_identity_for(frame, "frame")
            frame_identity = target_identity
        resource_type = str(getattr(request, "resource_type", "unknown"))
        initiator = "worker" if resource_type in {"script", "worker"} and frame is None else "frame" if frame is not None and getattr(frame, "parent_frame", None) is not None else "page"
        return {"frame_identity": frame_identity, "target_identity": target_identity, "initiator_class": initiator}

    def _redirect_context(self, request: Any) -> dict[str, Any]:
        chain = []
        current = getattr(request, "redirected_from", None)
        seen: set[int] = set()
        while current is not None and id(current) not in seen and len(chain) < 10:
            seen.add(id(current))
            chain.append(self._safe_url_shape(current.url))
            current = getattr(current, "redirected_from", None)
        chain.reverse()
        if chain:
            chain.append(self._safe_url_shape(request.url))
        return {"redirect_hop_count": len(chain) - 1 if chain else 0, "redirect_path_shapes": chain}

    def _on_request(self, request: Any) -> None:
        redirected = getattr(request, "redirected_from", None)
        try:
            post_data = getattr(request, "post_data", None)
            content_type = getattr(request, "headers", {}).get("content-type", "")
            safe_body = self._trace.redactor.redact_body(post_data, content_type) if self._trace is not None and post_data is not None else None
            safe_headers = self._trace.redactor.redact_headers(getattr(request, "headers", {})) if self._trace is not None else {}
            self._record({"kind": "request", "method": request.method, "url_shape": self._safe_url_shape(request.url), "resource_type": request.resource_type, "redirected_from": self._safe_url_shape(redirected.url) if redirected else None, **self._redirect_context(request), "request_body": safe_body, "headers": safe_headers, "loader_identity": self._loader_identity(request), **self._request_context(request)})
        except (RuntimeError, TypeError, ValueError) as error:
            self._redaction_failed(error)

    def _on_response(self, response: Any) -> None:
        try:
            request = response.request
            headers = self._trace.redactor.redact_headers(response.headers) if self._trace is not None else {}
            self._record({"kind": "response", "method": request.method, "url_shape": self._safe_url_shape(response.url), "status": response.status, "resource_type": request.resource_type, "headers": headers, "loader_identity": self._loader_identity(request), **self._request_context(request)})
        except (RuntimeError, TypeError, ValueError) as error:
            self._redaction_failed(error)

    def _on_request_failed(self, request: Any) -> None:
        try:
            self._record({"kind": "request_failed", "method": request.method, "url_shape": self._safe_url_shape(request.url), "resource_type": request.resource_type, "loader_identity": self._loader_identity(request), **self._request_context(request)})
        except (RuntimeError, TypeError, ValueError) as error:
            if self._trace is not None:
                self._trace.missing_evidence.add("failed_request_provenance")
            logging.getLogger(__name__).debug("failed request provenance incomplete: %s", error)

    def _listen_page(self, page: Any) -> None:
        self._listen(page, "console", self._on_console)
        self._listen(page, "framenavigated", self._on_frame_navigated)
        self._listen(page, "frameattached", self._on_frame_attached)
        self._listen(page, "framedetached", self._on_frame_detached)
        self._listen(page, "close", lambda: self._on_page_closed(page))

    def _on_page(self, page: Any) -> None:
        try:
            self._listen_page(page)
            target = self._page_target(page)
            self._record({"kind": "target_created", "target": target, "navigation_state": target["navigation_state"]})
        except Exception as error:
            if self._trace is not None:
                self._trace.missing_evidence.add("new_page_attachment")
            logging.getLogger(__name__).debug("new page attachment incomplete: %s", error)

    def _record_frame_event(self, kind: str, frame: Any) -> None:
        try:
            descriptor = self._frame_target(frame)
            self._record({"kind": kind, "target": descriptor, "navigation_state": descriptor["navigation_state"]})
        except Exception as error:
            if self._trace is not None:
                self._trace.missing_evidence.add(kind)
            logging.getLogger(__name__).debug("frame topology evidence incomplete: %s", error)

    def _on_frame_navigated(self, frame: Any) -> None:
        self._record_frame_event("target_navigated", frame)

    def _on_frame_attached(self, frame: Any) -> None:
        self._record_frame_event("target_attached", frame)

    def _on_frame_detached(self, frame: Any) -> None:
        self._record_frame_event("target_detached", frame)

    def _on_page_closed(self, page: Any) -> None:
        try:
            self._record({"kind": "target_closed", "target": self._page_target(page)})
        except Exception as error:
            if self._trace is not None:
                self._trace.missing_evidence.add("closed_page_descriptor")
            logging.getLogger(__name__).debug("closed page descriptor incomplete: %s", error)

    def _on_console(self, message: Any) -> None:
        try:
            values = [argument.json_value() for argument in getattr(message, "args", [])]
            safe_values = self._trace.redactor.redact_console(values) if self._trace is not None else []
            self._record({"kind": "console", "type": getattr(message, "type", "unknown"), "values": safe_values})
        except (RuntimeError, TypeError, ValueError) as error:
            self._redaction_failed(error)

    def _target_keys(self) -> set[str]:
        result = self.inventory()
        if not result.data:
            return set()
        return {str(item.get("target_identity") or f"{item['kind']}:{item['url_shape']}") for item in result.data.get("targets", [])}

    def _trace_delta(self, *, final: bool) -> ObserverResult:
        trace = self._trace
        assert trace is not None
        # Sync Playwright dispatches protocol events while its connection is
        # pumped. A bounded turn makes externally generated events observable
        # before the delta is materialized.
        self.pump_events()
        events = [] if trace.redaction_error else trace.events[trace.checkpoint:]
        trace.checkpoint = len(trace.events)
        inventory = self.inventory()
        inventory_targets = inventory.data.get("targets", []) if inventory.data else []
        current = {str(item.get("target_identity")) for item in inventory_targets}
        added = sorted(current - trace.started_targets)
        removed = sorted(trace.started_targets - current)
        status = "complete" if trace.dropped_events == 0 and not trace.redaction_error and not trace.missing_evidence else "partial"
        warnings = []
        if trace.dropped_events:
            warnings.append("event limit reached")
        if trace.redaction_error:
            warnings.append("redaction failed; evidence publication was blocked")
        if trace.missing_evidence:
            warnings.append("required target evidence is incomplete")
        next_actions = ("repeat one bounded observation for the missing evidence",) if warnings else ()
        return ObserverResult("stopped" if final and not warnings else status, "observation stopped" if final else "observation checkpoint", {"trace_id": trace.identity, "events": events, "inventory_targets": inventory_targets, "operation_boundary_sequence": trace.operation_boundary_sequence, "operation_boundary_ms": trace.operation_boundary_ms, "target_changes": {"added": added, "removed": removed}, "event_count": len(events), "dropped_events": trace.dropped_events, "missing_evidence": sorted(trace.missing_evidence)}, tuple(warnings), next_actions)

    @_synchronized
    def inspect(self) -> ObserverResult:
        """Inspect the current target set through the public observation seam."""
        return self.inventory()

    @_synchronized
    def inventory(self) -> ObserverResult:
        if not self.connected:
            return ObserverResult(
                "disconnected", "no borrowed browser connection",
                next_actions=("connect with an explicit CDP endpoint",),
            )
        targets: list[dict[str, Any]] = []
        limit_reached = False

        def append_target(target: dict[str, Any]) -> bool:
            nonlocal limit_reached
            targets.append(target)
            if len(targets) > self.max_targets:
                limit_reached = True
                return False
            return True

        for context in self._browser.contexts:
            context_identity = self._target_identity_for(context, "context")
            if not append_target({
                "kind": "context", "url_shape": "browser-context",
                "target_identity": context_identity, "parent_identity": None,
                "relationship": "browser_context", "navigation_state": "active",
                "capabilities": ["pages", "service_workers", "network"],
                "semantic_signature": {"kind": "context", "relationship": "browser_context", "depth": 0, "origin_class": "browser_internal"},
            }):
                break
            for page in context.pages:
                if not append_target(self._page_target(page)):
                    break
                for frame in page.frames[1:]:
                    if not append_target(self._frame_target(frame)):
                        break
                if limit_reached:
                    break
            if limit_reached:
                break
            for worker in getattr(context, "service_workers", []):
                if not append_target(self._worker_target(worker, "service_worker", context_identity)):
                    break
            if limit_reached:
                break
            for page in context.pages:
                page_identity = self._target_identity_for(page, "page")
                for worker in getattr(page, "workers", []):
                    if not append_target(self._worker_target(worker, "worker", page_identity)):
                        break
                if limit_reached:
                    break
            if limit_reached:
                break
        targets.sort(key=lambda item: (item["kind"], item["url_shape"], item.get("frame_depth", 0)))
        truncated = limit_reached
        targets = targets[: self.max_targets]
        return ObserverResult(
            "partial" if truncated else "complete",
            f"inventoried {len(targets)} browser targets",
            {"connection": "borrowed", "target_count": len(targets), "targets": targets},
            ("target inventory limit reached",) if truncated else (),
            ("narrow the browser target set and inspect again",) if truncated else (),
        )

    @_synchronized
    def shutdown(self) -> ObserverResult:
        """Invalidate local trace state and detach idempotently on reload/exit."""
        if self._trace is not None:
            self.stop_observation()
        return self.disconnect()

    @_synchronized
    def disconnect(self) -> ObserverResult:
        """Detach this client without closing the remote browser or its tabs."""
        if self._trace is not None:
            self.stop_observation()
        if self._playwright is None:
            return ObserverResult("disconnected", "borrowed browser is already detached", {"connection": "borrowed"})
        playwright, self._playwright = self._playwright, None
        self._browser = None
        self._endpoint = ""
        with self._owners_lock:
            if self._owners.get(self.session_id) is self:
                del self._owners[self.session_id]
        playwright.stop()
        return ObserverResult("disconnected", "detached from borrowed browser", {"connection": "borrowed"})

    def _connection_data(self) -> dict[str, Any]:
        return {"connection": "borrowed", "detach_only": True}

    @staticmethod
    def _origin_class(url: str, parent_url: str | None = None) -> str:
        parsed = urlsplit(str(url or ""))
        if str(url or "") in {"", "about:blank"}:
            return "unresolved"
        if parsed.scheme in {"data", "blob", "about"} or not parsed.hostname:
            return "opaque_origin"
        host = parsed.hostname.lower()
        if host == "localhost" or host == "::1" or host.startswith("127."):
            return "loopback"
        if parent_url:
            parent = urlsplit(str(parent_url))
            if parent.hostname:
                current_origin = (parsed.scheme.lower(), host, parsed.port)
                parent_origin = (parent.scheme.lower(), parent.hostname.lower(), parent.port)
                return "same_origin" if current_origin == parent_origin else "cross_origin"
        return "web"

    @staticmethod
    def _safe_url_shape(url: str) -> str:
        parsed = urlsplit(str(url or ""))
        if not parsed.scheme or not parsed.hostname:
            return "unknown"
        # Never return query, fragment, credentials, or an exact path.  Shape
        # remains useful for distinguishing portal origins without exposing
        # page content or personal path values.
        depth = len([part for part in parsed.path.split("/") if part])
        host = parsed.hostname.lower()
        host_fingerprint = hashlib.sha256(host.encode("utf-8")).hexdigest()[:12]
        return f"{parsed.scheme.lower()}://<host:{host_fingerprint}>/<path:{depth}>"

    def _target_identity_for(self, target: Any, kind: str) -> str:
        key = id(target)
        identity = self._target_ids.get(key)
        if identity is None:
            implementation = getattr(target, "_impl_obj", None)
            opaque_marker = str(getattr(implementation, "_guid", "")) or str(key)
            identity = hashlib.sha256(f"{self.session_id}:{kind}:{opaque_marker}".encode()).hexdigest()[:16]
            self._target_ids[key] = identity
        return identity

    def _page_target(self, page: Any) -> dict[str, Any]:
        url_shape = self._safe_url_shape(page.url)
        opener = getattr(page, "opener", None)
        opener_page = opener() if callable(opener) else None
        context_value = getattr(page, "context", None)
        context = context_value() if callable(context_value) else context_value
        if opener_page is not None:
            parent_identity = self._target_identity_for(opener_page, "page")
        elif context is not None:
            parent_identity = self._target_identity_for(context, "context")
        else:
            parent_identity = None
        navigation_state = "unresolved_blank" if str(page.url) in {"", "about:blank"} else "committed"
        relationship = "popup" if opener_page is not None else "top_level"
        return {
            "kind": "page", "url_shape": url_shape,
            "target_identity": self._target_identity_for(page, "page"),
            "parent_identity": parent_identity,
            "relationship": relationship,
            "navigation_state": navigation_state,
            "capabilities": ["document", "frames", "network"],
            "frame_count": max(0, len(page.frames) - 1),
            "semantic_signature": {"kind": "page", "relationship": relationship, "depth": 0, "origin_class": self._origin_class(page.url)},
        }

    def _frame_target(self, frame: Any) -> dict[str, Any]:
        depth = 0
        parent = getattr(frame, "parent_frame", None)
        while parent is not None:
            depth += 1
            parent = getattr(parent, "parent_frame", None)
        url_shape = self._safe_url_shape(frame.url)
        parent_frame = getattr(frame, "parent_frame", None)
        if parent_frame is None:
            parent_identity = None
        elif getattr(parent_frame, "parent_frame", None) is None:
            page = getattr(frame, "page", None)
            parent_identity = self._target_identity_for(page, "page") if page is not None else self._target_identity_for(parent_frame, "frame")
        else:
            parent_identity = self._target_identity_for(parent_frame, "frame")
        relationship = "child_frame" if parent_identity else "main_frame"
        parent_url = str(getattr(parent_frame, "url", "")) if parent_frame is not None else None
        return {
            "kind": "frame", "url_shape": url_shape, "frame_depth": depth,
            "target_identity": self._target_identity_for(frame, "frame"),
            "parent_identity": parent_identity,
            "relationship": relationship,
            "navigation_state": "unresolved_blank" if str(frame.url) in {"", "about:blank"} else "committed",
            "capabilities": ["document", "network"],
            "semantic_signature": {"kind": "frame", "relationship": relationship, "depth": depth, "origin_class": self._origin_class(frame.url, parent_url)},
        }

    def _worker_target(self, worker: Any, kind: str, parent_identity: str | None = None) -> dict[str, Any]:
        url_shape = self._safe_url_shape(worker.url)
        return {
            "kind": kind, "url_shape": url_shape,
            "target_identity": self._target_identity_for(worker, kind),
            "parent_identity": parent_identity, "relationship": "execution_context",
            "navigation_state": "committed", "capabilities": ["network", "runtime"],
            "semantic_signature": {"kind": kind, "relationship": "execution_context", "depth": 1 if parent_identity else 0, "origin_class": self._origin_class(worker.url)},
        }
