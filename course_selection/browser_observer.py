"""Non-destructive inspection of an explicitly supplied CDP browser.

This module is deliberately small: it is the harness-neutral seam used by
adapters and tests.  It never launches a browser, searches for endpoints, or
calls ``Browser.close``.  A connection made here is always borrowed.
"""

from __future__ import annotations

import re
from dataclasses import dataclass
from typing import Any
from urllib.parse import urlsplit

_ENDPOINT = re.compile(r"^(https?://|ws://|wss://)[^\s]+$", re.IGNORECASE)


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


class BorrowedBrowserObserver:
    """Attach to one already-running browser for read-only target inspection."""

    def __init__(self, *, session_id: str = "default") -> None:
        if not session_id or len(session_id) > 128:
            raise ValueError("session_id must be a non-empty bounded value")
        self.session_id = session_id
        self._playwright: Any | None = None
        self._browser: Any | None = None
        self._endpoint = ""

    @property
    def connected(self) -> bool:
        try:
            return self._browser is not None and self._browser.is_connected()
        except Exception:
            return False

    def connect(self, endpoint: str) -> ObserverResult:
        """Connect only to *endpoint*; repeated calls are idempotent."""
        endpoint = str(endpoint or "").strip()
        if not endpoint or not _ENDPOINT.match(endpoint):
            raise ValueError("an explicit HTTP or WebSocket CDP endpoint is required")
        if self.connected:
            if endpoint != self._endpoint:
                return ObserverResult(
                    "already_connected", "borrowed browser is already connected",
                    {"connection": "borrowed", "endpoint_changed": True},
                    ("disconnect before connecting to another endpoint",),
                )
            return ObserverResult("already_connected", "borrowed browser is already connected", self._connection_data())

        from playwright.sync_api import sync_playwright

        playwright = sync_playwright().start()
        try:
            browser = playwright.chromium.connect_over_cdp(endpoint)
        except Exception:
            # Stopping Playwright only tears down this client connection; no
            # browser close is attempted when attachment fails.
            playwright.stop()
            raise
        self._playwright, self._browser, self._endpoint = playwright, browser, endpoint
        return ObserverResult("connected", "connected to borrowed browser", self._connection_data())

    def inspect(self) -> ObserverResult:
        """Inspect the current target set through the public observation seam."""
        return self.inventory()

    def inventory(self) -> ObserverResult:
        if not self.connected:
            return ObserverResult(
                "disconnected", "no borrowed browser connection",
                next_actions=("connect with an explicit CDP endpoint",),
            )
        targets: list[dict[str, Any]] = []
        for context in self._browser.contexts:
            for page in context.pages:
                targets.append(self._page_target(page))
                for frame in page.frames[1:]:
                    targets.append(self._frame_target(frame))
            for worker in getattr(context, "service_workers", []):
                targets.append(self._worker_target(worker, "service_worker"))
            for page in context.pages:
                for worker in getattr(page, "workers", []):
                    targets.append(self._worker_target(worker, "worker"))
        targets.sort(key=lambda item: (item["kind"], item["url_shape"], item.get("frame_depth", 0)))
        return ObserverResult(
            "complete", f"inventoried {len(targets)} browser targets",
            {"connection": "borrowed", "target_count": len(targets), "targets": targets},
        )

    def disconnect(self) -> ObserverResult:
        """Detach this client without closing the remote browser or its tabs."""
        if self._playwright is None:
            return ObserverResult("disconnected", "borrowed browser is already detached", {"connection": "borrowed"})
        playwright, self._playwright = self._playwright, None
        self._browser = None
        self._endpoint = ""
        playwright.stop()
        return ObserverResult("disconnected", "detached from borrowed browser", {"connection": "borrowed"})

    def _connection_data(self) -> dict[str, Any]:
        return {"connection": "borrowed", "detach_only": True}

    @staticmethod
    def _safe_url_shape(url: str) -> str:
        parsed = urlsplit(str(url or ""))
        if not parsed.scheme or not parsed.hostname:
            return "unknown"
        # Never return query, fragment, credentials, or an exact path.  Shape
        # remains useful for distinguishing portal origins without exposing
        # page content or personal path values.
        depth = len([part for part in parsed.path.split("/") if part])
        return f"{parsed.scheme.lower()}://{parsed.hostname.lower()}/<path:{depth}>"

    @classmethod
    def _page_target(cls, page: Any) -> dict[str, Any]:
        return {
            "kind": "page", "url_shape": cls._safe_url_shape(page.url),
            "frame_count": max(0, len(page.frames) - 1),
        }

    @classmethod
    def _frame_target(cls, frame: Any) -> dict[str, Any]:
        depth = 0
        parent = getattr(frame, "parent_frame", None)
        while parent is not None:
            depth += 1
            parent = getattr(parent, "parent_frame", None)
        return {"kind": "frame", "url_shape": cls._safe_url_shape(frame.url), "frame_depth": depth}

    @classmethod
    def _worker_target(cls, worker: Any, kind: str) -> dict[str, Any]:
        return {"kind": kind, "url_shape": cls._safe_url_shape(worker.url)}
