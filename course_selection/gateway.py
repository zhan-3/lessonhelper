"""Framework-independent boundary for read-only university observations."""

from __future__ import annotations

import json
import time
from dataclasses import asdict
from datetime import datetime, timezone
from pathlib import Path
from typing import Any, Callable, Protocol
from urllib.parse import urlsplit

Progress = Callable[[str, dict[str, Any]], None]
Cancelled = Callable[[], bool]

_ACADEMIC_PROXY_PATH = (
    "/http/77726476706e69737468656265737421fae0558f693861446900c7a99c406d3667/"
)


class AcademicGateway(Protocol):
    def launch_shell(self, url: str, progress: Progress, cancelled: Cancelled) -> None: ...
    def connect(self, progress: Progress, cancelled: Cancelled) -> None: ...
    def refresh_selection(self, context: dict[str, Any], progress: Progress, cancelled: Cancelled) -> dict[str, Any]: ...
    def refresh_timetable(self, context: dict[str, Any], progress: Progress, cancelled: Cancelled) -> dict[str, Any]: ...
    def observe_timetable(self, request, progress: Progress, cancelled: Cancelled): ...
    def refresh_progress(self, context: dict[str, Any], progress: Progress, cancelled: Cancelled) -> dict[str, Any]: ...
    def observe_navigation(self, context: dict[str, Any], progress: Progress, cancelled: Cancelled, finished: Cancelled) -> dict[str, Any]: ...
    def reset_login(self) -> None: ...
    def close(self) -> None: ...
    def poll(self) -> None: ...


class UnconfirmedAcademicGateway:
    """Safe default until live read contracts have been independently verified."""

    def launch_shell(self, url: str, progress: Progress, cancelled: Cancelled) -> None:
        progress("connecting", {"message": "workbench shell is not configured"})

    def connect(self, progress: Progress, cancelled: Cancelled) -> None:
        progress("connecting", {"message": "academic gateway requires explicit configuration"})

    def refresh_selection(self, context: dict[str, Any], progress: Progress, cancelled: Cancelled) -> dict[str, Any]:
        return {"status": "interface_unconfirmed", "sections": []}

    def refresh_timetable(self, context: dict[str, Any], progress: Progress, cancelled: Cancelled) -> dict[str, Any]:
        progress("interface_unconfirmed", {"target": "timetable", "message": "timetable read interface is not configured"})
        return {"status": "interface_unconfirmed", "entries": []}

    def observe_timetable(self, request, progress: Progress, cancelled: Cancelled):
        from .deep_observation import TimetableObservationResult
        result = self.refresh_timetable(request.context, progress, cancelled)
        return TimetableObservationResult.incomplete(str(result.get("status", "interface_unconfirmed")))

    def refresh_progress(self, context: dict[str, Any], progress: Progress, cancelled: Cancelled) -> dict[str, Any]:
        progress("interface_unconfirmed", {"target": "progress", "message": "grade read interface is not configured"})
        return {"status": "interface_unconfirmed", "report": None}

    def observe_navigation(self, context: dict[str, Any], progress: Progress, cancelled: Cancelled, finished: Cancelled) -> dict[str, Any]:
        return {"status": "interface_unconfirmed", "report": {"events": [], "blocked_requests": []}}

    def reset_login(self) -> None:
        self.close()

    def close(self) -> None:
        return None

    def poll(self) -> None:
        return None


class PlaywrightAcademicGateway(UnconfirmedAcademicGateway):
    """Long-lived bundled-Chromium session, owned by the observation worker."""

    def __init__(self, profile_root=".private/course-progress", workspace_root=".private/academic-selection", portal_url: str | None = None, on_browser_closed: Callable[[], None] | None = None):
        from pathlib import Path
        from course_progress.explorer import DEFAULT_PORTAL_URL
        self.profile_root = Path(profile_root)
        self.workspace_root = Path(workspace_root)
        self.portal_url = portal_url or DEFAULT_PORTAL_URL
        self._playwright = None
        self._session = None
        self._shell_url = ""
        self._shell_page = None
        self._academic_page = None
        self._on_browser_closed = on_browser_closed or (lambda: None)
        self._closing = False
        self._browser_closed_notified = False
        self.browser_launches = 0

    @staticmethod
    def _is_loopback(url: str) -> bool:
        return (urlsplit(url).hostname or "").lower() in {"127.0.0.1", "localhost", "::1"}

    def _academic_pages(self) -> list[Any]:
        if self._session is None or self._session.context is None:
            return []
        return [page for page in self._session.context.pages if not self._is_loopback(page.url)]

    @staticmethod
    def _page_is_closed(page: Any) -> bool:
        checker = getattr(page, "is_closed", None)
        return bool(checker()) if checker is not None else False

    def _ensure_browser(self, progress: Progress) -> None:
        if self._session is not None and self._session.context is not None:
            try:
                _ = self._session.context.pages
                return
            except Exception:
                self.close()
        progress("connecting", {"message": "starting bundled Chromium"})
        from playwright.sync_api import sync_playwright
        from course_progress.session import AcademicBrowserSession
        self._playwright = sync_playwright().start()
        self._session = AcademicBrowserSession(
            self._playwright, browser_name="chromium", profile_root=self.profile_root,
            persistent=True,
        )
        self._session.__enter__()
        self.browser_launches += 1
        self._session.context.on("close", self._browser_closed)

    def _browser_closed(self, *_args) -> None:
        if not self._closing and not self._browser_closed_notified:
            self._browser_closed_notified = True
            self._on_browser_closed()

    def launch_shell(self, url: str, progress: Progress, cancelled: Cancelled) -> None:
        if not self._is_loopback(url):
            raise ValueError("workbench shell must use a loopback URL")
        self._ensure_browser(progress)
        self._shell_url = url.rstrip("/") + "/"
        context = self._session.context
        pages = context.pages
        page = next((item for item in pages if self._is_loopback(item.url)), None)
        if page is None:
            page = context.new_page()
        page.goto(self._shell_url, wait_until="domcontentloaded", timeout=30_000)
        page.bring_to_front()
        self._shell_page = page
        progress("connecting", {"message": "visible Chromium workbench opened"})

    def connect(self, progress: Progress, cancelled: Cancelled) -> None:
        self._ensure_browser(progress)
        context = self._session.context
        page = getattr(self, "_academic_page", None)
        if page is None:
            candidates = self._academic_pages()
            page = candidates[-1] if candidates else None
        if page is None or self._page_is_closed(page):
            page = next((item for item in context.pages if not self._is_loopback(item.url)), None)
        if page is None:
            page = context.new_page()
        self._academic_page = page
        try:
            # Authentication is an intentional, user-visible pause.  Publish
            # this before entering the blocking browser wait so task observers
            # can render the state while the user operates the browser.
            progress("waiting_for_authentication", {"message": "waiting for authentication in the bundled browser"})
            self._session.open_authenticated(self.portal_url, timeout_seconds=600, page=page)
            page.bring_to_front()
        except TimeoutError:
            progress("waiting_for_authentication", {"message": "authentication timed out; browser remains available"})
            raise

    def refresh_selection(self, context: dict[str, Any], progress: Progress, cancelled: Cancelled) -> dict[str, Any]:
        if self._session is None or self._session.context is None:
            raise RuntimeError("academic session is disconnected")
        from types import SimpleNamespace

        from .discovery import TARGET_SELECTION, InterfaceDiscovery
        from .selection_query import (
            SelectionContractError,
            VerifiedSelectionQueryAdapter,
        )

        categories = tuple(context.get("allowed_categories", ()))
        semester_label = str(context.get("semester_label", ""))
        if not categories or not semester_label:
            return {"status": "no_matching_round", "sections": []}
        raw_windows = context.get("allowed_windows", {})

        # Normal refreshes use the versioned, verified read contract directly.
        # The browser is retained for authentication, cookies, dynamic form
        # state and same-origin fetch; no selection-menu controls are clicked.
        page = getattr(self, "_academic_page", None)
        if page is None:
            try:
                pages = self._academic_pages()
            except (AttributeError, TypeError):
                pages = []
            page = pages[-1] if pages else None
        if page is not None and not self._page_is_closed(page):
            try:
                result = VerifiedSelectionQueryAdapter().read(
                    page,
                    categories=categories,
                    semester_label=semester_label,
                    allowed_windows=raw_windows,
                    progress=progress,
                    cancelled=cancelled,
                    authenticate=lambda url, target_page: self._session.open_authenticated(
                        url,
                        timeout_seconds=int(context.get("login_timeout_seconds", 600)),
                        page=target_page,
                    ),
                )
                result.setdefault("source_kind", "verified-selection-api")
                return result
            except SelectionContractError as error:
                progress(
                    "reading",
                    {
                        "target": "selection",
                        "mode": "contract-discovery",
                        "message": str(error),
                    },
                )

        # Only a missing or changed verified contract enters discovery.  A
        # transient category/page failure returned by the adapter remains an
        # incomplete attempt and is never repeated through another path.
        windows = {
            code: tuple(SimpleNamespace(**item) for item in items)
            for code, items in raw_windows.items()
        }
        navigator = InterfaceDiscovery(
            target=TARGET_SELECTION, output_root=self.workspace_root,
            allowed_selection_categories=categories,
            allowed_selection_windows=windows,
            notice_semester=semester_label,
        )
        browser_context = self._session.context
        response_handler = navigator._handle_response
        route_handler = navigator._guard_route
        browser_context.on("response", response_handler)
        browser_context.route("**/*", route_handler)
        try:
            progress("reading", {"target": "selection", "mode": "contract-discovery"})
            report = navigator.run(browser_context, max_clicks=8, wait_seconds=30)
        finally:
            # Lightweight test and embedding contexts may not expose unroute;
            # real Playwright BrowserContext instances do.
            unroute = getattr(browser_context, "unroute", None)
            if unroute is not None:
                unroute("**/*", route_handler)
            remove_listener = getattr(browser_context, "remove_listener", None)
            if remove_listener is not None:
                remove_listener("response", response_handler)
        if not report.selection_query_payload:
            return {"status": "interface_unconfirmed", "sections": []}
        # Discovery keeps a JSON artifact for read-only diagnostics, but the
        # refresh path consumes its structured result directly.  SQLite is
        # written by ObservationService after this method returns, so a stale
        # diagnostic file cannot be published as a snapshot.
        payload = report.selection_query_payload
        queries = payload.get("queries", [])
        complete = bool(queries) and all(item.get("complete") for item in queries)
        sections = [section for query in queries for section in query.get("sections", [])]
        return {
            "status": "complete" if complete else "incomplete",
            "source_kind": "discovered-selection-api",
            "sections": sections,
            "queries": queries,
        }

    def refresh_progress(self, context: dict[str, Any], progress: Progress, cancelled: Cancelled) -> dict[str, Any]:
        """Collect grade records and return only a score-free progress report."""
        if self._session is None or self._session.context is None:
            raise RuntimeError("academic session is disconnected")
        from course_progress.collector import PlaywrightGradeCollector, wait_for_reauthentication
        from course_progress.progress import RequirementBaseline, evaluate_progress, parse_requirements

        page = getattr(self, "_academic_page", None)
        if page is None:
            pages = self._academic_pages()
            page = pages[-1] if pages else None
        if page is None or self._page_is_closed(page):
            return {"status": "entry_unreachable", "report": None}

        requirements_path = Path(__file__).resolve().parents[1] / "docs" / "校园培养方案解读（2026年版）.md"
        if not requirements_path.is_file():
            return {"status": "interface_unconfirmed", "report": None}
        baseline = RequirementBaseline(
            version=str(context.get("baseline_version", "guide-2026")),
            requirements=parse_requirements(requirements_path),
            category_mapping={
                "本专业选修": "major_elective",
                "外专业选修": "outside_major_elective",
                "文理通识-文化素质教育课": "cultural_quality",
                "创新研修课": "innovation",
                "社会实践": "social_practice",
            },
        )
        timeout = int(context.get("login_timeout_seconds", 600))

        def on_page_data(semester, page_number, page_count, records):
            progress("reading", {
                "target": "progress",
                "semester": semester.label,
                "page": page_number,
                "page_count": page_count,
                "records": len(records),
            })

        collection = PlaywrightGradeCollector(
            page_size=int(context.get("page_size", 20))
        ).collect(
            page,
            frame_timeout_seconds=timeout,
            on_session_expired=lambda: wait_for_reauthentication(page, timeout),
            on_page_data=on_page_data,
            is_cancelled=cancelled,
        )
        report = evaluate_progress(collection.records, baseline)
        now = datetime.now(timezone.utc).isoformat(timespec="seconds")
        return {
            "status": "complete" if collection.complete else "incomplete",
            "source_kind": "academic",
            "source_at": now,
            "term": str(context.get("term", "")),
            "report": {
                "generated_at": now,
                "baseline_version": report.baseline_version,
                "data_complete": collection.complete,
                "semesters": [asdict(item) for item in collection.semesters],
                "collection_failures": [asdict(item) for item in collection.failures],
                "progress": [
                    {
                        "key": item.requirement.key,
                        "label": item.requirement.label,
                        "required_credits": item.requirement.minimum_credits,
                        "completed_credits": item.completed_credits,
                        "remaining_credits": item.remaining_credits,
                        "courses": [asdict(course) for course in item.courses],
                    }
                    for item in report.progress
                ],
                "conflicts": [asdict(item) for item in report.conflicts],
                "unclassified_courses": [asdict(course) for course in report.unclassified_courses],
            },
        }

    def _refresh_timetable_payload(self, context: dict[str, Any], progress: Progress, cancelled: Cancelled) -> dict[str, Any]:
        trace_requests: list[dict[str, Any]] = []
        if self._session is None or self._session.context is None:
            raise RuntimeError("academic session is disconnected")
        from .personal_timetable import parse_personal_timetable_html, personal_timetable_parameters
        from course_progress.collector import resolve_academic_url

        page = getattr(self, "_academic_page", None)
        if page is None:
            candidates = self._academic_pages()
            page = candidates[-1] if candidates else None
        if page is None or self._page_is_closed(page):
            return {"status": "entry_unreachable", "entries": [], "_trace_requests": trace_requests}
        trace_requests.append({"method": "GET", "url": page.url, "resource_type": "document"})

        def value(name: str) -> str:
            locator = page.locator(f"[name='{name}']").first
            if locator.count() == 0:
                return ""
            try:
                return str(locator.input_value() or "")
            except Exception:
                return ""

        try:
            parameters = personal_timetable_parameters(fhlj=value("fhlj"), xnxq=value("xnxq"))
        except ValueError as error:
            progress("interface_unconfirmed", {"target": "timetable", "message": str(error)})
            return {"status": "interface_unconfirmed", "entries": [], "reason": str(error), "_trace_requests": trace_requests}

        endpoint = resolve_academic_url(page.url, "/kbcx/queryXszkb")
        trace_requests.append({
            "method": "POST", "url": endpoint, "resource_type": "fetch",
            "post_data": "&".join(f"{name}={value}" for name, value in parameters.items()),
        })
        progress("reading", {"target": "timetable", "endpoint": "/kbcx/queryXszkb"})
        response = self._session.context.request.post(endpoint, form=parameters, timeout=60_000)
        if response.status != 200:
            return {"status": "entry_unreachable", "entries": [], "http_status": response.status, "_trace_requests": trace_requests}
        html = response.text()
        term = str(context.get("term") or parameters["xnxq"])
        try:
            entries = parse_personal_timetable_html(html, term=term)
        except ValueError as error:
            progress("interface_unconfirmed", {"target": "timetable", "message": str(error)})
            return {"status": "interface_unconfirmed", "entries": [], "reason": str(error), "_trace_requests": trace_requests}
        from .timetable import timetable_snapshot_payload
        payload = timetable_snapshot_payload(
            entries, source_name="/kbcx/queryXszkb", source_kind="personal-timetable-api",
        )
        return {"status": "complete", "term": term, "source_kind": "personal-timetable-api", "_trace_requests": trace_requests, **payload}

    def observe_timetable(self, request, progress: Progress, cancelled: Cancelled):
        """Perform one typed personal-timetable observation with a safe trace."""
        from .deep_observation import AcademicRequestTrace, TimetableObservationResult

        context = request.context
        result = self._refresh_timetable_payload(context, progress, cancelled)
        # BrowserContext.request does not emit page-level request events. The
        # verified adapter records each request at the exact call site, before
        # it is sent, including failed attempts and the resolved proxy URL.
        trace = AcademicRequestTrace.from_requests(result.pop("_trace_requests", ()))
        if cancelled():
            return TimetableObservationResult.cancelled(trace=trace)
        if result.get("status") != "complete":
            return TimetableObservationResult.incomplete(str(result.get("reason") or result.get("status")), trace=trace)
        return TimetableObservationResult.complete(
            term=str(result.get("term") or request.term), entries=list(result.get("entries", ())), trace=trace,
        )

    def refresh_timetable(self, context: dict[str, Any], progress: Progress, cancelled: Cancelled) -> dict[str, Any]:
        """Compatibility adapter for callers not yet on the typed seam."""
        return self._refresh_timetable_payload(context, progress, cancelled)

    def observe_navigation(self, context: dict[str, Any], progress: Progress, cancelled: Cancelled, finished: Cancelled) -> dict[str, Any]:
        """Observe user-directed navigation in the existing browser session."""
        if self._session is None or self._session.context is None:
            raise RuntimeError("academic session is disconnected")
        from playwright.sync_api import Error
        from .manual_observation import ManualObservationPolicy, summarize_json_structure, url_evidence

        browser_context = self._session.context
        policy = ManualObservationPolicy(context.get("allowed_post_requests", ()))
        events: list[dict[str, Any]] = []
        blocked_requests: list[dict[str, Any]] = []
        maximum_events = min(max(int(context.get("maximum_events", 300)), 1), 1000)
        timeout_seconds = min(max(int(context.get("timeout_seconds", 1800)), 30), 7200)

        def append(event: dict[str, Any]) -> None:
            if len(events) < maximum_events:
                events.append({"sequence": len(events) + 1, **event})

        for page in self._academic_pages():
            append({"kind": "page", **url_evidence(page.url)})

        def guard(route) -> None:
            request = route.request
            if self._is_loopback(request.url):
                route.continue_()
                return
            if request.resource_type not in {"document", "xhr", "fetch"}:
                route.continue_()
                return
            allowed, evidence = policy.inspect_request(
                request.method,
                request.url,
                request.post_data,
                request.headers.get("content-type", ""),
                request.resource_type,
            )
            append({"kind": "request", **evidence})
            if allowed:
                route.continue_()
            else:
                blocked_requests.append(evidence)
                route.abort("blockedbyclient")

        def record_response(response) -> None:
            request = response.request
            if self._is_loopback(response.url):
                return
            if request.resource_type not in {"document", "xhr", "fetch"}:
                return
            content_type = response.headers.get("content-type", "")
            event: dict[str, Any] = {
                "kind": "response",
                "method": request.method,
                **url_evidence(response.url),
                "status": response.status,
                "content_type": content_type.split(";", 1)[0].strip().lower(),
                "resource_type": request.resource_type,
            }
            if "json" in content_type.lower():
                try:
                    body = response.body()
                    if len(body) <= 1024 * 1024:
                        event["structure"] = summarize_json_structure(json.loads(body.decode("utf-8-sig")))
                except (Error, UnicodeDecodeError, json.JSONDecodeError):
                    pass
            append(event)

        browser_context.route("**/*", guard)
        browser_context.on("response", record_response)
        started = time.monotonic()
        last_progress = 0.0
        try:
            while not cancelled() and not finished() and time.monotonic() - started < timeout_seconds:
                pages = self._academic_pages()
                if pages:
                    pages[-1].wait_for_timeout(250)
                else:
                    time.sleep(0.25)
                elapsed = time.monotonic() - started
                if elapsed - last_progress >= 1:
                    progress("observing", {
                        "message": "manual navigation observation active",
                        "event_count": len(events),
                        "blocked_count": len(blocked_requests),
                        "pages": [url_evidence(page.url) for page in pages],
                    })
                    last_progress = elapsed
        finally:
            browser_context.unroute("**/*", guard)
            browser_context.remove_listener("response", record_response)
        return {
            "status": "complete",
            "report": {
                "browser_instances": 1,
                "browser_launches": self.browser_launches,
                "events": events,
                "blocked_requests": blocked_requests,
                "finished_by_user": finished(),
                "timed_out": time.monotonic() - started >= timeout_seconds,
            },
        }

    def reset_login(self) -> None:
        """Close the owned browser and remove authentication-bearing profile state."""
        import shutil

        from course_progress.explorer import resolve_profile_dir

        self.close()
        profile_dir = resolve_profile_dir(self.profile_root)
        if profile_dir.exists():
            shutil.rmtree(profile_dir)
        if profile_dir.exists():
            raise RuntimeError("无法清除旧教务浏览器会话；已阻止重新连接")
        (self.profile_root / "webvpn-auth-state.json").unlink(missing_ok=True)
        self._academic_page = None
        self._shell_page = None
        self._closing = False
        self._browser_closed_notified = False

    def close(self) -> None:
        self._closing = True
        if self._session is not None:
            self._session.__exit__(None, None, None)
            self._session = None
        if self._playwright is not None:
            self._playwright.stop()
            self._playwright = None

    def poll(self) -> None:
        if self._session is None or self._session.context is None:
            return
        try:
            context = self._session.context
            pages = context.pages
            if pages:
                pages[0].wait_for_timeout(1)
            if self._shell_url and (self._shell_page is None or self._page_is_closed(self._shell_page)) and pages:
                page = context.new_page()
                page.goto(self._shell_url, wait_until="domcontentloaded", timeout=30_000)
                page.bring_to_front()
                self._shell_page = page
        except Exception:
            self._browser_closed()
            raise
