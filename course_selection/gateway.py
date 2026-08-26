"""Framework-independent boundary for read-only university observations."""

from __future__ import annotations

from typing import Any, Callable, Protocol

Progress = Callable[[str, dict[str, Any]], None]
Cancelled = Callable[[], bool]


class AcademicGateway(Protocol):
    def connect(self, progress: Progress, cancelled: Cancelled) -> None: ...
    def refresh_selection(self, context: dict[str, Any], progress: Progress, cancelled: Cancelled) -> dict[str, Any]: ...
    def refresh_timetable(self, context: dict[str, Any], progress: Progress, cancelled: Cancelled) -> dict[str, Any]: ...
    def close(self) -> None: ...


class UnconfirmedAcademicGateway:
    """Safe default until live read contracts have been independently verified."""

    def connect(self, progress: Progress, cancelled: Cancelled) -> None:
        progress("connecting", {"message": "academic gateway requires explicit configuration"})

    def refresh_selection(self, context: dict[str, Any], progress: Progress, cancelled: Cancelled) -> dict[str, Any]:
        return {"status": "interface_unconfirmed", "sections": []}

    def refresh_timetable(self, context: dict[str, Any], progress: Progress, cancelled: Cancelled) -> dict[str, Any]:
        # No timetable request contract has been confirmed.  Returning an
        # explicit blocked result is important: an empty list would mean
        # "free" to downstream conflict checks.  This method intentionally
        # performs no navigation, guessing, or write request.
        progress("interface_unconfirmed", {"target": "timetable", "message": "timetable read interface is not confirmed"})
        return {
            "status": "interface_unconfirmed",
            "reason": "timetable_read_contract_unconfirmed",
            "source_kind": "academic-unconfirmed",
            "entries": [],
        }

    def close(self) -> None:
        return None


class PlaywrightAcademicGateway(UnconfirmedAcademicGateway):
    """Long-lived bundled-Chromium session, owned by the observation worker."""

    def __init__(self, profile_root=".private/course-progress", workspace_root=".private/academic-selection", portal_url: str | None = None):
        from pathlib import Path
        from course_progress.explorer import DEFAULT_PORTAL_URL
        self.profile_root = Path(profile_root)
        self.workspace_root = Path(workspace_root)
        self.portal_url = portal_url or DEFAULT_PORTAL_URL
        self._playwright = None
        self._session = None

    def connect(self, progress: Progress, cancelled: Cancelled) -> None:
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
            persistent=False,
        )
        self._session.__enter__()
        try:
            # Authentication is an intentional, user-visible pause.  Publish
            # this before entering the blocking browser wait so task observers
            # can render the state while the user operates the browser.
            progress("waiting_for_authentication", {"message": "waiting for authentication in the bundled browser"})
            self._session.open_authenticated(self.portal_url, timeout_seconds=600)
        except TimeoutError:
            progress("waiting_for_authentication", {"message": "authentication timed out; browser remains available"})
            raise

    def refresh_selection(self, context: dict[str, Any], progress: Progress, cancelled: Cancelled) -> dict[str, Any]:
        if self._session is None or self._session.context is None:
            raise RuntimeError("academic session is disconnected")
        from types import SimpleNamespace
        from .discovery import InterfaceDiscovery, TARGET_SELECTION
        categories = tuple(context.get("allowed_categories", ()))
        if not categories or not context.get("semester_label"):
            return {"status": "no_matching_round", "sections": []}
        windows = {
            code: tuple(SimpleNamespace(**item) for item in items)
            for code, items in context.get("allowed_windows", {}).items()
        }
        navigator = InterfaceDiscovery(
            target=TARGET_SELECTION, output_root=self.workspace_root,
            allowed_selection_categories=categories,
            allowed_selection_windows=windows,
            notice_semester=context["semester_label"],
        )
        self._session.context.on("response", navigator._handle_response)
        self._session.context.route("**/*", navigator._guard_route)
        progress("reading", {"category": categories[0], "page": 1})
        report = navigator.run(self._session.context, max_clicks=8, wait_seconds=30)
        if not report.selection_query_path:
            return {"status": "interface_unconfirmed", "sections": []}
        import json
        payload = json.loads(report.selection_query_path.read_text(encoding="utf-8"))
        queries = payload.get("queries", [])
        complete = bool(queries) and all(item.get("complete") for item in queries)
        sections = [section for query in queries for section in query.get("sections", [])]
        return {"status": "complete" if complete else "incomplete", "sections": sections, "queries": queries}

    def close(self) -> None:
        if self._session is not None:
            self._session.__exit__(None, None, None)
            self._session = None
        if self._playwright is not None:
            self._playwright.stop()
            self._playwright = None
