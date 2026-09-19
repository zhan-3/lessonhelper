"""Browser-backed lab session adapter.

``BrowserLabSession`` implements the core ``LabSession`` protocol on top of a
``LabTransport``.  It lives outside ``lab_booking`` so the booking core stays
free of browser and network imports; only the CLI and other composition roots
reach for this adapter.

Two ways to obtain a session, differing only in who owns the browser:

* :meth:`attach` borrows a tab from a browser you started with a CDP port;
* :meth:`from_profile` launches the project's own persistent Chromium, whose
  profile already carries the openlab sign-in — no manual browser, no port.
  The profile is shared with ``course_progress``, so a sign-in survives across
  runs.

The app token is a credential: it is learned in memory and never persisted.
"""

from __future__ import annotations

import time
from collections.abc import Mapping
from pathlib import Path
from typing import Any

from .lab_booking import LabSlot
from .lab_ports import DEFAULT_ORIGIN, LabTransport
from .lab_transport import (
    TRIGGER_ROUTES,
    BrowserLabTransport,
    install_token_hook,
    page_token,
)

# The centre apps are Vue SPAs and the token only appears once a protected
# route is visited, so nudge the router through each candidate until the
# installed hook captures it.
_ROUTE_PUSH_JS = """(route) => {
  const app = document.querySelector('#app') && document.querySelector('#app').__vue_app__;
  if (app) app.config.globalProperties.$router.push(route).catch(() => {});
}"""

TOKEN_CAPTURE_SECONDS = 60.0


def capture_app_token(page: Any, *, timeout_seconds: float = TOKEN_CAPTURE_SECONDS) -> str:
    """Drive the SPA until the app token appears, then return it (memory only)."""
    install_token_hook(page)
    token = page_token(page)
    deadline = time.monotonic() + timeout_seconds
    for route in TRIGGER_ROUTES:
        if token:
            break
        page.evaluate(_ROUTE_PUSH_JS, route)
        while time.monotonic() < deadline and not token:
            time.sleep(0.7)
            token = page_token(page)
    if not token:
        raise RuntimeError(
            "app token not captured; sign in to openlab and interact with the page once, then retry"
        )
    return token


def _page_for_center(context: Any, center: str, *, open_if_missing: bool) -> Any:
    """Pick the tab already on ``/<center>/``, optionally opening one."""
    page = next((item for item in context.pages if f"/{center}/" in item.url), None)
    if page is not None:
        return page
    if not open_if_missing:
        raise RuntimeError(f"no open tab for center '{center}'; open it from the gateway first")
    page = context.new_page()
    page.goto(f"{DEFAULT_ORIGIN}/{center}/booking/", wait_until="domcontentloaded", timeout=60_000)
    return page


class BrowserLabSession:
    """LabSession over a transport; the default keeps requests inside the app's page."""

    def __init__(
        self,
        page: Any,
        center: str,
        token: str,
        *,
        pause: float = 0.2,
        transport: LabTransport | None = None,
    ):
        self.page, self.center, self.token, self.pause = page, center, token, pause
        self._transport = transport or BrowserLabTransport(page, center, token, pause=pause)
        # Set by ``attach`` / ``from_profile``; declared here so the attribute is
        # typed and ``close`` needs no ``getattr`` fallback.
        self._playwright: Any = None

    @classmethod
    def attach(cls, cdp_url: str, center: str) -> BrowserLabSession:
        """Borrow an already-authenticated tab from a CDP-attached browser."""
        from playwright.sync_api import sync_playwright

        playwright = sync_playwright().start()
        try:
            browser = playwright.chromium.connect_over_cdp(cdp_url)
            if not browser.contexts:
                raise RuntimeError("CDP browser has no persistent context")
            page = _page_for_center(browser.contexts[0], center, open_if_missing=False)
            token = capture_app_token(page)
        except Exception:
            playwright.stop()
            raise
        return cls._bound(page, center, token, playwright)

    @classmethod
    def from_profile(cls, profile_root: Path | str, center: str) -> BrowserLabSession:
        """Launch the project's persistent Chromium and reuse its openlab sign-in.

        No externally started browser and no CDP port are needed.  If the centre
        page is not already open it is opened, which is also what surfaces the
        sign-in window the first time the stored session has expired.
        """
        from playwright.sync_api import sync_playwright

        from course_progress.explorer import launch_browser_context, resolve_profile_dir

        playwright = sync_playwright().start()
        try:
            context = launch_browser_context(
                playwright, "chromium", resolve_profile_dir(Path(profile_root))
            )
            page = _page_for_center(context, center, open_if_missing=True)
            token = capture_app_token(page)
        except Exception:
            playwright.stop()
            raise
        return cls._bound(page, center, token, playwright)

    @classmethod
    def _bound(cls, page: Any, center: str, token: str, playwright: Any) -> BrowserLabSession:
        session = cls(page, center, token)
        session._playwright = playwright
        return session

    def as_transport(self) -> LabTransport:
        """Expose the transport, transferring ownership of the browser session.

        The returned transport's ``close`` then releases the Playwright
        instance, so a caller that only needs reads can keep using the existing
        ``try/finally: transport.close()`` shape.
        """
        transport = self._transport
        if isinstance(transport, BrowserLabTransport):
            transport._playwright = self._playwright
            self._playwright = None
        return transport

    def close(self) -> None:
        self._transport.close()
        if self._playwright is not None:
            self._playwright.stop()

    def _call(self, path: str, form: Mapping[str, Any] | None = None) -> Mapping[str, Any]:
        return self._transport.call(path, form)

    def _result(self, path: str, form: Mapping[str, Any] | None = None) -> Any:
        payload = self._call(path, form)
        if payload.get("transport") != "ok" or payload.get("code") != 0:
            raise RuntimeError(f"{path} failed: {payload.get('code')} {payload.get('message')}")
        return payload.get("result")

    def booked(self) -> list[Mapping[str, Any]]:
        return list(self._result("view/lesson/ckkb") or [])

    def subjects(self) -> list[Mapping[str, Any]]:
        return list(self._result("view/subjects") or [])

    def schedule(self, subject_id: int) -> list[Mapping[str, Any]]:
        result = self._result("view/booking/yyxh", {"id": subject_id}) or {}
        return list(result.get("data") or [])

    def free_seat(self, subject_id: int, class_date: str, timer: int) -> Mapping[str, Any] | None:
        payload = self._call("view/booking/yyxkzw",
                             {"subjectId": subject_id, "classDate": class_date, "timer": timer})
        if payload.get("transport") != "ok" or payload.get("code") != 0:
            return None
        data = (payload.get("result") or {}).get("data") or {}
        for seat in data.get("labList") or []:
            if str(seat.get("status")) == "空闲":
                return {"seatId": seat.get("id"), "tableNo": seat.get("TableNo"),
                        "lab": (data.get("lab") or {}).get("LabsName")}
        return None

    def submit(self, slot: LabSlot) -> Mapping[str, Any]:
        return self._call("view/booking/doyyxkzw", {
            "subjectId": slot.subject_id,
            "classDate": slot.class_date,
            "timer": f"{slot.start}:00" if len(slot.start) == 5 else slot.start,
            "timercode": slot.timer,
            "id": slot.seat_id,
        })
