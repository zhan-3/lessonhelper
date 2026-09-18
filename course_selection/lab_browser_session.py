"""Browser-backed lab session adapter.

``BrowserLabSession`` implements the core ``LabSession`` protocol on top of a
``LabTransport``.  It lives outside ``lab_booking`` so the booking core stays
free of browser and network imports; only the CLI and other composition roots
reach for this adapter.

The app token is a credential: it is learned in memory and never persisted.
"""

from __future__ import annotations

import time
from collections.abc import Mapping
from typing import Any

from .lab_booking import LabSlot
from .lab_ports import LabTransport
from .lab_transport import (
    TRIGGER_ROUTES,
    BrowserLabTransport,
    install_token_hook,
    page_token,
)


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

    @classmethod
    def attach(cls, cdp_url: str, center: str) -> BrowserLabSession:
        """Borrow the already-authenticated tab and learn its app token in memory."""
        from playwright.sync_api import sync_playwright

        playwright = sync_playwright().start()
        browser = playwright.chromium.connect_over_cdp(cdp_url)
        page = next((item for item in browser.contexts[0].pages if f"/{center}/" in item.url), None)
        if page is None:
            raise RuntimeError(f"no open tab for center '{center}'; open it from the gateway first")
        install_token_hook(page)
        token = page_token(page)
        deadline = time.monotonic() + 60
        for route in TRIGGER_ROUTES:
            if token:
                break
            page.evaluate(
                """(route) => {
                  const app = document.querySelector('#app') && document.querySelector('#app').__vue_app__;
                  if (app) app.config.globalProperties.$router.push(route).catch(() => {});
                }""",
                route,
            )
            while time.monotonic() < deadline and not token:
                time.sleep(0.7)
                token = page_token(page)
        if not token:
            raise RuntimeError("app token not captured; interact with the app once, then retry")
        session = cls(page, center, token)
        session._playwright = playwright
        return session

    def close(self) -> None:
        self._transport.close()
        playwright = getattr(self, "_playwright", None)
        if playwright is not None:
            playwright.stop()

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
