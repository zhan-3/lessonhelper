"""Transport seam for openlab teaching-center apps.

Every center app is reached through one method, ``call(path, form)``, so the
booking and contract layers never care how the request travels:

* :class:`HttpLabTransport` sends a plain HTTP POST with the
  ``vctchauthorization`` header.  It needs no cookie and no browser, which is
  what the campus network allows and what keeps the library testable.
* :class:`BrowserLabTransport` runs the request as the app's own ``fetch``
  inside its page, for networks where only the browser can reach the host.
* :class:`RoutedLabTransport` prefers the primary backend and falls back only
  when the primary fails at the transport level; a business failure is never
  replayed against the other backend.

The app token is a credential: it is kept in memory for the life of the
process and never written to disk.
"""

from __future__ import annotations

import json
import time
import urllib.parse
import urllib.request
from collections.abc import Mapping, Sequence
from typing import Any

# The abstract seam now lives with the core; re-exported here so the existing
# ``from .lab_transport import ...`` call sites keep working unchanged.
from .lab_ports import CONTENT_TYPE, DEFAULT_ORIGIN, TOKEN_HEADER, Fetch, LabTransport


def _urlopen_fetch(url: str, data: bytes, headers: Mapping[str, str], timeout: float) -> tuple[int, str]:
    request = urllib.request.Request(url, data=data, method="POST", headers=dict(headers))
    with urllib.request.urlopen(request, timeout=timeout) as response:
        return response.status, response.read().decode("utf-8", "replace")


class HttpLabTransport:
    """Plain-HTTP backend: header authentication only, no cookie jar."""

    def __init__(
        self,
        center: str,
        token: str,
        *,
        origin: str = DEFAULT_ORIGIN,
        fetch: Fetch | None = None,
        timeout: float = 15.0,
        pause: float = 0.0,
    ):
        self.center = center
        self.token = token
        self.origin = origin.rstrip("/")
        self.timeout = timeout
        self.pause = pause
        self._fetch = fetch or _urlopen_fetch

    def close(self) -> None:
        """Nothing to release; present so both backends share one interface."""

    def call(self, path: str, form: Mapping[str, Any] | None = None) -> dict[str, Any]:
        body = urllib.parse.urlencode({k: v for k, v in (form or {}).items() if v is not None}).encode()
        headers = {TOKEN_HEADER: self.token, "Content-Type": CONTENT_TYPE}
        url = f"{self.origin}/{self.center}/StuApi/{path}"
        try:
            status, text = self._fetch(url, body, headers, self.timeout)
        except Exception as error:  # transport failures are typed, never raised
            return {"transport": "error", "detail": f"{type(error).__name__}: {error}"[:160]}
        finally:
            if self.pause:
                time.sleep(self.pause)
        try:
            return {"transport": "ok", "status": status, **json.loads(text)}
        except ValueError:
            return {"transport": "non_json", "detail": str(status)}


TOKEN_HOOK = """
() => {
  if (window.__vctchHook) return 'already';
  window.__vctchHook = true;
  window.__vctch = '';
  const grab = (list) => {
    for (const [name, value] of list) {
      if (String(name).toLowerCase() === 'vctchauthorization' && value) window.__vctch = value;
    }
  };
  const open = XMLHttpRequest.prototype.open;
  XMLHttpRequest.prototype.open = function (...args) { this.__h = []; return open.apply(this, args); };
  const setHeader = XMLHttpRequest.prototype.setRequestHeader;
  XMLHttpRequest.prototype.setRequestHeader = function (name, value) {
    if (this.__h) this.__h.push([name, value]);
    return setHeader.call(this, name, value);
  };
  const send = XMLHttpRequest.prototype.send;
  XMLHttpRequest.prototype.send = function (...args) { grab(this.__h || []); return send.apply(this, args); };
  const fetchOrig = window.fetch;
  window.fetch = function (input, init) {
    const headers = init && init.headers;
    if (headers) {
      if (headers instanceof Headers) grab([...headers.entries()]);
      else if (Array.isArray(headers)) grab(headers);
      else grab(Object.entries(headers));
    }
    return fetchOrig.apply(this, arguments);
  };
  return 'installed';
}
"""

# Routes whose components fetch data, used only to make the app issue a request
# so the hook can learn the token.  Switching the app's own route submits nothing.
TRIGGER_ROUTES: Sequence[str] = ("/lesson", "/booking", "/lab", "/result")


def install_token_hook(page: Any) -> str:
    """Install the read-only header observer in the app's page."""
    return str(page.evaluate(TOKEN_HOOK))


def page_token(page: Any) -> str:
    return str(page.evaluate("() => window.__vctch || ''"))


def acquire_token(
    cdp_url: str,
    center: str,
    *,
    trigger_routes: Sequence[str] = TRIGGER_ROUTES,
    timeout: float = 60.0,
) -> tuple[str, str]:
    """Borrow the authenticated tab and learn its app token from its own requests.

    Returns ``(token, page_url)``.  The token stays in memory; the browser is
    only detached, never closed.
    """
    from playwright.sync_api import sync_playwright

    playwright = sync_playwright().start()
    try:
        browser = playwright.chromium.connect_over_cdp(cdp_url)
        page = next((item for item in browser.contexts[0].pages if f"/{center}/" in item.url), None)
        if page is None:
            raise RuntimeError(f"no open tab for center '{center}'; open it from the gateway first")
        install_token_hook(page)
        token = page_token(page)
        deadline = time.monotonic() + timeout
        for route in trigger_routes:
            if token:
                break
            page.evaluate(
                """(route) => {
                  const app = document.querySelector('#app') && document.querySelector('#app').__vue_app__;
                  if (app) app.config.globalProperties.$router.push(route).catch(() => {});
                }""",
                route,
            )
            while time.monotonic() < deadline:
                time.sleep(0.7)
                token = page_token(page)
                if token:
                    break
        if not token:
            raise RuntimeError("app token not captured; keep the center app open and try again")
        return token, page.url
    finally:
        playwright.stop()


class BrowserLabTransport:
    """Backend that issues the request inside the app's own page."""

    def __init__(self, page: Any, center: str, token: str, *, pause: float = 0.2):
        self.page, self.center, self.token, self.pause = page, center, token, pause

    def close(self) -> None:
        """Nothing to release; the borrowed page belongs to the user."""

    def call(self, path: str, form: Mapping[str, Any] | None = None) -> dict[str, Any]:
        payload = self.page.evaluate(
            """async ([path, form, token]) => {
              try {
                const response = await fetch(path, {
                  method: 'POST', credentials: 'include',
                  headers: {'Content-Type': 'application/x-www-form-urlencoded;charset=UTF-8',
                            vctchauthorization: token},
                  body: new URLSearchParams(form).toString(),
                });
                const text = await response.text();
                try { return {transport: 'ok', status: response.status, ...JSON.parse(text)}; }
                catch (error) { return {transport: 'non_json', detail: String(response.status)}; }
              } catch (error) {
                return {transport: 'error', detail: String(error).slice(0, 160)};
              }
            }""",
            [f"/{self.center}/StuApi/{path}", dict(form or {}), self.token],
        )
        time.sleep(self.pause)
        return payload


def is_transport_failure(payload: Mapping[str, Any]) -> bool:
    return payload.get("transport") == "error"


class RoutedLabTransport:
    """Prefer ``primary``; use ``fallback`` only when the network path fails."""

    def __init__(self, primary: LabTransport, fallback: LabTransport):
        self.primary, self.fallback = primary, fallback
        self.center = primary.center
        self.last_used = "primary"

    def close(self) -> None:
        self.primary.close()
        self.fallback.close()

    def call(self, path: str, form: Mapping[str, Any] | None = None) -> dict[str, Any]:
        payload = self.primary.call(path, form)
        if not is_transport_failure(payload):
            self.last_used = "primary"
            return payload
        self.last_used = "fallback"
        return self.fallback.call(path, form)


def http_lab_transport(cdp_url: str, center: str, **options: Any) -> HttpLabTransport:
    """Acquire the app token from the live tab, then talk to the app over plain HTTP."""
    token, _url = acquire_token(cdp_url, center)
    return HttpLabTransport(center, token, **options)
