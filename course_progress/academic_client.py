"""Authenticated, browser-backed transport for fixed academic-system contracts.

The browser owns authentication only. Readers provide versioned paths and form
parameters; this client performs no menu discovery or control clicking.
"""

from __future__ import annotations

from dataclasses import dataclass
from typing import Any, Callable, Mapping
from urllib.parse import urljoin, urlsplit

from course_progress.explorer import _is_login_url


class AcademicClientError(RuntimeError):
    """Base error for a fixed academic request."""


class AcademicAuthenticationRequired(AcademicClientError):
    """The fixed request resolved to a login page."""


class AcademicContractError(AcademicClientError):
    """A fixed endpoint no longer matches its verified contract."""


def resolve_academic_url(current_url: str, endpoint: str) -> str:
    """Resolve a direct or WebVPN-rewritten academic endpoint."""
    import re

    # An already-absolute endpoint (e.g. the verified selection entry URL) is
    # used verbatim; only relative paths are joined onto the current prefix.
    endpoint_parts = urlsplit(endpoint)
    if endpoint_parts.scheme and endpoint_parts.netloc:
        return endpoint
    parts = urlsplit(current_url)
    match = re.match(r"^(/https?/[0-9a-fA-F]+)(?:/|$)", parts.path)
    if match:
        return f"{parts.scheme}://{parts.netloc}{match.group(1)}/{endpoint.lstrip('/')}"
    return urljoin(current_url, endpoint)


_FETCH_FORM = """
async ({url, parameters, timeoutMs}) => {
  const controller = new AbortController();
  const timer = setTimeout(() => controller.abort(), timeoutMs);
  try {
    const body = new URLSearchParams(parameters).toString();
    const response = await fetch(url, {
      method: 'POST', credentials: 'same-origin', redirect: 'follow',
      headers: {'Content-Type': 'application/x-www-form-urlencoded'},
      body, signal: controller.signal,
    });
    return {status: response.status, url: response.url,
            requestBody: body, body: await response.text()};
  } finally { clearTimeout(timer); }
}
"""

_FETCH_PAGE_FORM = """
async ({url, overrides, remove, timeoutMs}) => {
  const form = document.querySelector('form#queryform, form[name="queryform"]');
  if (!form) throw new Error('verified query form is missing');
  const parameters = new URLSearchParams(new FormData(form));
  for (const [name, value] of Object.entries(overrides)) parameters.set(name, value);
  for (const name of remove) parameters.delete(name);
  const controller = new AbortController();
  const timer = setTimeout(() => controller.abort(), timeoutMs);
  try {
    const response = await fetch(url, {
      method: 'POST', credentials: 'same-origin', redirect: 'follow',
      headers: {'Content-Type': 'application/x-www-form-urlencoded'},
      body: parameters.toString(), signal: controller.signal,
    });
    return {status: response.status, url: response.url,
            requestBody: parameters.toString(), body: await response.text()};
  } finally { clearTimeout(timer); }
}
"""


@dataclass(frozen=True)
class AcademicResponse:
    status: int
    url: str
    body: str
    request_body: str = ""


class AuthenticatedAcademicClient:
    """Send bounded fixed requests through one authenticated browser page."""

    def __init__(
        self,
        page: Any,
        *,
        authenticate: Callable[[str, Any], Any],
        timeout_seconds: int = 15,
        cancelled: Callable[[], bool] = lambda: False,
    ) -> None:
        self.page = page
        self.authenticate = authenticate
        self.timeout_seconds = max(1, int(timeout_seconds))
        self.cancelled = cancelled
        self.trace_requests: list[dict[str, Any]] = []

    def endpoint(self, path: str) -> str:
        return resolve_academic_url(str(getattr(self.page, "url", "")), path)

    def get(self, path: str) -> AcademicResponse:
        if self.cancelled():
            raise AcademicClientError("request cancelled")
        url = self.endpoint(path)
        self.trace_requests.append({"method": "GET", "url": url, "resource_type": "document"})
        self.page = self.authenticate(url, self.page)
        final_url = str(getattr(self.page, "url", url))
        if _is_login_url(final_url):
            raise AcademicAuthenticationRequired("academic authentication is required")
        content = getattr(self.page, "content", None)
        body = str(content()) if callable(content) else ""
        return AcademicResponse(200, final_url, body)

    def post_form(
        self,
        path: str,
        parameters: Mapping[str, Any],
        *,
        retry_read_once: bool = False,
    ) -> AcademicResponse:
        attempts = 2 if retry_read_once else 1
        last_error: Exception | None = None
        for attempt in range(attempts):
            if self.cancelled():
                raise AcademicClientError("request cancelled")
            url = self.endpoint(path)
            try:
                result = self.page.evaluate(
                    _FETCH_FORM,
                    {"url": url, "parameters": {str(k): str(v) for k, v in parameters.items()},
                     "timeoutMs": self.timeout_seconds * 1000},
                )
                request_body = str(result.get("requestBody") or "")
                self.trace_requests.append({
                    "method": "POST", "url": str(result.get("url") or url),
                    "resource_type": "fetch", "post_data": request_body,
                })
                response = AcademicResponse(
                    int(result.get("status") or 0), str(result.get("url") or url),
                    str(result.get("body") or ""), request_body,
                )
                if _is_login_url(response.url) or "authserver/login" in response.body.lower():
                    raise AcademicAuthenticationRequired("academic authentication is required")
                return response
            except AcademicAuthenticationRequired:
                raise
            except Exception as error:  # Playwright errors vary by runtime version.
                last_error = error
                if attempt + 1 < attempts:
                    self.get(path)
        raise AcademicClientError(str(last_error or "academic request failed"))

    def post_page_form(
        self,
        path: str,
        *,
        overrides: Mapping[str, Any] | None = None,
        remove: tuple[str, ...] = (),
        retry_read_once: bool = False,
    ) -> AcademicResponse:
        """Submit the verified fixed page form without navigating or clicking."""
        attempts = 2 if retry_read_once else 1
        last_error: Exception | None = None
        for attempt in range(attempts):
            if self.cancelled():
                raise AcademicClientError("request cancelled")
            url = self.endpoint(path)
            try:
                result = self.page.evaluate(
                    _FETCH_PAGE_FORM,
                    {"url": url, "overrides": {str(k): str(v) for k, v in (overrides or {}).items()},
                     "remove": list(remove), "timeoutMs": self.timeout_seconds * 1000},
                )
                request_body = str(result.get("requestBody") or "")
                self.trace_requests.append({"method": "POST", "url": str(result.get("url") or url),
                                            "resource_type": "fetch", "post_data": request_body})
                return AcademicResponse(int(result.get("status") or 0), str(result.get("url") or url),
                                        str(result.get("body") or ""), request_body)
            except Exception as error:
                last_error = error
                if attempt + 1 < attempts:
                    self.get(path)
        raise AcademicClientError(str(last_error or "academic request failed"))
