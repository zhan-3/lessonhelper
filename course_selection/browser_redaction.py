"""Trace-local redaction for browser observation evidence."""

from __future__ import annotations

import hashlib
import hmac
import json
import re
import secrets
from dataclasses import dataclass, field
from typing import Any
from urllib.parse import parse_qsl, urlencode, urlsplit, urlunsplit

_DEFAULT_SENSITIVE_NAMES = {
    "authorization", "cookie", "set-cookie", "password", "passwd", "secret",
    "token", "access_token", "refresh_token", "id_token", "session", "sessionid",
    "csrf", "csrf_token", "xsrf", "ticket", "code", "account", "account_id",
    "username", "user_id", "student", "student_id", "student_number", "学号",
}
_NAME_PARTS = re.compile(r"(?:auth|cookie|credential|password|secret|session|token|csrf|xsrf|ticket|student|account)", re.IGNORECASE)


@dataclass
class TraceRedactor:
    """Apply one policy to all evidence channels for the lifetime of one trace."""

    additional_sensitive_names: tuple[str, ...] = ()
    _key: bytearray = field(default_factory=lambda: bytearray(secrets.token_bytes(32)), init=False, repr=False)
    _closed: bool = field(default=False, init=False, repr=False)

    def _is_sensitive_name(self, name: str) -> bool:
        normalized = str(name).strip().lower().replace("-", "_")
        configured = {item.strip().lower().replace("-", "_") for item in self.additional_sensitive_names}
        unprefixed = normalized.removeprefix("x_")
        return normalized in _DEFAULT_SENSITIVE_NAMES or normalized in configured or unprefixed in configured or bool(_NAME_PARTS.search(normalized))

    def _replacement(self, value: Any) -> dict[str, Any]:
        if self._closed:
            raise RuntimeError("trace redactor is closed")
        encoded = str(value).encode("utf-8", errors="replace")
        fingerprint = hmac.new(bytes(self._key), encoded, hashlib.sha256).hexdigest()[:16]
        return {"redacted": True, "type": type(value).__name__, "length": len(encoded), "fingerprint": fingerprint}

    def redact_url(self, url: str) -> dict[str, Any]:
        parsed = urlsplit(str(url or ""))
        query_names = [name for name, _ in parse_qsl(parsed.query, keep_blank_values=True)]
        safe_query = urlencode([(name, "<redacted>") for name in query_names])
        depth = len([part for part in parsed.path.split("/") if part])
        sanitized = urlunsplit((parsed.scheme, parsed.hostname or "", f"/<path:{depth}>", safe_query, ""))
        return {"url": sanitized, "query_names": query_names}

    def redact_headers(self, headers: dict[str, Any] | None) -> dict[str, Any]:
        result: dict[str, Any] = {}
        for name, value in (headers or {}).items():
            # Header names are useful contract evidence; values are not needed
            # and may hide URLs, identifiers, cookies, or vendor-specific tokens.
            result[str(name)] = self._replacement(value)
        return result

    def redact_value(self, value: Any, *, field_name: str = "", redact_scalars: bool = False) -> Any:
        if self._is_sensitive_name(field_name):
            return self._replacement(value)
        if isinstance(value, dict):
            safe_items = {}
            for key, item in value.items():
                key_text = str(key)
                safe_key = "<redacted-key>" if self._is_sensitive_name(key_text) or key_text.isdigit() else key_text
                safe_items[safe_key] = self.redact_value(item, field_name=key_text, redact_scalars=redact_scalars)
            return safe_items
        if isinstance(value, list):
            return [self.redact_value(item, redact_scalars=redact_scalars) for item in value]
        if isinstance(value, tuple):
            return [self.redact_value(item, redact_scalars=redact_scalars) for item in value]
        if isinstance(value, str) and redact_scalars:
            return self._replacement(value)
        return value

    def redact_body(self, body: str | bytes | None, content_type: str = "") -> Any:
        if body is None:
            return None
        text = body.decode("utf-8", errors="replace") if isinstance(body, bytes) else str(body)
        if "json" in content_type.lower():
            try:
                return self.redact_value(json.loads(text), redact_scalars=True)
            except json.JSONDecodeError:
                return self._replacement(text)
        if "form" in content_type.lower():
            return {name: self.redact_value(value, field_name=name) for name, value in parse_qsl(text, keep_blank_values=True)}
        return {"type": "opaque", "length": len(text.encode("utf-8")), "fingerprint": self._replacement(text)["fingerprint"]}

    def redact_console(self, values: list[Any]) -> list[Any]:
        # Console arguments have no trustworthy field provenance. Preserve only
        # structure and correlate scalar values through trace-local fingerprints.
        result = []
        for value in values:
            if isinstance(value, (dict, list, tuple)):
                result.append(self.redact_value(value, redact_scalars=True))
            else:
                result.append(self._replacement(value))
        return result

    def close(self) -> None:
        for index in range(len(self._key)):
            self._key[index] = 0
        self._key.clear()
        self._closed = True
