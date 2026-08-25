"""Remove authentication and personal identifiers from captured portal data."""

from __future__ import annotations

import json
import re
from collections.abc import Mapping, Sequence
from typing import Any
from urllib.parse import parse_qsl, urlencode, urlsplit, urlunsplit

REDACTED = "[REDACTED]"

_SECRET_KEY_PARTS = (
    "authorization",
    "token",
    "accesstoken",
    "refreshtoken",
    "password",
    "passwd",
    "secret",
    "cookie",
    "ticket",
    "jsessionid",
    "sessionid",
)

_PERSONAL_KEYS = {
    "xh",
    "xuehao",
    "studentid",
    "studentno",
    "studentnumber",
    "userid",
    "usercode",
    "username",
    "xm",
    "realname",
    "studentname",
    "mobile",
    "phone",
    "email",
    "idcard",
    "identitynumber",
    "yhxx",
    "yhid",
}

_JWT_RE = re.compile(r"\beyJ[A-Za-z0-9_-]+\.[A-Za-z0-9_-]+\.[A-Za-z0-9_-]+\b")
_CAS_TICKET_RE = re.compile(r"\b(?:ST|TGT)-[A-Za-z0-9._~-]+", re.IGNORECASE)
_EMAIL_RE = re.compile(r"(?<![\w.+-])[\w.+-]+@[\w.-]+\.[A-Za-z]{2,}(?![\w.-])")
_PHONE_RE = re.compile(r"(?<!\d)1[3-9]\d{9}(?!\d)")
_STUDENT_NUMBER_RE = re.compile(r"(?<!\d)20\d{8}(?!\d)")


def _normalise_key(key: object) -> str:
    return re.sub(r"[^a-z0-9]", "", str(key).lower())


def is_sensitive_key(key: object) -> bool:
    normalised = _normalise_key(key)
    return normalised in _PERSONAL_KEYS or any(
        part in normalised for part in _SECRET_KEY_PARTS
    )


def sanitize_text(value: str) -> str:
    """Redact common secrets embedded in otherwise useful text."""
    value = _JWT_RE.sub(REDACTED, value)
    value = _CAS_TICKET_RE.sub(REDACTED, value)
    value = _EMAIL_RE.sub(REDACTED, value)
    value = _PHONE_RE.sub(REDACTED, value)
    return _STUDENT_NUMBER_RE.sub(REDACTED, value)


def sanitize_url(url: str) -> str:
    """Redact sensitive query parameters while preserving endpoint identity."""
    try:
        parts = urlsplit(url)
        query = [
            (key, REDACTED if is_sensitive_key(key) else sanitize_text(value))
            for key, value in parse_qsl(parts.query, keep_blank_values=True)
        ]
        return urlunsplit(
            (parts.scheme, parts.netloc, parts.path, urlencode(query), parts.fragment)
        )
    except ValueError:
        return sanitize_text(url)


def sanitize_request_body(body: str) -> str:
    """Sanitize JSON request bodies structurally, with a text fallback."""
    try:
        parsed = json.loads(body)
    except json.JSONDecodeError:
        if "=" in body:
            pairs = parse_qsl(body, keep_blank_values=True)
            if pairs:
                return urlencode(
                    [
                        (
                            key,
                            REDACTED
                            if is_sensitive_key(key)
                            else sanitize_text(value),
                        )
                        for key, value in pairs
                    ]
                )
        return sanitize_text(body)
    return json.dumps(sanitize_data(parsed), ensure_ascii=False)


def sanitize_data(value: Any, *, parent_key: object | None = None) -> Any:
    """Recursively redact secrets while retaining response shape and course data."""
    if parent_key is not None and is_sensitive_key(parent_key):
        return REDACTED

    if isinstance(value, Mapping):
        normalised_keys = {_normalise_key(key) for key in value}
        personal_context = any(
            key in _PERSONAL_KEYS
            or key.startswith("student")
            or key.startswith("user")
            for key in normalised_keys
        )
        return {
            str(key): (
                REDACTED
                if _normalise_key(key) == "name" and personal_context
                else sanitize_data(item, parent_key=key)
            )
            for key, item in value.items()
        }

    if isinstance(value, Sequence) and not isinstance(value, (str, bytes, bytearray)):
        return [sanitize_data(item) for item in value]

    if isinstance(value, str):
        return sanitize_text(value)

    return value
