"""Strict, immutable discovery of official selection-arrangement notices."""

from __future__ import annotations

import hashlib
from dataclasses import asdict
from difflib import unified_diff
from urllib.parse import urlparse

from .notice import parse_notice

ARRANGEMENT_MARKERS = ("选课时间安排", "各类课程选课时间", "课程选课安排")


def candidate_from_text(source_url: str, text: str, *, official_hosts: tuple[str, ...]) -> dict:
    host = (urlparse(source_url).hostname or "").lower()
    if host not in {item.lower() for item in official_hosts}:
        raise ValueError("notice source is not an approved official host")
    notice = parse_notice(text, source_url=source_url, source_kind="official")
    if not any(marker in notice.title for marker in ARRANGEMENT_MARKERS):
        raise ValueError("article is not a course-selection time arrangement")
    payload = asdict(notice)
    payload["content_hash"] = hashlib.sha256(text.encode("utf-8")).hexdigest()
    payload["version_id"] = hashlib.sha256(f"{source_url}\n{text}".encode("utf-8")).hexdigest()
    payload["status"] = "candidate"
    payload["query_eligible"] = not notice.missing_fields and bool(notice.windows)
    return payload


def notice_diff(previous: dict, candidate: dict) -> str:
    return "\n".join(unified_diff(previous.get("source_text", "").splitlines(), candidate.get("source_text", "").splitlines(), fromfile="confirmed", tofile="candidate", lineterm=""))
