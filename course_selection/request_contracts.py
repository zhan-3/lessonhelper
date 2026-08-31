"""Diagnostic request-contract candidates derived from sanitized deltas."""

from __future__ import annotations

import hashlib
from collections import Counter, defaultdict
from typing import Any
from urllib.parse import urlsplit


def _shape(url: str) -> str:
    parsed = urlsplit(str(url or ""))
    if not parsed.scheme or not parsed.hostname:
        return "unknown"
    depth = len([part for part in parsed.path.split("/") if part])
    host = hashlib.sha256(parsed.hostname.lower().encode()).hexdigest()[:12]
    return f"{parsed.scheme.lower()}://<host:{host}>/<path:{depth}>"


def generate_contract_candidates(delta: dict[str, Any], *, limit: int = 20) -> list[dict[str, Any]]:
    """Rank sanitized request events without claiming a fixed production contract."""
    if not isinstance(delta, dict) or not isinstance(delta.get("events"), list):
        return []
    grouped: dict[tuple[str, str, str], list[dict[str, Any]]] = defaultdict(list)
    responses: dict[tuple[str, str, str], list[dict[str, Any]]] = defaultdict(list)
    for event in delta["events"]:
        if not isinstance(event, dict):
            continue
        key = (str(event.get("method", "GET")).upper(), _shape(event.get("url_shape", "")), str(event.get("target_identity", "unknown")))
        if event.get("kind") == "request":
            grouped[key].append(event)
        elif event.get("kind") == "response":
            responses[key].append(event)
    candidates = []
    for key, requests in grouped.items():
        method, path_shape, target = key
        matching_responses = responses.get(key, [])
        statuses = [item.get("status") for item in matching_responses if item.get("status") is not None]
        resources = Counter(str(item.get("resource_type", "unknown")) for item in requests)
        score = 0
        reasons = []
        if any(item.get("redirected_from") for item in requests):
            score += 5
            reasons.append("redirect chain links this request to the observed navigation")
        if any(item.get("request_body") is not None for item in requests):
            score += 2
            reasons.append("request carries a structured or opaque body")
        if any(item in {"xhr", "fetch"} for item in resources):
            score += 3
            reasons.append("XHR/fetch resource is more likely to be a business read")
        if any(status in {200, 201} for status in statuses):
            score += 2
            reasons.append("response completed successfully")
        if method == "POST":
            reasons.append("POST retained as diagnostic candidate; method alone does not prove a write")
        candidates.append({
            "method": method, "path_shape": path_shape, "target_identity": target,
            "count": len(requests), "repetition": "repeated" if len(requests) > 1 else "single",
            "header_names": sorted({name for item in requests for name in (item.get("headers") or {})}),
            "dynamic_fields": sorted({name for item in requests for name in (item.get("request_body") or {}) if name != "fingerprint"}),
            "response_statuses": statuses[:10], "resource_types": dict(resources),
            "response_schema_fingerprint": hashlib.sha256(str(sorted(statuses)).encode()).hexdigest()[:16] if statuses else "",
            "identity_indicators": ["target_identity", "frame_identity", "initiator_class"],
            "pagination_indicators": [], "score": score, "reasons": reasons,
            "completeness": "complete" if matching_responses else "partial",
            "warnings": ["diagnostic candidate only; not a verified academic read contract"],
        })
    return sorted(candidates, key=lambda item: (-item["score"], item["path_shape"], item["method"]))[: max(1, min(limit, 20))]
