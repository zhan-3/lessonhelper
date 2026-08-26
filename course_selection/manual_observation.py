"""Exact, value-free policy for guided academic-browser observation."""

from __future__ import annotations

import json
from typing import Any, Iterable, Mapping
from urllib.parse import parse_qs, urlsplit, urlunsplit


def _canonical_url(url: str) -> str:
    parsed = urlsplit(url)
    return urlunsplit((parsed.scheme.lower(), parsed.netloc.lower(), parsed.path, "", ""))


def url_evidence(url: str) -> dict[str, Any]:
    """Retain endpoint identity and parameter names, never parameter values."""
    parsed = urlsplit(url)
    fragment_route = parsed.fragment.split("?", 1)[0]
    return {
        "url": _canonical_url(url),
        "query_field_names": sorted(parse_qs(parsed.query, keep_blank_values=True).keys()),
        "fragment_route": fragment_route if fragment_route.startswith("!/") else "",
    }


def summarize_json_structure(value: Any, *, depth: int = 0) -> dict[str, Any]:
    """Describe JSON types and field names without retaining any values."""
    if depth >= 4:
        return {"type": "truncated"}
    if isinstance(value, dict):
        return {
            "type": "object",
            "fields": {
                str(key): summarize_json_structure(child, depth=depth + 1)
                for key, child in list(value.items())[:100]
            },
        }
    if isinstance(value, list):
        return {
            "type": "array",
            "item": summarize_json_structure(value[0], depth=depth + 1) if value else {"type": "unknown"},
        }
    if value is None:
        return {"type": "null"}
    if isinstance(value, bool):
        return {"type": "boolean"}
    if isinstance(value, (int, float)):
        return {"type": "number"}
    return {"type": "string"}


def _body_values(post_data: str | None, content_type: str) -> dict[str, list[str]]:
    if not post_data:
        return {}
    if "json" in content_type.lower():
        try:
            payload = json.loads(post_data)
        except (TypeError, json.JSONDecodeError):
            return {}
        if not isinstance(payload, dict):
            return {}
        return {
            str(key): [str(item) for item in value] if isinstance(value, list) else [str(value)]
            for key, value in payload.items()
        }
    return {str(key): [str(item) for item in values] for key, values in parse_qs(post_data, keep_blank_values=True).items()}


def _field_names(post_data: str | None, content_type: str) -> list[str]:
    return sorted(_body_values(post_data, content_type))


class ManualObservationPolicy:
    """Allow navigation reads and only explicitly confirmed POST signatures."""

    def __init__(self, allowed_posts: Iterable[Mapping[str, Any]]):
        self.allowed_posts = {
            (
                _canonical_url(str(item.get("url", ""))),
                tuple(sorted(str(name) for name in item.get("query_field_names", ()))),
                tuple(sorted(str(name) for name in item.get("field_names", ()))),
            ): {
                "allowed_values": {
                    str(name): {str(value) for value in values}
                    for name, values in item.get("allowed_values", {}).items()
                },
                "integer_ranges": {
                    str(name): (int(bounds[0]), int(bounds[1]))
                    for name, bounds in item.get("integer_ranges", {}).items()
                },
            }
            for item in allowed_posts
            if item.get("url")
        }

    def inspect_request(
        self,
        method: str,
        url: str,
        post_data: str | None,
        content_type: str,
        resource_type: str,
    ) -> tuple[bool, dict[str, Any]]:
        normalized_method = method.upper()
        body_values = _body_values(post_data, content_type)
        fields = sorted(body_values)
        canonical_url = _canonical_url(url)
        query_fields = tuple(sorted(parse_qs(urlsplit(url).query, keep_blank_values=True).keys()))
        allowed = normalized_method in {"GET", "HEAD", "OPTIONS"}
        if normalized_method == "POST":
            constraints = self.allowed_posts.get((canonical_url, query_fields, tuple(fields)))
            allowed = constraints is not None
            if allowed:
                allowed = all(
                    name in body_values and set(body_values[name]).issubset(values)
                    for name, values in constraints["allowed_values"].items()
                )
            if allowed:
                for name, (minimum, maximum) in constraints["integer_ranges"].items():
                    try:
                        values = [int(value) for value in body_values[name]]
                    except (KeyError, TypeError, ValueError):
                        allowed = False
                        break
                    if not values or any(value < minimum or value > maximum for value in values):
                        allowed = False
                        break
        evidence = {
            "method": normalized_method,
            **url_evidence(url),
            "field_names": fields,
            "content_type": content_type.split(";", 1)[0].strip().lower(),
            "resource_type": resource_type,
            "blocked": not allowed,
        }
        return allowed, evidence
