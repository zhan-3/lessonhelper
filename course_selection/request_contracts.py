"""Diagnostic request-contract candidates derived from sanitized deltas."""

from __future__ import annotations

import hashlib
from collections import Counter, defaultdict
from typing import Any
from urllib.parse import urlsplit


def _shape(url: str) -> str:
    value = str(url or "")
    if "<host:" in value and "<path:" in value:
        return value
    parsed = urlsplit(value)
    if not parsed.scheme or not parsed.hostname:
        return "unknown"
    depth = len([part for part in parsed.path.split("/") if part])
    host = hashlib.sha256(parsed.hostname.lower().encode()).hexdigest()[:12]
    return f"{parsed.scheme.lower()}://<host:{host}>/<path:{depth}>"


def generate_contract_candidates(delta: dict[str, Any], *, limit: int = 20) -> list[dict[str, Any]]:
    """Rank sanitized request events without claiming a fixed production contract."""
    if not isinstance(delta, dict) or not isinstance(delta.get("events"), list):
        return []
    boundary = int(delta.get("operation_boundary_sequence", 0) or 0)
    inventory_targets = delta.get("inventory_targets", ())
    origin_by_target = {
        str(item.get("target_identity")): str(
            (item.get("semantic_signature") or {}).get("origin_class", "unresolved")
        )
        for item in inventory_targets
        if isinstance(item, dict) and isinstance(item.get("semantic_signature"), dict)
    } if isinstance(inventory_targets, list) else {}
    events = delta["events"]
    if boundary:
        events = [event for event in events if isinstance(event, dict) and int(event.get("sequence", 0) or 0) >= boundary]
    grouped: dict[tuple[str, str, str], list[dict[str, Any]]] = defaultdict(list)
    responses: dict[tuple[str, str, str], list[dict[str, Any]]] = defaultdict(list)
    for event in events:
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
        initiators = Counter(str(item.get("initiator_class", "unknown")) for item in requests)
        frame_ids = sorted({str(item.get("frame_identity", "unknown")) for item in requests})
        loader_ids = sorted({str(item.get("loader_identity", "unknown")) for item in requests})
        elapsed = [float(item.get("elapsed_ms", 0)) for item in requests]
        redirect_event = max(requests, key=lambda item: int(item.get("redirect_hop_count", 0)))
        redirect_hop_count = int(redirect_event.get("redirect_hop_count", 0))
        redirect_path_shapes = list(redirect_event.get("redirect_path_shapes", []))[:11]
        score = 0
        reasons = []
        if elapsed:
            score += 1
            reasons.append("request occurred inside the bounded observation window")
        if redirect_hop_count or any(item.get("redirected_from") for item in requests):
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
        if target != "unknown" and frame_ids != ["unknown"]:
            score += 2
            reasons.append("target and frame provenance are correlated")
        if len(requests) > 2 and all(item in {"xhr", "fetch"} for item in resources):
            score -= 8
            reasons.append("high repetition may indicate background polling")
        if method == "POST":
            reasons.append("POST retained as diagnostic candidate; method alone does not prove a write")
        body_fields = sorted({name for item in requests for name in (item.get("request_body") or {}) if name not in {"fingerprint", "length", "type", "redacted"}})
        pagination = sorted(name for name in body_fields if any(marker in name.lower() for marker in ("page", "offset", "limit", "cursor")))
        response_header_names = sorted({name for item in matching_responses for name in (item.get("headers") or {})})
        schema_material = [sorted(statuses), response_header_names, sorted(resources)]
        evidence_identity = hashlib.sha256(str([delta.get("trace_id", "trace"), key]).encode()).hexdigest()[:16]
        provenance_category = "target_frame_correlated" if target != "unknown" and frame_ids != ["unknown"] else "uncorrelated"
        target_frame_relation = "correlated" if provenance_category == "target_frame_correlated" else "uncorrelated"
        primary_resource_type = min(resources, key=lambda name: (-resources[name], name)) if resources else "unknown"
        completeness = "complete" if matching_responses else "partial"
        semantic_signature = {
            "rank_inputs": {"completeness": completeness, "provenance_category": provenance_category},
            "request": {"method": method, "path_shape": path_shape, "resource_type": primary_resource_type, "target_frame_relation": target_frame_relation},
            "redirect": {"hop_count": redirect_hop_count, "method": method, "path_shape_sequence": redirect_path_shapes},
        }
        local_or_internal = origin_by_target.get(target) in {
            "loopback", "browser_internal",
        }
        eligible = (
            not local_or_internal
            and completeness == "complete"
            and primary_resource_type in {"xhr", "fetch"}
            and (len(requests) <= 2 or bool(redirect_hop_count or pagination or body_fields))
        )
        candidates.append({
            "method": method, "path_shape": path_shape, "target_identity": target,
            "count": len(requests), "repetition": "repeated" if len(requests) > 1 else "single",
            "evidence_identity": evidence_identity,
            "header_names": sorted({name for item in requests for name in (item.get("headers") or {})}),
            "dynamic_fields": body_fields,
            "response_statuses": statuses[:10], "resource_types": dict(resources),
            "response_schema_fingerprint": hashlib.sha256(str(schema_material).encode()).hexdigest()[:16] if matching_responses else "",
            "identity_indicators": {"target": target, "frames": frame_ids, "loaders": loader_ids, "initiators": dict(initiators)},
            "pagination_indicators": pagination,
            "provenance_category": provenance_category,
            "primary_resource_type": primary_resource_type,
            "target_frame_relation": target_frame_relation,
            "semantic_signature": semantic_signature,
            "causal_eligibility": (
                "local_or_internal" if local_or_internal
                else "eligible" if eligible
                else "background_or_insufficient"
            ),
            "selected": False,
            "redirect_hop_count": redirect_hop_count,
            "redirect_path_shapes": redirect_path_shapes,
            "provenance": {"first_elapsed_ms": min(elapsed) if elapsed else None, "last_elapsed_ms": max(elapsed) if elapsed else None},
            "score": score, "reasons": reasons,
            "completeness": completeness,
            "warnings": ["diagnostic candidate only; not a verified academic read contract"],
        })
    ranked = sorted(candidates, key=lambda item: (-item["score"], item["path_shape"], item["method"]))[: max(1, min(limit, 20))]
    selected = False
    for rank, candidate in enumerate(ranked, 1):
        candidate["rank"] = rank
        if not selected and candidate["causal_eligibility"] == "eligible":
            candidate["selected"] = True
            selected = True
    return ranked


def project_operation_targets(delta: dict[str, Any], candidates: list[dict[str, Any]]) -> list[dict[str, Any]]:
    """Project selected request evidence onto bounded top-level target semantics."""
    targets = delta.get("inventory_targets") if isinstance(delta, dict) else None
    if not isinstance(targets, list):
        return []
    by_identity = {str(item.get("target_identity")): item for item in targets if isinstance(item, dict) and item.get("target_identity")}
    projected = []
    seen: set[str] = set()
    for candidate in candidates:
        if not candidate.get("selected"):
            continue
        observed = by_identity.get(str(candidate.get("target_identity")))
        if observed is None:
            continue
        relevant = observed
        owner = observed
        visited: set[str] = set()
        while owner.get("kind") != "page" and owner.get("parent_identity"):
            identity = str(owner["parent_identity"])
            if identity in visited or identity not in by_identity:
                break
            visited.add(identity)
            owner = by_identity[identity]
        owner_identity = str(owner.get("target_identity", ""))
        if not owner_identity or owner_identity in seen:
            continue
        seen.add(owner_identity)
        relevant_signature = relevant.get("semantic_signature") if isinstance(relevant.get("semantic_signature"), dict) else {}
        projected.append({
            "target_identity": owner_identity,
            "candidate_evidence_identity": candidate.get("evidence_identity"),
            "navigation_state": owner.get("navigation_state"),
            "semantic_signature": {
                "kind": owner.get("kind", "unknown"),
                "relationship": owner.get("relationship", "unknown"),
                "depth": int(relevant_signature.get("depth", 0) or 0),
                "origin_class": relevant_signature.get("origin_class", "unresolved"),
            },
        })
    return projected[:4]
