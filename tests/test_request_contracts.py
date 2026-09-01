from course_selection.request_contracts import generate_contract_candidates


def test_candidates_rank_redirected_fetch_and_aggregate_repetition():
    events = [
        {"kind": "request", "method": "GET", "url_shape": "https://<host:a>/<path:1>", "target_identity": "target", "resource_type": "fetch", "redirected_from": "https://<host:a>/<path:1>", "redirect_hop_count": 2, "redirect_path_shapes": ["https://<host:a>/<path:1>", "https://<host:b>/<path:1>", "https://<host:a>/<path:1>"], "headers": {"X-Trace": "safe"}},
        {"kind": "request", "method": "GET", "url_shape": "https://<host:a>/<path:1>", "target_identity": "target", "resource_type": "fetch", "request_body": {"page": {"redacted": True}}},
        {"kind": "response", "method": "GET", "url_shape": "https://<host:a>/<path:1>", "target_identity": "target", "resource_type": "fetch", "status": 200},
    ]
    candidates = generate_contract_candidates({"events": events})
    assert len(candidates) == 1
    candidate = candidates[0]
    assert candidate["count"] == 2
    assert candidate["repetition"] == "repeated"
    assert candidate["score"] >= 10
    assert candidate["redirect_hop_count"] == 2
    assert len(candidate["redirect_path_shapes"]) == 3
    assert "diagnostic candidate only" in candidate["warnings"][0]


def test_causally_linked_request_outranks_similar_polling_decoy():
    events = []
    for sequence in range(4):
        events.append({
            "kind": "request", "method": "GET", "url_shape": "https://<host:a>/<path:3>",
            "target_identity": "background", "frame_identity": "frame-bg", "loader_identity": "loader-bg",
            "initiator_class": "page", "resource_type": "fetch", "elapsed_ms": sequence + 1,
        })
    events.extend([
        {"kind": "request", "method": "POST", "url_shape": "https://<host:a>/<path:2>",
         "target_identity": "business", "frame_identity": "frame-business", "loader_identity": "loader-nav",
         "initiator_class": "frame", "resource_type": "xhr", "elapsed_ms": 10,
         "redirected_from": "https://<host:a>/<path:1>", "request_body": {"page": {"redacted": True}}},
        {"kind": "response", "method": "POST", "url_shape": "https://<host:a>/<path:2>",
         "target_identity": "business", "resource_type": "xhr", "status": 200, "headers": {"content-type": {"redacted": True}}},
    ])
    candidates = generate_contract_candidates({"trace_id": "trace-1", "events": events})
    assert candidates[0]["target_identity"] == "business"
    assert candidates[0]["pagination_indicators"] == ["page"]
    assert candidates[0]["evidence_identity"]
    assert candidates[0]["identity_indicators"]["loaders"] == ["loader-nav"]
    assert candidates[0]["provenance_category"] == "target_frame_correlated"


def test_post_reads_remain_diagnostic_and_unknown_response_is_partial():
    candidates = generate_contract_candidates({"events": [{
        "kind": "request", "method": "POST", "url_shape": "https://<host:a>/<path:2>",
        "target_identity": "target", "resource_type": "xhr",
    }]})
    assert candidates[0]["completeness"] == "partial"
    assert any("POST" in reason for reason in candidates[0]["reasons"])
