from course_selection.request_contracts import generate_contract_candidates


def test_candidates_rank_redirected_fetch_and_aggregate_repetition():
    events = [
        {"kind": "request", "method": "GET", "url_shape": "https://<host:a>/<path:1>", "target_identity": "target", "resource_type": "fetch", "redirected_from": "https://<host:a>/<path:1>", "headers": {"X-Trace": "safe"}},
        {"kind": "request", "method": "GET", "url_shape": "https://<host:a>/<path:1>", "target_identity": "target", "resource_type": "fetch", "request_body": {"page": {"redacted": True}}},
        {"kind": "response", "method": "GET", "url_shape": "https://<host:a>/<path:1>", "target_identity": "target", "resource_type": "fetch", "status": 200},
    ]
    candidates = generate_contract_candidates({"events": events})
    assert len(candidates) == 1
    candidate = candidates[0]
    assert candidate["count"] == 2
    assert candidate["repetition"] == "repeated"
    assert candidate["score"] >= 10
    assert "diagnostic candidate only" in candidate["warnings"][0]


def test_post_reads_remain_diagnostic_and_unknown_response_is_partial():
    candidates = generate_contract_candidates({"events": [{
        "kind": "request", "method": "POST", "url_shape": "https://<host:a>/<path:2>",
        "target_identity": "target", "resource_type": "xhr",
    }]})
    assert candidates[0]["completeness"] == "partial"
    assert any("POST" in reason for reason in candidates[0]["reasons"])
