from datetime import datetime, timezone

import pytest

from course_selection.shadow_acceptance import ShadowWitness, evaluate_shadow_acceptance


def witness(**overrides):
    value = {
        "operation": "current_timetable_query",
        "authorized": True,
        "mode": "observe",
        "witness_date": datetime.now(timezone.utc).date().isoformat(),
        "devtools_confirmed": True,
        "target_confirmed": True,
        "redirects_confirmed": True,
        "candidate_shape_confirmed": True,
        "redaction_confirmed": True,
        "read_only_attested": True,
        "prohibited_actions_absent": True,
        "trace_complete": True,
        "disconnect_confirmed": True,
        "borrowed_browser_alive": True,
        "profile_session_tabs_alive": True,
        "missing_evidence": [],
        "semantic_contradictions": [],
        "manual_interventions": 1,
        "removed_manual_devtools_step": True,
    }
    value.update(overrides)
    return ShadowWitness.from_dict(value)


def test_pass_requires_authorized_dated_independent_live_witness():
    report = evaluate_shadow_acceptance(witness())
    assert report["status"] == "passed"
    assert report["evidence_tiers"]["authorized_shadow_witness"] == "real-environment verified"
    assert report["allowed_actions"] == ["connect", "inventory", "start", "checkpoint", "stop", "disconnect"]


def test_missing_evidence_and_browser_disruption_fail_closed():
    report = evaluate_shadow_acceptance(witness(
        borrowed_browser_alive=False,
        missing_evidence=["target"],
        semantic_contradictions=["candidate_shape"],
    ))
    assert report["status"] == "partial"
    assert report["evidence_tiers"]["authorized_shadow_witness"] == "not_verified"
    assert report["release_risks"]


def test_rejects_unapproved_or_sensitive_witness_fields():
    report = evaluate_shadow_acceptance(witness(authorized=False))
    assert report["status"] == "failed"
    with pytest.raises(ValueError, match="potentially sensitive"):
        ShadowWitness.from_dict({"operation": "other_read_only_query", "cookie": "secret"})
    with pytest.raises(ValueError, match="sensitive evidence"):
        witness(missing_evidence=["https://secret.example/student/2025000000"])
