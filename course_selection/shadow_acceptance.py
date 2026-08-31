"""Fail-closed reporting for an authorized read-only Shadow acceptance."""

from __future__ import annotations

import json
import sys
from dataclasses import dataclass
from datetime import date, datetime, timezone
from pathlib import Path
from typing import Any

_ALLOWED_ACTIONS = ("connect", "inventory", "start", "checkpoint", "stop", "disconnect")
_ALLOWED_OPERATIONS = {"current_timetable_query", "selection_list_query", "grade_progress_query", "other_read_only_query"}
_ALLOWED_EVIDENCE_CODES = {"target", "redirect", "candidate_shape", "redaction", "completeness", "browser_survival", "devtools"}


@dataclass(frozen=True)
class ShadowWitness:
    operation: str
    authorized: bool
    mode: str
    witness_date: str
    devtools_confirmed: bool
    target_confirmed: bool
    redirects_confirmed: bool
    candidate_shape_confirmed: bool
    redaction_confirmed: bool
    read_only_attested: bool
    prohibited_actions_absent: bool
    trace_complete: bool
    disconnect_confirmed: bool
    borrowed_browser_alive: bool
    profile_session_tabs_alive: bool
    missing_evidence: tuple[str, ...]
    semantic_contradictions: tuple[str, ...]
    manual_interventions: int
    removed_manual_devtools_step: bool

    @classmethod
    def from_dict(cls, value: dict[str, Any]) -> ShadowWitness:
        allowed = set(cls.__dataclass_fields__)
        unknown = set(value) - allowed
        if unknown:
            raise ValueError("unsupported or potentially sensitive witness fields: " + ", ".join(sorted(unknown)))
        return cls(
            operation=str(value.get("operation", "")).strip(),
            authorized=value.get("authorized") is True,
            mode=str(value.get("mode", "")),
            witness_date=str(value.get("witness_date", "")),
            devtools_confirmed=value.get("devtools_confirmed") is True,
            target_confirmed=value.get("target_confirmed") is True,
            redirects_confirmed=value.get("redirects_confirmed") is True,
            candidate_shape_confirmed=value.get("candidate_shape_confirmed") is True,
            redaction_confirmed=value.get("redaction_confirmed") is True,
            read_only_attested=value.get("read_only_attested") is True,
            prohibited_actions_absent=value.get("prohibited_actions_absent") is True,
            trace_complete=value.get("trace_complete") is True,
            disconnect_confirmed=value.get("disconnect_confirmed") is True,
            borrowed_browser_alive=value.get("borrowed_browser_alive") is True,
            profile_session_tabs_alive=value.get("profile_session_tabs_alive") is True,
            missing_evidence=_evidence_codes(value.get("missing_evidence", [])),
            semantic_contradictions=_evidence_codes(value.get("semantic_contradictions", [])),
            manual_interventions=max(0, int(value.get("manual_interventions", 0))),
            removed_manual_devtools_step=value.get("removed_manual_devtools_step") is True,
        )


def _evidence_codes(values: Any) -> tuple[str, ...]:
    if not isinstance(values, list):
        raise TypeError("evidence codes must be an array")
    result = tuple(map(str, values))
    invalid = set(result) - _ALLOWED_EVIDENCE_CODES
    if invalid:
        raise ValueError("unsupported or potentially sensitive evidence codes: " + ", ".join(sorted(invalid)))
    return result


def evaluate_shadow_acceptance(witness: ShadowWitness) -> dict[str, Any]:
    blockers = []
    if witness.operation not in _ALLOWED_OPERATIONS:
        blockers.append("specific read-only operation code was not identified")
    if not witness.authorized:
        blockers.append("specific operation was not explicitly authorized")
    if witness.mode != "observe":
        blockers.append("mode must be observe")
    try:
        parsed_date = date.fromisoformat(witness.witness_date)
        if parsed_date > datetime.now(timezone.utc).date():
            blockers.append("witness date is in the future")
    except ValueError:
        blockers.append("witness date is invalid")
    confirmations = {
        "devtools": witness.devtools_confirmed,
        "target": witness.target_confirmed,
        "redirects": witness.redirects_confirmed,
        "candidate_shape": witness.candidate_shape_confirmed,
        "redaction": witness.redaction_confirmed,
        "read_only_attestation": witness.read_only_attested,
        "prohibited_actions_absent": witness.prohibited_actions_absent,
        "trace_completeness": witness.trace_complete,
        "disconnect": witness.disconnect_confirmed,
        "borrowed_browser_survival": witness.borrowed_browser_alive,
        "profile_session_tabs_survival": witness.profile_session_tabs_alive,
    }
    blockers.extend(f"{name} was not independently confirmed" for name, confirmed in confirmations.items() if not confirmed)
    blockers.extend(f"missing evidence: {item}" for item in witness.missing_evidence)
    blockers.extend(f"semantic contradiction: {item}" for item in witness.semantic_contradictions)
    status = "passed" if not blockers else "partial" if witness.authorized and witness.mode == "observe" else "failed"
    real_tier = "real-environment verified" if status == "passed" else "not_verified"
    return {
        "status": status,
        "mode": "observe",
        "allowed_actions": list(_ALLOWED_ACTIONS),
        "operation": witness.operation,
        "witness_date": witness.witness_date,
        "blockers": blockers,
        "evidence_tiers": {
            "observer_implementation": "implemented",
            "synthetic_behavior": "automated-test verified",
            "authorized_shadow_witness": real_tier,
        },
        "manual_interventions": witness.manual_interventions,
        "removed_manual_devtools_step": witness.removed_manual_devtools_step,
        "privacy_boundary": "no raw URLs, headers, bodies, HTML, screenshots, credentials, cookies, tokens, or student records",
        "release_risks": ["live academic interfaces may drift after the dated witness"] if status == "passed" else ["current university-system compatibility remains unverified"],
        "next_safe_step": "retain the dated sanitized witness report" if status == "passed" else "repeat one explicitly authorized read-only observation after resolving blockers",
    }


def _private_path(value: str) -> Path:
    path = Path(value).resolve()
    private_root = (Path.cwd() / ".private").resolve()
    if not path.is_relative_to(private_root):
        raise ValueError("shadow acceptance files must remain below .private")
    return path


def main() -> int:
    if len(sys.argv) != 3:
        print("usage: python -m course_selection.shadow_acceptance .private/WITNESS.json .private/REPORT.json")
        return 2
    witness_path = _private_path(sys.argv[1])
    output = _private_path(sys.argv[2])
    witness = ShadowWitness.from_dict(json.loads(witness_path.read_text(encoding="utf-8")))
    report = evaluate_shadow_acceptance(witness)
    output.parent.mkdir(parents=True, exist_ok=True)
    output.write_text(json.dumps(report, ensure_ascii=False, indent=2), encoding="utf-8")
    print(f"shadow acceptance: {report['status']} ({output})")
    return 0 if report["status"] == "passed" else 1


if __name__ == "__main__":
    raise SystemExit(main())
