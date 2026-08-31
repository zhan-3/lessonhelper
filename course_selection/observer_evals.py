"""Local-only value gates for the academic browser observer."""

from __future__ import annotations

import json
import statistics
import sys
from dataclasses import dataclass
from pathlib import Path
from typing import Any


@dataclass(frozen=True)
class RunMetrics:
    scenario_id: str
    variant: str
    tokens: int
    elapsed_seconds: float
    tool_calls: int
    interventions: int
    status: str
    targets: frozenset[str]
    requests: frozenset[str]
    redirects: frozenset[str]
    candidates_correct: int
    candidates_total: int
    browser_alive: bool
    serialized_output: str

    @classmethod
    def from_dict(cls, value: dict[str, Any]) -> RunMetrics:
        return cls(
            scenario_id=str(value["scenario_id"]), variant=str(value["variant"]),
            tokens=max(0, int(value.get("tokens", 0))),
            elapsed_seconds=max(0.0, float(value.get("elapsed_seconds", 0))),
            tool_calls=max(0, int(value.get("tool_calls", 0))),
            interventions=max(0, int(value.get("interventions", 0))),
            status=str(value.get("status", "failed")),
            targets=frozenset(map(str, value.get("targets", []))),
            requests=frozenset(map(str, value.get("requests", []))),
            redirects=frozenset(map(str, value.get("redirects", []))),
            candidates_correct=max(0, int(value.get("candidates_correct", 0))),
            candidates_total=max(0, int(value.get("candidates_total", 0))),
            browser_alive=bool(value.get("browser_alive", False)),
            serialized_output=str(value.get("serialized_output", "")),
        )


def hidden_variant(scenario: dict[str, Any], seed: int) -> dict[str, Any]:
    """Vary irrelevant fixture details while retaining declared expectations."""
    variants = {
        "frame_depth": 1 + seed % 4,
        "origin_index": seed % 3,
        "name": f"fixture-{(seed * 17) % 997}",
        "delay_ms": (seed * 37) % 400,
        "target_order": "reverse" if seed % 2 else "forward",
        "request_method": "POST" if seed % 3 == 0 else "GET",
        "field_name": f"field_{(seed * 11) % 101}",
    }
    return {"scenario_id": scenario["id"], "failure_class": scenario["failure_class"], "variant": variants, "expected": {key: scenario.get(key, []) for key in ("expected_targets", "expected_requests", "expected_redirects")}}


def evaluate(manifest: dict[str, Any], runs: list[RunMetrics]) -> dict[str, Any]:
    scenarios = {item["id"]: item for item in manifest.get("scenarios", [])}
    baits = tuple(map(str, manifest.get("sensitive_baits", [])))
    findings = []
    accepted = 0
    recalled = {"targets": 0, "requests": 0, "redirects": 0}
    expected_counts = {"targets": 0, "requests": 0, "redirects": 0}
    for run in runs:
        expected = scenarios.get(run.scenario_id)
        if expected is None:
            findings.append(f"unknown scenario: {run.scenario_id}")
            continue
        missing = {
            "targets": set(expected.get("expected_targets", [])) - run.targets,
            "requests": set(expected.get("expected_requests", [])) - run.requests,
            "redirects": set(expected.get("expected_redirects", [])) - run.redirects,
        }
        expected_status = str(expected.get("expected_status", "complete"))
        if run.status != expected_status:
            findings.append(f"unexpected status: {run.variant}/{run.scenario_id} expected {expected_status}, got {run.status}")
        if run.status == "complete":
            for kind, values in missing.items():
                total = len(expected.get(f"expected_{kind}", []))
                expected_counts[kind] += total
                recalled[kind] += total - len(values)
            if any(missing.values()):
                findings.append(f"false complete: {run.variant}/{run.scenario_id}")
            else:
                accepted += 1
        if not run.browser_alive:
            findings.append(f"browser disruption: {run.variant}/{run.scenario_id}")
        if any(bait and bait in run.serialized_output for bait in baits):
            findings.append(f"sensitive leakage: {run.variant}/{run.scenario_id}")
    by_variant = {name: [run for run in runs if run.variant == name] for name in ("baseline", "observer")}

    def median(name: str, field: str) -> float:
        values = [float(getattr(run, field)) for run in by_variant[name]]
        return statistics.median(values) if values else 0.0

    def reduction(field: str) -> float:
        baseline = median("baseline", field)
        treatment = median("observer", field)
        return (baseline - treatment) / baseline if baseline > 0 else 0.0

    token_reduction = reduction("tokens")
    intervention_reduction = reduction("interventions")
    elapsed_reduction = reduction("elapsed_seconds")
    required_scenarios = set(scenarios)
    coverage_passed = all({run.scenario_id for run in by_variant[name]} >= required_scenarios for name in by_variant)
    if not coverage_passed:
        findings.append("incomplete paired scenario coverage")
    safety_passed = not any(item.startswith(("false complete", "browser disruption", "sensitive leakage")) for item in findings)
    recall_values = {kind: (recalled[kind] / expected_counts[kind] if expected_counts[kind] else 1.0) for kind in recalled}
    outcome_passed = not any(item.startswith("unexpected status") for item in findings)
    value_passed = safety_passed and outcome_passed and coverage_passed and all(value == 1.0 for value in recall_values.values()) and token_reduction >= 0.5 and intervention_reduction >= 0.5 and elapsed_reduction >= 0.2
    return {
        "status": "demonstrated" if value_passed else "not_demonstrated",
        "local_only": True,
        "run_count": len(runs), "accepted_complete_runs": accepted,
        "recall": recall_values,
        "candidate_accuracy": {name: (sum(run.candidates_correct for run in values) / sum(run.candidates_total for run in values)) if sum(run.candidates_total for run in values) else 0.0 for name, values in by_variant.items()},
        "median": {name: {field: median(name, field) for field in ("tokens", "elapsed_seconds", "tool_calls", "interventions")} for name in by_variant},
        "token_reduction": token_reduction, "intervention_reduction": intervention_reduction, "elapsed_reduction": elapsed_reduction,
        "coverage_passed": coverage_passed,
        "findings": findings,
        "recommendation": "continue observer development" if value_passed else "use existing Chrome DevTools or Playwright tooling until the value gates pass",
    }


def main() -> int:
    if len(sys.argv) != 3:
        print("usage: python -m course_selection.observer_evals MANIFEST.json RUNS.json")
        return 2
    manifest = json.loads(Path(sys.argv[1]).read_text(encoding="utf-8"))
    payload = json.loads(Path(sys.argv[2]).read_text(encoding="utf-8"))
    report = evaluate(manifest, [RunMetrics.from_dict(item) for item in payload])
    print(json.dumps(report, ensure_ascii=False, indent=2))
    return 0 if report["status"] == "demonstrated" else 1


if __name__ == "__main__":
    raise SystemExit(main())
