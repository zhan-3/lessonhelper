import json
from pathlib import Path

from course_selection.observer_evals import RunMetrics, evaluate, hidden_variant

MANIFEST = Path(__file__).parents[1] / ".scratch/academic-browser-observer/evals/public-scenarios.json"


def run(variant, *, tokens, interventions, output="", browser_alive=True, status="complete"):
    return RunMetrics.from_dict({
        "scenario_id": "auth-redirect", "variant": variant,
        "tokens": tokens, "elapsed_seconds": tokens / 100,
        "tool_calls": 8 if variant == "baseline" else 3,
        "interventions": interventions, "status": status,
        "targets": ["protected-page"], "requests": ["protected-read"],
        "redirects": ["login-to-protected"],
        "candidates_correct": 1, "candidates_total": 1,
        "browser_alive": browser_alive, "serialized_output": output,
    })


def one_scenario_manifest():
    manifest = json.loads(MANIFEST.read_text(encoding="utf-8"))
    manifest["scenarios"] = [item for item in manifest["scenarios"] if item["id"] == "auth-redirect"]
    return manifest


def test_value_gate_requires_safety_and_half_reduction():
    manifest = one_scenario_manifest()
    report = evaluate(manifest, [
        run("baseline", tokens=1000, interventions=4),
        run("observer", tokens=400, interventions=1),
    ])
    assert report["status"] == "demonstrated"
    assert report["recall"] == {"targets": 1.0, "requests": 1.0, "redirects": 1.0}

    leaked = evaluate(manifest, [
        run("baseline", tokens=1000, interventions=4),
        run("observer", tokens=400, interventions=1, output="EVAL_TOKEN_DO_NOT_LOG"),
    ])
    assert leaked["status"] == "not_demonstrated"
    assert "Chrome DevTools or Playwright" in leaked["recommendation"]


def test_false_complete_and_browser_disruption_fail_closed():
    manifest = one_scenario_manifest()
    missing = run("observer", tokens=1, interventions=0, browser_alive=False)
    missing = RunMetrics(**{**missing.__dict__, "requests": frozenset()})
    report = evaluate(manifest, [run("baseline", tokens=10, interventions=1), missing])
    assert any(item.startswith("false complete") for item in report["findings"])
    assert any(item.startswith("browser disruption") for item in report["findings"])


def test_manifest_expected_partial_rejects_false_complete():
    manifest = json.loads(MANIFEST.read_text(encoding="utf-8"))
    manifest["scenarios"] = [item for item in manifest["scenarios"] if item["id"] == "blank-popup"]
    value = run("observer", tokens=1, interventions=0)
    value = RunMetrics(**{**value.__dict__, "scenario_id": "blank-popup", "targets": frozenset({"popup-unresolved"}), "requests": frozenset(), "redirects": frozenset()})
    baseline = RunMetrics(**{**value.__dict__, "variant": "baseline"})
    report = evaluate(manifest, [baseline, value])
    assert any(item.startswith("unexpected status") for item in report["findings"])
    assert report["status"] == "not_demonstrated"


def test_hidden_variants_change_details_not_expected_outcome():
    manifest = json.loads(MANIFEST.read_text(encoding="utf-8"))
    scenario = manifest["scenarios"][1]
    first = hidden_variant(scenario, 1)
    second = hidden_variant(scenario, 9)
    assert first["variant"] != second["variant"]
    assert first["expected"] == second["expected"]
    assert first["failure_class"] == second["failure_class"]
