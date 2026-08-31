import json

import pytest

from course_selection.browser_redaction import TraceRedactor

SECRET = "student-secret-2025000000"


def serialized(value):
    return json.dumps(value, ensure_ascii=False, sort_keys=True)


def test_redacts_sensitive_values_across_supported_channels():
    redactor = TraceRedactor(additional_sensitive_names=("campusKey",))
    evidence = {
        "url": redactor.redact_url(f"https://student.example/query?student_id={SECRET}&page=1"),
        "headers": redactor.redact_headers({"Authorization": SECRET, "X-CampusKey": SECRET}),
        "json": redactor.redact_body(json.dumps({"csrfToken": SECRET, "safe": "shape"}), "application/json"),
        "form": redactor.redact_body(f"password={SECRET}&page=1", "application/x-www-form-urlencoded"),
        "opaque": redactor.redact_body(SECRET, "text/plain"),
        "console": redactor.redact_console([SECRET, {"studentNumber": SECRET}]),
    }
    output = serialized(evidence)
    assert SECRET not in output
    assert "student_id" in output
    assert "Authorization" in output
    assert "safe" in output


def test_fingerprints_correlate_only_within_one_trace():
    first = TraceRedactor()
    same_a = first.redact_headers({"token": SECRET})["token"]["fingerprint"]
    same_b = first.redact_value(SECRET, field_name="session")["fingerprint"]
    second = TraceRedactor()
    other = second.redact_value(SECRET, field_name="session")["fingerprint"]
    assert same_a == same_b
    assert same_a != other
    first.close()
    with pytest.raises(RuntimeError):
        first.redact_value(SECRET, field_name="token")
