"""Transport-backed ``LabExamSession`` for one teaching-center app.

Reading and submitting assessments needs nothing beyond the existing
``LabTransport`` seam: every call is a POST to
``/{center}/StuApi/view/exam/...`` carrying the ``vctchauthorization`` header,
exactly like the booking endpoints.  No browser object is involved, so this
adapter works with either backend (plain HTTP or in-page fetch).

Note the asymmetry in error handling: ``exam_status`` and ``exam_sheet`` raise
on a failed envelope because there is nothing useful to show, while
``submit_exam`` returns the envelope untouched so the guard in
:mod:`course_selection.lab_exam` can distinguish "not sent" from "sent, outcome
unknown".
"""

from __future__ import annotations

from collections.abc import Mapping
from typing import Any

from .lab_exam import ExamSheet, ExamStatus, parse_exam_sheet, parse_exam_status
from .lab_ports import LabTransport


class TransportLabExam:
    """Read and submit assessments over a ``LabTransport``."""

    def __init__(self, transport: LabTransport):
        self._transport = transport

    def _result(self, path: str, form: Mapping[str, Any]) -> Mapping[str, Any]:
        payload = self._transport.call(path, form)
        if payload.get("transport") != "ok" or payload.get("code") != 0:
            raise RuntimeError(
                f"{path} failed: {payload.get('code')} {payload.get('message')}"
            )
        result = payload.get("result")
        return result if isinstance(result, Mapping) else {}

    def exam_status(self, subject_id: int) -> ExamStatus:
        return parse_exam_status(
            self._result("view/exam/view", {"subjectId": subject_id}),
            subject_id=subject_id,
        )

    def exam_sheet(self, subject_id: int) -> ExamSheet:
        return parse_exam_sheet(self._result("view/exam/list", {"subjectId": subject_id}))

    def submit_exam(self, subject_id: int, payload: str) -> Mapping[str, Any]:
        return self._transport.call(
            "view/exam/do", {"subjectId": subject_id, "answer": payload}
        )
