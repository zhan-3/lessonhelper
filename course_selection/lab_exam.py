"""Read and submit openlab pre-exam assessments for one teaching-center app.

Reading (status, question sheet) is safe and repeatable.  Submission is a write
and is guarded at every step the web client simply trusts:

* the status endpoint is re-read immediately before submitting, and a
  ``msgCode`` of ``Y`` or ``N`` blocks the submission outright;
* the answers are validated against the *freshly read* question set (ids match,
  nothing empty, single-choice questions carry exactly one option);
* the submission carries an explicit confirmation token bound to the subject
  and the exact answer set;
* at most one POST is issued, and an unclear transport outcome becomes
  ``possibly_applied`` and is never retried.

``parse_exam_sheet`` deliberately **ignores** the ``answerResult`` field that
the API returns alongside each question.  This module reads questions and
submits the answers a human supplies; it does not read, expose, or forward the
server-provided answer key, so it cannot drift into auto-answering.
"""

from __future__ import annotations

import hashlib
import json
import re
from collections.abc import Mapping, Sequence
from dataclasses import dataclass
from typing import Any, Protocol

# ``Y``/``N`` both mean "this assessment cannot be taken right now"; anything
# else means the sheet may be loaded.
BLOCKING_MESSAGE_CODES = frozenset({"Y", "N"})

OPTION_LABELS = ("A", "B", "C", "D", "E", "F")

# Question groups carried by ``view/exam/list``.  kg1s questions take a single
# option; kg2s questions take an array.  The web client encodes them the same
# way, so the wire shape is preserved rather than normalised.
GROUP_SINGLE = 1
GROUP_MULTI = 2
_GROUP_FIELDS = ((GROUP_SINGLE, "kg1s"), (GROUP_MULTI, "kg2s"))


@dataclass(frozen=True)
class ExamStatus:
    """Pre-check result from ``view/exam/view``."""

    subject_id: int
    message: str
    message_code: str

    @property
    def allowed(self) -> bool:
        return self.message_code not in BLOCKING_MESSAGE_CODES

    def to_dict(self) -> dict[str, Any]:
        return {
            "subject_id": self.subject_id,
            "message": self.message,
            "message_code": self.message_code,
            "allowed": self.allowed,
        }


@dataclass(frozen=True)
class ExamQuestion:
    """One question, without its answer key."""

    question_id: str
    group: int
    text: str
    options: tuple[tuple[str, str], ...]

    def to_dict(self) -> dict[str, Any]:
        return {
            "question_id": self.question_id,
            "group": self.group,
            "text": self.text,
            "options": [{"label": label, "text": text} for label, text in self.options],
        }


@dataclass(frozen=True)
class ExamSheet:
    """The full question set for one assessment."""

    subject_id: int
    subject_name: str
    questions: tuple[ExamQuestion, ...]

    def to_dict(self) -> dict[str, Any]:
        return {
            "subject_id": self.subject_id,
            "subject_name": self.subject_name,
            "questions": [question.to_dict() for question in self.questions],
        }


@dataclass(frozen=True)
class ExamAnswer:
    """A human-supplied answer for one question."""

    question_id: str
    group: int
    selected: tuple[str, ...]

    @property
    def key(self) -> str:
        return self.question_id

    def to_payload(self) -> dict[str, Any]:
        """Encode for ``view/exam/do``: single choice as a value, multi as a list."""
        answer: Any = list(self.selected) if self.group == GROUP_MULTI else (
            self.selected[0] if self.selected else None
        )
        return {"id": self.question_id, "type": self.group, "answer": answer}

    def to_dict(self) -> dict[str, Any]:
        return {
            "question_id": self.question_id,
            "group": self.group,
            "selected": list(self.selected),
        }


class LabExamSession(Protocol):
    """Read/write surface for one center's assessments; faked in tests."""

    def exam_status(self, subject_id: int) -> ExamStatus: ...

    def exam_sheet(self, subject_id: int) -> ExamSheet: ...

    def submit_exam(self, subject_id: int, payload: str) -> Mapping[str, Any]: ...


def _as_int(value: Any) -> int:
    try:
        return int(value)
    except (TypeError, ValueError):
        return 0


_HTML_TAG = re.compile(r"<[^>]+>")


def _plain_text(value: Any) -> str:
    """Strip the inline HTML the API embeds in question and option text.

    The observed payloads wrap every stem and option in ``<p>`` (and use
    ``<br/>`` for line breaks); the terminal has no use for them.
    """
    return " ".join(_HTML_TAG.sub(" ", str(value or "")).split())


def parse_exam_status(payload: Mapping[str, Any], *, subject_id: int) -> ExamStatus:
    """Build a status from the ``{msg, msgCode}`` result envelope."""
    return ExamStatus(
        subject_id=subject_id,
        message=str(payload.get("msg") or ""),
        message_code=str(payload.get("msgCode") or ""),
    )


def parse_exam_sheet(payload: Mapping[str, Any]) -> ExamSheet:
    """Build a sheet from the ``{kg1s, kg2s}`` result envelope.

    ``answerResult`` is intentionally not read — see the module docstring.
    """
    questions: list[ExamQuestion] = []
    for group, field in _GROUP_FIELDS:
        for raw in payload.get(field) or ():
            if not isinstance(raw, Mapping):
                continue
            question_id = str(raw.get("questionId") or "").strip()
            if not question_id:
                continue
            options: list[tuple[str, str]] = []
            for label in OPTION_LABELS:
                text = _plain_text(raw.get(f"answer{label}"))
                if text:
                    options.append((label, text))
            questions.append(ExamQuestion(
                question_id=question_id,
                group=group,
                text=_plain_text(raw.get("questionTxt")),
                options=tuple(options),
            ))
    return ExamSheet(
        subject_id=_as_int(payload.get("subjectId")),
        subject_name=str(payload.get("subjectName") or "").strip(),
        questions=tuple(questions),
    )


def validate_answers(sheet: ExamSheet, answers: Sequence[ExamAnswer]) -> tuple[str, ...]:
    """Return every reason the answer set cannot be submitted (empty = valid)."""
    problems: list[str] = []
    expected = {question.question_id: question for question in sheet.questions}
    given: dict[str, ExamAnswer] = {}
    for answer in answers:
        if answer.question_id in given:
            problems.append(f"重复作答: {answer.question_id}")
        given[answer.question_id] = answer

    for missing in sorted(set(expected) - set(given)):
        problems.append(f"未作答: {missing}")
    for unknown in sorted(set(given) - set(expected)):
        problems.append(f"题目不存在: {unknown}")

    for question_id, answer in given.items():
        question = expected.get(question_id)
        if question is None:
            continue
        if not answer.selected or any(not value.strip() for value in answer.selected):
            problems.append(f"答案为空: {question_id}")
            continue
        allowed = {label for label, _ in question.options}
        outside = sorted(set(answer.selected) - allowed)
        if outside:
            problems.append(f"选项不在题干中: {question_id} -> {', '.join(outside)}")
        if question.group == GROUP_SINGLE and len(answer.selected) != 1:
            problems.append(f"单选题必须且只能选一项: {question_id}")
        if question.group == GROUP_MULTI and len(set(answer.selected)) != len(answer.selected):
            problems.append(f"多选题存在重复选项: {question_id}")
    return tuple(problems)


ANSWER_LINE = re.compile(r"^答\s*[:：]\s*([A-Za-z](?:\s*[,、]?\s*[A-Za-z])*)?\s*$")
_QID_LINE = re.compile(r"^#qid (\S+)")


def render_sheet_plain(sheet: ExamSheet) -> str:
    """Render a sheet for offline work: numbered questions plus a blank answer line.

    ``#qid`` and ``答:`` are emitted per question so a filled-in copy can be fed
    straight back via :func:`parse_answer_sheet` — the id is never transcribed
    by hand, so it cannot drift from the question it belongs to.
    """
    lines: list[str] = []
    for index, question in enumerate(sheet.questions, start=1):
        lines.append(f"[{index}] {question.text}")
        for label, text in question.options:
            lines.append(f"{label}. {text}")
        lines.append(f"#qid {question.question_id}")
        lines.append("答: ")
        lines.append("")
    return "\n".join(lines)


def parse_answer_sheet(text: str) -> dict[str, list[str]]:
    """Read ``答: X`` lines back, keyed by the ``#qid`` that precedes them.

    Unanswered questions are simply absent, which :func:`validate_answers`
    then reports as ``未作答`` — an untouched template never submits silently.
    """
    answers: dict[str, list[str]] = {}
    question_id: str | None = None
    for line in text.splitlines():
        if match := _QID_LINE.match(line.strip()):
            question_id = match.group(1)
            continue
        if match := ANSWER_LINE.match(line.strip()):
            if question_id is None:
                raise ValueError("答案行出现在任何 #qid 之前")
            chosen = [letter.upper() for letter in re.findall(r"[A-Za-z]", match.group(1) or "")]
            if chosen:
                answers[question_id] = chosen
    return answers


def parse_answer_mapping(data: Mapping[str, Any], sheet: ExamSheet) -> tuple[ExamAnswer, ...]:
    """Build answers from a user-supplied ``{question_id: [option, ...]}`` mapping.

    The question *group* is taken from the freshly read sheet rather than from
    the file, so a stale or hand-edited answer file cannot mislabel a
    single-choice question as multi-choice (or the reverse).
    """
    by_id = {question.question_id: question for question in sheet.questions}
    answers: list[ExamAnswer] = []
    for raw_id, raw_selected in data.items():
        question = by_id.get(str(raw_id))
        if question is None:
            raise ValueError(f"答案里的题目不在本次考核中: {raw_id}")
        if isinstance(raw_selected, Sequence) and not isinstance(raw_selected, (str, bytes)):
            values = [str(value) for value in raw_selected]
        else:
            values = [str(raw_selected)]
        answers.append(ExamAnswer(
            question_id=question.question_id,
            group=question.group,
            selected=tuple(values),
        ))
    return tuple(answers)


def encode_answers(answers: Sequence[ExamAnswer]) -> str:
    """Encode answers the way the web client does: a JSON *string* body value."""
    ordered = sorted(answers, key=lambda answer: answer.key)
    return json.dumps([answer.to_payload() for answer in ordered], ensure_ascii=False)


def exam_token(subject_id: int, answers: Sequence[ExamAnswer]) -> str:
    """Deterministic confirmation token binding a run to its exact content."""
    digest = hashlib.sha256()
    digest.update(f"{subject_id}\n".encode())
    for answer in sorted(answers, key=lambda item: item.key):
        digest.update(f"{answer.question_id}|{','.join(answer.selected)}\n".encode())
    return digest.hexdigest()[:16]


def submit_exam(
    session: LabExamSession,
    subject_id: int,
    answers: Sequence[ExamAnswer],
    *,
    confirmation: str,
) -> dict[str, Any]:
    """Submit one assessment at most once; refuse anything unclear."""
    targets = tuple(answers)
    expected = exam_token(subject_id, targets)
    if confirmation != expected:
        raise ValueError(f"confirmation token mismatch: expected {expected}")

    status = session.exam_status(subject_id)
    if not status.allowed:
        return {
            "outcome": "blocked_by_status",
            "status": status.to_dict(),
            "detail": status.message,
        }

    sheet = session.exam_sheet(subject_id)
    problems = validate_answers(sheet, targets)
    if problems:
        return {"outcome": "invalid_answers", "problems": list(problems)}

    payload = session.submit_exam(subject_id, encode_answers(targets))
    if payload.get("transport") != "ok":
        return {
            "outcome": "possibly_applied",
            "detail": str(payload.get("detail") or payload.get("transport")),
        }

    code = payload.get("code")
    message = str(payload.get("message") or "")
    if code != 0:
        return {"outcome": "confirmed_failure", "detail": f"{code} {message}"}

    result = payload.get("result")
    verdict = str(result.get("jg") or "") if isinstance(result, Mapping) else ""
    return {"outcome": "confirmed_success", "verdict": verdict, "detail": message}
