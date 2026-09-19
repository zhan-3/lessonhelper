"""Tests for the openlab pre-exam assessment core and its guard."""

import json
import unittest
from typing import Any

from course_selection.lab_exam import (
    GROUP_MULTI,
    GROUP_SINGLE,
    ExamAnswer,
    ExamSheet,
    ExamStatus,
    encode_answers,
    exam_token,
    parse_answer_mapping,
    parse_answer_sheet,
    parse_exam_sheet,
    parse_exam_status,
    render_sheet_plain,
    submit_exam,
    validate_answers,
)
from course_selection.lab_exam_transport import TransportLabExam

# Synthetic payloads shaped like the real /view/exam/* responses.  These are
# test fixtures, not captured data.
ANSWER_KEY = "B"
STATUS_PAYLOAD: dict[str, Any] = {"msg": "可以参加", "msgCode": ""}
BLOCKED_STATUS_PAYLOAD: dict[str, Any] = {"msg": "已参加", "msgCode": "Y"}
SHEET_PAYLOAD: dict[str, Any] = {
    "subjectId": 4242,
    "subjectName": "实验安全预考核",
    "kg1s": [
        {
            "questionId": "q1", "questionTxt": "题干一",
            "answerA": "选项A", "answerB": "选项B",
            "answerC": "", "answerD": "",
            "answerResult": ANSWER_KEY,
        },
        {
            "questionId": "q2", "questionTxt": "题干二",
            "answerA": "甲", "answerB": "乙", "answerC": "丙", "answerD": "丁",
            "answerResult": "C",
        },
    ],
    "kg2s": [
        {
            "questionId": "q3", "questionTxt": "题干三（多选）",
            "answerA": "甲", "answerB": "乙", "answerC": "丙", "answerD": "",
            "answerResult": "A,C",
        },
    ],
}


def sheet() -> ExamSheet:
    return parse_exam_sheet(SHEET_PAYLOAD)


def valid_answers() -> tuple[ExamAnswer, ...]:
    return (
        ExamAnswer("q1", GROUP_SINGLE, ("A",)),
        ExamAnswer("q2", GROUP_SINGLE, ("C",)),
        ExamAnswer("q3", GROUP_MULTI, ("A", "C")),
    )


class FakeExamSession:
    """Records calls and returns canned envelopes."""

    def __init__(
        self,
        *,
        status: ExamStatus | None = None,
        sheet_value: ExamSheet | None = None,
        submit: dict[str, Any] | None = None,
    ):
        self._status = status or parse_exam_status(STATUS_PAYLOAD, subject_id=4242)
        self._sheet = sheet_value or sheet()
        self._submit = submit if submit is not None else {
            "transport": "ok", "code": 0, "message": "", "result": {"jg": "通过"},
        }
        self.calls: list[tuple[str, Any]] = []

    def exam_status(self, subject_id: int) -> ExamStatus:
        self.calls.append(("status", subject_id))
        return self._status

    def exam_sheet(self, subject_id: int) -> ExamSheet:
        self.calls.append(("sheet", subject_id))
        return self._sheet

    def submit_exam(self, subject_id: int, payload: str) -> dict[str, Any]:
        self.calls.append(("submit", payload))
        return self._submit


class ParsingTests(unittest.TestCase):
    def test_status_allows_when_code_is_not_blocking(self):
        self.assertTrue(parse_exam_status(STATUS_PAYLOAD, subject_id=4242).allowed)

    def test_status_blocks_on_y_and_n(self):
        for code in ("Y", "N"):
            status = parse_exam_status({"msg": "x", "msgCode": code}, subject_id=1)
            self.assertFalse(status.allowed, code)

    def test_sheet_parses_questions_and_options(self):
        parsed = sheet()
        self.assertEqual(4242, parsed.subject_id)
        self.assertEqual("实验安全预考核", parsed.subject_name)
        self.assertEqual(["q1", "q2", "q3"], [q.question_id for q in parsed.questions])
        self.assertEqual([GROUP_SINGLE, GROUP_SINGLE, GROUP_MULTI],
                         [q.group for q in parsed.questions])
        # 空选项被丢弃：q1 只有 A、B
        self.assertEqual(("A", "B"), tuple(label for label, _ in parsed.questions[0].options))
        self.assertEqual(("A", "B", "C", "D"),
                         tuple(label for label, _ in parsed.questions[1].options))

    def test_inline_html_is_stripped(self):
        """The API wraps stems and options in ``<p>`` / ``<br/>``."""
        payload = {
            "subjectId": 1, "subjectName": "s",
            "kg1s": [{
                "questionId": "q", "questionTxt": "<p>题干<br/>换行</p>",
                "answerA": "<p>选项  A</p>", "answerB": "", "answerC": "", "answerD": "",
                "answerResult": "A",
            }],
        }
        question = parse_exam_sheet(payload).questions[0]
        self.assertEqual("题干 换行", question.text)
        self.assertEqual((("A", "选项 A"),), question.options)

    def test_parser_ignores_the_server_answer_key(self):
        """The core must not read, expose, or forward ``answerResult``.

        The answer key is a letter that also appears as an option label, so a
        string search would be meaningless — assert the *shape* instead: a
        parsed question has no field that could hold an answer.
        """
        parsed = sheet()
        for question in parsed.questions:
            self.assertEqual(
                {"question_id", "group", "text", "options"},
                set(question.to_dict()),
                "parsed question must carry no answer field",
            )
        self.assertNotIn("answerResult", json.dumps(parsed.to_dict(), ensure_ascii=False))


class ValidationTests(unittest.TestCase):
    def test_complete_answer_set_is_valid(self):
        self.assertEqual((), validate_answers(sheet(), valid_answers()))

    def test_missing_and_unknown_questions_are_reported(self):
        problems = validate_answers(sheet(), (ExamAnswer("q1", GROUP_SINGLE, ("A",)),))
        self.assertTrue(any("未作答: q2" in p for p in problems))
        self.assertTrue(any("未作答: q3" in p for p in problems))
        problems = validate_answers(
            sheet(), valid_answers() + (ExamAnswer("nope", GROUP_SINGLE, ("A",)),)
        )
        self.assertTrue(any("题目不存在: nope" in p for p in problems))

    def test_empty_answer_is_reported(self):
        answers = (ExamAnswer("q1", GROUP_SINGLE, ()),) + valid_answers()[1:]
        self.assertTrue(any("答案为空: q1" in p for p in validate_answers(sheet(), answers)))

    def test_single_choice_must_have_exactly_one_option(self):
        answers = (ExamAnswer("q1", GROUP_SINGLE, ("A", "B")),) + valid_answers()[1:]
        self.assertTrue(any("单选题" in p for p in validate_answers(sheet(), answers)))

    def test_option_outside_the_question_is_reported(self):
        answers = (ExamAnswer("q1", GROUP_SINGLE, ("Z",)),) + valid_answers()[1:]
        self.assertTrue(any("选项不在题干中: q1" in p for p in validate_answers(sheet(), answers)))

    def test_duplicate_question_is_reported(self):
        problems = validate_answers(sheet(), valid_answers() + valid_answers()[:1])
        self.assertTrue(any("重复作答: q1" in p for p in problems))


class AnswerMappingTests(unittest.TestCase):
    """The answer file carries options only; the group comes from the sheet."""

    def test_group_is_taken_from_the_sheet_not_the_file(self):
        answers = parse_answer_mapping({"q1": ["A"], "q3": ["A", "C"]}, sheet())
        by_id = {answer.question_id: answer for answer in answers}
        self.assertEqual(GROUP_SINGLE, by_id["q1"].group)
        self.assertEqual(GROUP_MULTI, by_id["q3"].group)

    def test_scalar_value_is_accepted_as_one_option(self):
        answers = parse_answer_mapping({"q1": "A"}, sheet())
        self.assertEqual(("A",), answers[0].selected)

    def test_unknown_question_is_rejected(self):
        with self.assertRaisesRegex(ValueError, "不在本次考核中"):
            parse_answer_mapping({"nope": ["A"]}, sheet())

    def test_incomplete_mapping_surfaces_through_validation(self):
        problems = validate_answers(sheet(), parse_answer_mapping({"q1": ["A"]}, sheet()))
        self.assertTrue(any("未作答: q2" in problem for problem in problems))


class PlainSheetTests(unittest.TestCase):
    """The offline template: render, fill in ``答:``, read back."""

    def test_untouched_template_yields_no_answers(self):
        self.assertEqual({}, parse_answer_sheet(render_sheet_plain(sheet())))

    def test_filled_template_binds_to_the_right_question(self):
        rendered = render_sheet_plain(sheet())
        filled = rendered.replace("答: \n", "答: A\n", 1)   # 只填第一题
        parsed = parse_answer_sheet(filled)
        self.assertEqual(["A"], parsed[sheet().questions[0].question_id])
        self.assertEqual(1, len(parsed))

    def test_multi_select_line_is_split(self):
        self.assertEqual({"q9": ["A", "C"]}, parse_answer_sheet("#qid q9\n答: A,C\n"))
        self.assertEqual({"q9": ["A", "C"]}, parse_answer_sheet("#qid q9\n答: AC\n"))

    def test_answer_before_any_qid_is_rejected(self):
        with self.assertRaisesRegex(ValueError, "之前"):
            parse_answer_sheet("答: A\n#qid q1\n")

    def test_template_carries_every_question_id(self):
        rendered = render_sheet_plain(sheet())
        for question in sheet().questions:
            self.assertIn(f"#qid {question.question_id}", rendered)
            self.assertIn(question.text, rendered)

    def test_empty_template_cannot_pass_validation(self):
        parsed = parse_answer_mapping(parse_answer_sheet(render_sheet_plain(sheet())), sheet())
        self.assertEqual((), parsed)
        problems = validate_answers(sheet(), parsed)
        self.assertEqual(3, len([p for p in problems if "未作答" in p]))


class EncodingTests(unittest.TestCase):
    def test_single_choice_encodes_as_value_and_multi_as_list(self):
        encoded = json.loads(encode_answers(valid_answers()))
        by_id = {item["id"]: item for item in encoded}
        self.assertEqual({"id": "q1", "type": GROUP_SINGLE, "answer": "A"}, by_id["q1"])
        self.assertEqual({"id": "q3", "type": GROUP_MULTI, "answer": ["A", "C"]}, by_id["q3"])

    def test_encoding_is_deterministic_regardless_of_input_order(self):
        shuffled = tuple(reversed(valid_answers()))
        self.assertEqual(encode_answers(valid_answers()), encode_answers(shuffled))

    def test_token_is_stable_and_content_bound(self):
        answers = valid_answers()
        self.assertEqual(exam_token(4242, answers), exam_token(4242, answers))
        self.assertNotEqual(exam_token(4242, answers), exam_token(4243, answers))
        changed = (ExamAnswer("q1", GROUP_SINGLE, ("B",)),) + answers[1:]
        self.assertNotEqual(exam_token(4242, answers), exam_token(4242, changed))


class SubmitGuardTests(unittest.TestCase):
    def test_confirmation_token_must_match(self):
        with self.assertRaisesRegex(ValueError, "confirmation token mismatch"):
            submit_exam(FakeExamSession(), 4242, valid_answers(), confirmation="wrong")

    def test_blocked_status_stops_before_reading_questions(self):
        session = FakeExamSession(
            status=parse_exam_status(BLOCKED_STATUS_PAYLOAD, subject_id=4242)
        )
        outcome = submit_exam(
            session, 4242, valid_answers(),
            confirmation=exam_token(4242, valid_answers()),
        )
        self.assertEqual("blocked_by_status", outcome["outcome"])
        self.assertEqual(["status"], [name for name, _ in session.calls])

    def test_invalid_answers_stop_before_submitting(self):
        session = FakeExamSession()
        incomplete = (ExamAnswer("q1", GROUP_SINGLE, ("A",)),)
        outcome = submit_exam(
            session, 4242, incomplete, confirmation=exam_token(4242, incomplete)
        )
        self.assertEqual("invalid_answers", outcome["outcome"])
        self.assertNotIn("submit", [name for name, _ in session.calls])

    def test_unclear_transport_becomes_possibly_applied(self):
        session = FakeExamSession(submit={"transport": "timeout", "detail": "no route"})
        outcome = submit_exam(
            session, 4242, valid_answers(),
            confirmation=exam_token(4242, valid_answers()),
        )
        self.assertEqual("possibly_applied", outcome["outcome"])
        self.assertEqual(1, len([1 for name, _ in session.calls if name == "submit"]))

    def test_business_failure_is_confirmed_failure(self):
        session = FakeExamSession(submit={"transport": "ok", "code": 500, "message": "服务器错误"})
        outcome = submit_exam(
            session, 4242, valid_answers(),
            confirmation=exam_token(4242, valid_answers()),
        )
        self.assertEqual("confirmed_failure", outcome["outcome"])

    def test_success_reports_the_verdict(self):
        session = FakeExamSession()
        outcome = submit_exam(
            session, 4242, valid_answers(),
            confirmation=exam_token(4242, valid_answers()),
        )
        self.assertEqual("confirmed_success", outcome["outcome"])
        self.assertEqual("通过", outcome["verdict"])
        # 恰好一次提交
        self.assertEqual(1, len([1 for name, _ in session.calls if name == "submit"]))

    def test_status_and_sheet_are_reread_before_submitting(self):
        session = FakeExamSession()
        submit_exam(session, 4242, valid_answers(),
                    confirmation=exam_token(4242, valid_answers()))
        self.assertEqual(["status", "sheet", "submit"], [name for name, _ in session.calls])


class FakeTransport:
    """Records transport calls; returns a canned envelope per path."""

    def __init__(self, responses: dict[str, dict[str, Any]]):
        self.center = "dxwl"
        self._responses = responses
        self.calls: list[tuple[str, Any]] = []

    def call(self, path: str, form: Any = None) -> dict[str, Any]:
        self.calls.append((path, form))
        return self._responses.get(path, {"transport": "ok", "code": 0, "result": {}})

    def close(self) -> None:
        pass


class TransportAdapterTests(unittest.TestCase):
    def test_paths_and_parameters_match_the_web_client(self):
        transport = FakeTransport({
            "view/exam/view": {"transport": "ok", "code": 0, "result": STATUS_PAYLOAD},
            "view/exam/list": {"transport": "ok", "code": 0, "result": SHEET_PAYLOAD},
            "view/exam/do": {"transport": "ok", "code": 0, "result": {"jg": "通过"}},
        })
        exam = TransportLabExam(transport)

        status = exam.exam_status(4242)
        sheet_value = exam.exam_sheet(4242)
        raw = exam.submit_exam(4242, encode_answers(valid_answers()))

        self.assertTrue(status.allowed)
        self.assertEqual(3, len(sheet_value.questions))
        self.assertEqual({"jg": "通过"}, raw["result"])
        self.assertEqual(
            [("view/exam/view", {"subjectId": 4242}),
             ("view/exam/list", {"subjectId": 4242}),
             ("view/exam/do", {"subjectId": 4242,
                               "answer": encode_answers(valid_answers())})],
            transport.calls,
        )

    def test_failed_envelope_raises_for_reads(self):
        transport = FakeTransport({
            "view/exam/view": {"transport": "ok", "code": 500, "message": "boom"},
        })
        with self.assertRaisesRegex(RuntimeError, "view/exam/view failed"):
            TransportLabExam(transport).exam_status(4242)

    def test_submit_returns_the_envelope_untouched(self):
        """The guard needs the raw envelope to tell 'not sent' from 'unknown'."""
        transport = FakeTransport({
            "view/exam/do": {"transport": "timeout", "detail": "no route"},
        })
        payload = TransportLabExam(transport).submit_exam(4242, "[]")
        self.assertEqual("timeout", payload["transport"])


if __name__ == "__main__":
    unittest.main()
