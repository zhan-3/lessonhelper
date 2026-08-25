import unittest

from course_selection.selection_entry import (
    STATUS_EMPTY,
    STATUS_LOGIN_REQUIRED,
    STATUS_READY,
    STATUS_ROUND_NOT_OPEN,
    classify_selection_response,
    observe_json_exchange,
    selection_page_matches_notice,
)


class SelectionEntryTests(unittest.TestCase):
    def test_classifies_a_course_list_and_extracts_safe_section_fields(self):
        observation = classify_selection_response(
            200,
            {
                "data": {
                    "list": [
                        {
                            "courseId": "EL-101",
                            "courseName": "人工智能导论",
                            "credits": 2,
                            "teacherName": "李老师",
                            "classTime": "周一 1-2",
                            "remaining": 12,
                            "selected": False,
                        }
                    ]
                }
            },
            request_url="https://academic.example/selection/list",
        )

        self.assertEqual(observation.status, STATUS_READY)
        self.assertEqual(observation.sections[0].identity, "EL-101")
        self.assertEqual(observation.sections[0].credits, "2")
        self.assertFalse(observation.sections[0].selected)

    def test_distinguishes_login_round_not_open_and_empty_states(self):
        self.assertEqual(
            classify_selection_response(401, {"message": "未登录"}).status,
            STATUS_LOGIN_REQUIRED,
        )
        self.assertEqual(
            classify_selection_response(200, {"message": "尚未开始选课"}).status,
            STATUS_ROUND_NOT_OPEN,
        )
        self.assertEqual(
            classify_selection_response(200, {"data": {"list": []}}).status,
            STATUS_EMPTY,
        )

    def test_contract_redacts_sensitive_request_and_response_values(self):
        observation, contract = observe_json_exchange(
            method="POST",
            url="https://academic.example/selection/list?ticket=secret&term=2026",
            status_code=200,
            request_body='{"studentNo":"2025000000","term":"2026"}',
            payload={
                "studentNo": "2025000000",
                "courseName": "人工智能导论",
                "courseId": "EL-101",
            },
        )

        self.assertEqual(observation.status, STATUS_READY)
        self.assertNotIn("secret", observation.request_url)
        self.assertNotIn("2025000000", observation.request_url)
        self.assertNotIn("secret", contract.url)
        self.assertNotIn("2025000000", contract.request_body)
        self.assertEqual(contract.response_fields, ("courseId", "courseName"))

    def test_selection_cli_requires_a_confirmed_notice(self):
        from course_selection.cli import build_parser

        args = build_parser().parse_args(["explore-entry", "--wait-seconds", "3"])

        self.assertEqual(args.command, "explore-entry")
        self.assertEqual(args.wait_seconds, 3)

    def test_requires_notice_specific_text_before_accepting_a_selection_page(self):
        class Notice:
            selection_type = "文化素质"
            term = "2026年秋季学期"

        self.assertTrue(selection_page_matches_notice("学生选课｜文化素质｜2026年秋季学期", Notice()))
        self.assertFalse(selection_page_matches_notice("学生选课｜文化素质｜其他学期", Notice()))
        self.assertFalse(selection_page_matches_notice("学生选课｜其他轮次", Notice()))
