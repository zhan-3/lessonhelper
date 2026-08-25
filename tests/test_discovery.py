import unittest

from course_selection.cli import build_parser
from course_selection.discovery import (
    TARGET_SELECTION,
    TARGET_TIMETABLE,
    find_academic_portal_redirect,
    is_mutating_request,
    score_discovery_control,
)


class InterfaceDiscoveryTests(unittest.TestCase):
    def test_selection_navigation_is_allowed_but_write_controls_are_blocked(self):
        self.assertGreater(
            score_discovery_control(TARGET_SELECTION, text="学生选课"),
            0,
        )
        self.assertEqual(
            score_discovery_control(TARGET_SELECTION, text="提交选课"),
            -1,
        )
        self.assertEqual(
            score_discovery_control(TARGET_SELECTION, text="退课"),
            -1,
        )

    def test_timetable_and_intermediate_portal_controls_are_ranked(self):
        timetable = score_discovery_control(TARGET_TIMETABLE, text="我的课表")
        portal = score_discovery_control(TARGET_TIMETABLE, text="本科生综合服务")

        self.assertGreater(timetable, portal)
        self.assertGreater(portal, 0)

    def test_read_queries_pass_and_mutation_requests_are_blocked(self):
        self.assertFalse(
            is_mutating_request("POST", "/course/list", '{"term":"2026"}')
        )
        self.assertTrue(
            is_mutating_request("POST", "/selection/save", '{"courseId":"1"}')
        )
        self.assertTrue(
            is_mutating_request("POST", "/course/drop-course", None)
        )

    def test_cli_exposes_both_automatic_discovery_commands(self):
        timetable = build_parser().parse_args(["discover-timetable"])
        selection = build_parser().parse_args(["discover-selection"])

        self.assertEqual(timetable.discovery_target, TARGET_TIMETABLE)
        self.assertEqual(selection.discovery_target, TARGET_SELECTION)
        self.assertEqual(selection.max_clicks, 8)

    def test_finds_new_academic_system_in_portal_catalog(self):
        payload = {
            "groups": [
                {
                    "resources": [
                        {"name": "教务调查问卷", "redirect": "/survey"},
                        {"name": "新教务系统", "redirect": "/vpn/academic"},
                    ]
                }
            ]
        }

        self.assertEqual(find_academic_portal_redirect(payload), "/vpn/academic")


if __name__ == "__main__":
    unittest.main()
