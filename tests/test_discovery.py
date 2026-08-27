import unittest
from pathlib import Path

from course_selection.cli import build_parser
from course_selection.discovery import (
    TARGET_SELECTION,
    TARGET_TIMETABLE,
    find_academic_portal_redirect,
    is_control_allowed_in_stage,
    is_mutating_request,
    score_discovery_control,
)


class InterfaceDiscoveryTests(unittest.TestCase):
    def test_selection_navigation_is_allowed_but_write_controls_are_blocked(self):
        self.assertGreater(
            score_discovery_control(TARGET_SELECTION, text="学生选课"),
            0,
        )
        self.assertGreater(
            score_discovery_control(TARGET_SELECTION, text="全校任选课选课"),
            score_discovery_control(TARGET_SELECTION, text="学生选课"),
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
        portal = score_discovery_control(
            TARGET_TIMETABLE,
            text="学生选课",
            href="javascript:void(0)",
        )

        self.assertGreater(timetable, portal)
        self.assertGreater(portal, 0)
        self.assertGreater(
            score_discovery_control(TARGET_TIMETABLE, text="统一身份认证登录"),
            0,
        )
        self.assertEqual(
            score_discovery_control(TARGET_TIMETABLE, text="学生事务"),
            -1,
        )
        self.assertEqual(
            score_discovery_control(
                TARGET_TIMETABLE,
                text="导出课表",
                target_page_reached=True,
            ),
            -1,
        )
        self.assertEqual(
            score_discovery_control(
                TARGET_TIMETABLE,
                text="个人课表查询",
                target_page_reached=True,
            ),
            -1,
        )
        self.assertGreater(
            score_discovery_control(
                TARGET_TIMETABLE,
                text="查询",
                target_page_reached=True,
            ),
            0,
        )

    def test_selection_categories_beat_microprogramme_fallback(self):
        core = score_discovery_control(TARGET_SELECTION, text="文化素质核心")
        elective = score_discovery_control(TARGET_SELECTION, text="全校任选课")
        limited = score_discovery_control(TARGET_SELECTION, text="限选课")
        microprogramme = score_discovery_control(TARGET_SELECTION, text="微专业选课")

        self.assertGreater(core, elective)
        self.assertGreater(elective, limited)
        self.assertGreater(limited, microprogramme)

    def test_selection_categories_require_the_student_selection_menu(self):
        category_href = "/xsxk/queryXsxk?pageXklb=szhx"

        self.assertTrue(
            is_control_allowed_in_stage(
                TARGET_SELECTION,
                text="学生选课",
                href="javascript:void(0)",
                selection_menu_expanded=False,
            )
        )
        self.assertFalse(
            is_control_allowed_in_stage(
                TARGET_SELECTION,
                text="创新创业",
                href="javascript:void(0)",
                selection_menu_expanded=False,
            )
        )
        self.assertTrue(
            is_control_allowed_in_stage(
                TARGET_SELECTION,
                text="文化素质核心",
                href=category_href,
                selection_menu_expanded=True,
                allowed_selection_categories=("szhx",),
            )
        )
        self.assertFalse(
            is_control_allowed_in_stage(
                TARGET_SELECTION,
                text="创新创业",
                href="javascript:void(0)",
                selection_menu_expanded=True,
                allowed_selection_categories=("szhx",),
            )
        )
        self.assertFalse(
            is_control_allowed_in_stage(
                TARGET_SELECTION,
                text="创新实验",
                href="/xsxk/queryXsxk?pageXklb=cxsy",
                selection_menu_expanded=True,
                allowed_selection_categories=("szhx",),
            )
        )

    def test_read_queries_pass_and_mutation_requests_are_blocked(self):
        self.assertEqual(
            score_discovery_control(
                TARGET_SELECTION,
                text="查询",
                target_page_reached=True,
            ),
            -1,
        )
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
        self.assertFalse(selection.persistent_session)
        self.assertIsNone(selection.grade)
        self.assertEqual(
            str(selection.notice),
            str(Path(".private") / "academic-selection" / "selection-notice.json"),
        )

        configure = build_parser().parse_args(
            ["configure-login", "--username", "2025000000"]
        )
        self.assertEqual(configure.command, "configure-login")
        self.assertEqual(configure.username, "2025000000")

        profile = build_parser().parse_args(
            ["configure-profile", "--grade", "2025"]
        )
        self.assertEqual(profile.command, "configure-profile")
        self.assertEqual(profile.grade, "2025")

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
