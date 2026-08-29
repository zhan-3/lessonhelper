import unittest
from pathlib import Path
from unittest import mock

from click.testing import CliRunner

from course_selection.cli import (
    configure_login_cmd,
    configure_profile_cmd,
    discover_selection_cmd,
    discover_timetable_cmd,
)
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
        runner = CliRunner()
        with mock.patch("course_selection.cli._run_discovery") as discovery:
            result = runner.invoke(discover_timetable_cmd, [], catch_exceptions=False)
            self.assertEqual(0, result.exit_code)
            result = runner.invoke(discover_selection_cmd, [], catch_exceptions=False)
            self.assertEqual(0, result.exit_code)

        timetable_call, selection_call = discovery.call_args_list
        self.assertEqual(TARGET_TIMETABLE, timetable_call.kwargs["target"])
        self.assertEqual(TARGET_SELECTION, selection_call.kwargs["target"])
        self.assertEqual(8, selection_call.kwargs["max_clicks"])
        self.assertFalse(selection_call.kwargs["persistent_session"])
        self.assertIsNone(selection_call.kwargs["grade"])
        self.assertEqual(
            str(Path(".private") / "academic-selection" / "selection-notice.json"),
            str(selection_call.kwargs["notice"]),
        )

        configure = configure_login_cmd.make_context(
            "configure-login", ["--username", "2025000000"]
        )
        self.assertEqual("2025000000", configure.params["username"])

        profile = configure_profile_cmd.make_context(
            "configure-profile", ["--grade", "2025"]
        )
        self.assertEqual("2025", profile.params["grade"])

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
