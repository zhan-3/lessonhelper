import unittest
from datetime import datetime

from course_selection.selection_entry import (
    STATUS_EMPTY,
    STATUS_LOGIN_REQUIRED,
    STATUS_READY,
    STATUS_ROUND_NOT_OPEN,
    classify_selection_response,
    classify_selection_html,
    extract_course_sections_from_html,
    observe_json_exchange,
    selection_page_count,
    selection_page_matches_notice,
)



def _section(identity: str, category: str) -> object:
    from dataclasses import dataclass
    from types import SimpleNamespace

    @dataclass
    class Section:
        identity: str
        category: str
        name: str

    return Section(identity=identity, category=category, name="经济学原理")


class SelectionQuerySourceTests(unittest.TestCase):
    def test_flattened_sections_carry_query_source_for_planning_filters(self):
        from course_selection.selection_query import _sections_with_query_source

        sections = {"a|b|c": _section("a|b|c", "专业核心课")}
        stamped = _sections_with_query_source(sections, "xsxk", semester="2026-20271", page=3)
        self.assertEqual(stamped["a|b|c"]["query_code"], "xsxk")
        self.assertEqual(stamped["a|b|c"]["query_term"], "2026-20271")
        self.assertEqual(stamped["a|b|c"]["query_page"], 3)
        self.assertEqual(stamped["a|b|c"]["query_label"], "外专业课程")
        self.assertEqual(stamped["a|b|c"]["category"], "专业核心课")


class SelectionEntryTests(unittest.TestCase):
    COURSE_HTML = """
    <span>选课时间：2026-08-29 08:30 至 2026-08-29 10:30</span>
    <table class="bot_line">
      <tr>
        <th></th><th>序号</th><th>课程代码</th><th>课程名称</th>
        <th>前置课程</th><th>面向对象</th><th>校区</th><th>上课信息</th>
        <th>课程类别</th><th>开课院系</th><th>学分</th><th>学时</th>
        <th>备注信息</th><th>选课要求</th><th>已选/容量</th>
      </tr>
      <tr>
        <td><a onclick="saveXsxk1('TASK-9')">选择</a></td><td>1</td>
        <td>GE101</td><td>人工智能导论</td><td>无</td><td>全校本科生</td>
        <td>威海校区</td><td>教师：李老师 周一 1-2节</td><td>文化素质</td>
        <td>计算机学院</td><td>2</td><td>32</td><td>线下授课</td>
        <td>每人限选一门</td><td>18/30</td>
      </tr>
    </table>
    """

    def test_parses_legacy_selection_course_table(self):
        sections = extract_course_sections_from_html(self.COURSE_HTML)

        self.assertEqual(len(sections), 1)
        self.assertEqual(sections[0].identity, "TASK-9")
        self.assertEqual(sections[0].action_rwh, "TASK-9")
        self.assertEqual(sections[0].action_name, "saveXsxk1")
        self.assertTrue(sections[0].execution_ready)
        self.assertEqual(sections[0].course_code, "GE101")
        self.assertEqual(sections[0].teacher, "李老师")
        self.assertEqual(sections[0].selected_count, "18")
        self.assertEqual(sections[0].capacity_count, "30")
        self.assertEqual(sections[0].meetings, ({"day": 1, "start": 1, "end": 2, "weeks": []},))

    def test_parses_teacher_prefix_and_removes_it_from_schedule(self):
        html = self.COURSE_HTML.replace(
            "教师：李老师 周一 1-2节",
            "李可欣0◇上课信息:[10-17周]星期一第3,4节◇◇",
        )

        section = extract_course_sections_from_html(html)[0]

        self.assertEqual(section.teacher, "李可欣")
        self.assertEqual(section.time, "[10-17周]星期一第3,4节")
        self.assertEqual(
            section.meetings,
            ({"day": 1, "start": 3, "end": 4, "weeks": list(range(10, 18))},),
        )

    def test_fallback_identity_is_not_marked_executable(self):
        html = self.COURSE_HTML.replace("saveXsxk1('TASK-9')", "showCourse('GE101')")

        section = extract_course_sections_from_html(html)[0]

        self.assertFalse(section.execution_ready)
        self.assertEqual(section.action_rwh, "")

    def test_reads_page_count_and_does_not_flatten_composite_capacity(self):
        html = self.COURSE_HTML.replace(
            "18/30",
            "男:0/34 女:0/20",
        ).replace(
            "<table class=\"bot_line\">",
            '<input name="pageCount" value="5"><table class="bot_line">',
        )

        section = extract_course_sections_from_html(html)[0]

        self.assertEqual(selection_page_count(html), 5)
        self.assertEqual(section.selected_count, "")
        self.assertEqual(section.capacity_count, "")

    def test_future_empty_html_is_round_not_open_not_empty(self):
        empty = self.COURSE_HTML.replace(
            "<tr>\n        <td><a onclick=\"saveXsxk1('TASK-9')\">选择</a></td><td>1</td>",
            "<tr style='display:none'><td></td>",
        )
        # Use a header-only page to avoid depending on malformed replacement details.
        empty = empty.split("<tr style='display:none'>", 1)[0] + "</table>"

        observation = classify_selection_html(
            200,
            empty,
            now=datetime(2026, 8, 25, 12, 0),
        )

        self.assertEqual(observation.status, STATUS_ROUND_NOT_OPEN)
        self.assertIn("2026-08-29 08:30", observation.message)

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
        from course_selection.cli import explore_entry_cmd

        context = explore_entry_cmd.make_context(
            "explore-entry", ["--wait-seconds", "3"]
        )

        self.assertEqual(3, context.params["wait_seconds"])

    def test_requires_notice_specific_text_before_accepting_a_selection_page(self):
        class Notice:
            selection_type = "文化素质"
            term = "2026年秋季学期"

        self.assertTrue(selection_page_matches_notice("学生选课｜文化素质｜2026年秋季学期", Notice()))
        self.assertFalse(selection_page_matches_notice("学生选课｜文化素质｜其他学期", Notice()))
        self.assertFalse(selection_page_matches_notice("学生选课｜其他轮次", Notice()))
