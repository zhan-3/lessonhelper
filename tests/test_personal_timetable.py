import unittest
from types import SimpleNamespace

from course_selection.deep_observation import TimetableObservationRequest
from course_selection.gateway import PlaywrightAcademicGateway
from course_selection.personal_timetable import (
    parse_personal_timetable_html,
    parse_timetable_grid_html,
    personal_timetable_parameters,
)


class PersonalTimetableParserTests(unittest.TestCase):
    def test_request_body_contains_only_verified_fields(self):
        self.assertEqual(
            {"fhlj": "kbcx/queryGrkb", "xnxq": "2026-20271"},
            personal_timetable_parameters(
                fhlj=" kbcx/queryGrkb ", xnxq="2026-20271"
            ),
        )

    def test_parses_complete_personal_timetable_rows(self):
        html = """
        <table>
          <tr><th>课程代码</th><th>课程名称</th><th>教学班号</th><th>教师</th>
              <th>星期</th><th>节次</th><th>周次</th><th>单双周</th><th>上课地点</th></tr>
          <tr><td>CS101</td><td>程序设计</td><td>01</td><td>张老师</td>
              <td>星期一</td><td>1-2</td><td>1-16</td><td>全周</td><td>主楼101</td></tr>
        </table>
        """

        entries = parse_personal_timetable_html(html, term="2026-1")

        self.assertEqual(1, len(entries))
        self.assertEqual("CS101", entries[0].course_code)
        self.assertEqual("程序设计", entries[0].course_name)
        self.assertEqual(1, entries[0].weekday)
        self.assertEqual((1, 2), (entries[0].start_period, entries[0].end_period))
        self.assertEqual(tuple(range(1, 17)), entries[0].week_numbers)
        self.assertEqual("张老师", entries[0].teacher)
        self.assertEqual("主楼101", entries[0].location)

    def test_grid_parser_preserves_single_br_between_course_and_schedule(self):
        entries = parse_timetable_grid_html(
            """
            <span>2026 秋季学期 学生个人课表</span>
            <table>
              <tr><th>节次</th><th>星期一</th><th>星期二</th><th>星期三</th>
                  <th>星期四</th><th>星期五</th><th>星期六</th><th>星期日</th></tr>
              <tr><td>上午</td><td>1-2</td>
                  <td>程序设计<br>张老师[1-15单]周A101</td>
                  <td>&nbsp;</td><td></td><td></td><td></td><td></td><td></td></tr>
            </table>
            """
        )

        self.assertEqual(1, len(entries))
        self.assertEqual("程序设计", entries[0].course_name)
        self.assertEqual("张老师", entries[0].teacher)
        self.assertEqual("odd", entries[0].week_parity)
        self.assertEqual("A101", entries[0].location)

    def test_grid_parser_handles_multiple_br_and_later_isolated_week_range(self):
        entries = parse_timetable_grid_html(
            """
            <span>2026秋季学期 学生个人课表</span>
            <table>
              <tr><th>节次</th><th>星期一</th><th>星期二</th><th>星期三</th>
                  <th>星期四</th><th>星期五</th><th>星期六</th><th>星期日</th></tr>
              <tr><td>下午</td><td>5-6</td><td></td>
                  <td>离散数学<br>教师甲[1-5，7-9]周，教师乙[16]周<br>M楼-201</td>
                  <td></td><td></td><td></td><td></td><td></td></tr>
            </table>
            """
        )

        self.assertEqual("离散数学", entries[0].course_name)
        self.assertEqual("教师甲，教师乙", entries[0].teacher)
        self.assertEqual((1, 2, 3, 4, 5, 7, 8, 9, 16), entries[0].week_numbers)
        self.assertEqual("M楼-201", entries[0].location)

    def test_grid_parser_imports_colspan_other_course_as_conflict_unknown(self):
        entries = parse_timetable_grid_html(
            """
            <span>2026 秋季学期 学生个人课表</span>
            <table>
              <tr><th>节次</th><th>星期一</th><th>星期二</th><th>星期三</th>
                  <th>星期四</th><th>星期五</th><th>星期六</th><th>星期日</th></tr>
              <tr><td colspan="9">其它课程：大学物理实验◇岳老师◇1-16</td></tr>
            </table>
            """
        )

        self.assertEqual("2026年秋季学期", entries[0].term)
        self.assertEqual("大学物理实验", entries[0].course_name)
        self.assertEqual("岳老师", entries[0].teacher)
        self.assertEqual(tuple(range(1, 17)), entries[0].week_numbers)
        self.assertIsNone(entries[0].weekday)
        self.assertIsNone(entries[0].start_period)
        self.assertEqual("unknown", entries[0].conflict_status)

    def test_rejects_a_table_without_required_schedule_fields(self):
        with self.assertRaisesRegex(ValueError, "required timetable columns"):
            parse_personal_timetable_html(
                "<table><tr><th>课程名称</th></tr><tr><td>程序设计</td></tr></table>",
                term="2026-1",
            )

    def test_keeps_unparseable_schedule_as_conflict_unknown(self):
        entries = parse_personal_timetable_html(
            """
            <table><tr><th>课程名称</th><th>星期</th><th>节次</th><th>周次</th></tr>
            <tr><td>待定课程</td><td>待定</td><td>待定</td><td>待定</td></tr></table>
            """,
            term="2026-1",
        )
        self.assertIsNone(entries[0].weekday)
        self.assertEqual("unknown", entries[0].conflict_status)

    def test_gateway_posts_verified_request_and_returns_structured_snapshot(self):
        html = """
        <span>2026秋季学期 学生个人课表</span>
        <table>
          <tr><th>节次</th><th>星期一</th><th>星期二</th><th>星期三</th>
              <th>星期四</th><th>星期五</th><th>星期六</th><th>星期日</th></tr>
          <tr><td>上午</td><td>1-2</td><td>程序设计<br>张老师[1-16]周A101</td>
              <td></td><td></td><td></td><td></td><td></td><td></td></tr>
        </table>
        """

        class Page:
            url = "https://webvpn.example/http/abc123/home"

            def __init__(self):
                self.evaluate_calls = []

            def evaluate(self, script, arguments):
                self.evaluate_calls.append((script, arguments))
                return {
                    "status": 200,
                    "url": arguments["url"],
                    "requestBody": "fhlj=kbcx%2FqueryGrkb&xnxq=2026-20271",
                    "body": html,
                }

        class Session:
            def __init__(self, page):
                self.context = SimpleNamespace(pages=[page])
                self.open_calls = []

            def open_authenticated(self, url, *, timeout_seconds, page):
                self.open_calls.append((url, timeout_seconds, page))
                return page

        page = Page()
        session = Session(page)
        gateway = PlaywrightAcademicGateway.__new__(PlaywrightAcademicGateway)
        gateway._session = session

        result = gateway.observe_timetable(
            TimetableObservationRequest("2026年秋季学期", {"term": "2026年秋季学期"}),
            lambda *_: None,
            lambda: False,
        )

        self.assertEqual("complete", result.status)
        self.assertEqual("personal-timetable-api", result.source_kind)
        self.assertEqual(1, len(result.entries))
        endpoint = session.open_calls[0][0]
        self.assertEqual("/kbcx/queryGrkb", endpoint.split("/http/abc123", 1)[1])
        self.assertEqual({"url": endpoint}, page.evaluate_calls[0][1])
        self.assertTrue(result.trace.events[0].url.endswith("/kbcx/queryGrkb"))
        self.assertEqual(
            ("fhlj", "xnxq"), result.trace.events[1].field_names
        )

    def test_gateway_trace_records_retry_navigation_after_page_fetch_failure(self):
        from playwright.sync_api import Error as PlaywrightError

        html = """
        <span>2026秋季学期 学生个人课表</span>
        <table>
          <tr><th>节次</th><th>星期一</th><th>星期二</th><th>星期三</th>
              <th>星期四</th><th>星期五</th><th>星期六</th><th>星期日</th></tr>
          <tr><td>上午</td><td>1-2</td><td>程序设计<br>张老师[1-16]周A101</td>
              <td></td><td></td><td></td><td></td><td></td><td></td></tr>
        </table>
        """

        class Page:
            url = "https://webvpn.example/http/abc123/home"

            def __init__(self):
                self.attempts = 0

            def evaluate(self, _script, arguments):
                self.attempts += 1
                if self.attempts == 1:
                    raise PlaywrightError("query form unavailable after authentication")
                return {
                    "status": 200,
                    "url": arguments["url"],
                    "requestBody": "fhlj=kbcx%2FqueryGrkb&xnxq=2026-20271",
                    "body": html,
                }

        page = Page()
        session = SimpleNamespace(
            context=SimpleNamespace(pages=[page]),
            open_authenticated=lambda _url, **_kwargs: page,
        )
        gateway = PlaywrightAcademicGateway.__new__(PlaywrightAcademicGateway)
        gateway._session = session

        result = gateway.observe_timetable(
            TimetableObservationRequest("2026年秋季学期", {"term": "2026年秋季学期"}),
            lambda *_: None,
            lambda: False,
        )

        self.assertEqual("complete", result.status)
        self.assertEqual(["GET", "GET", "POST"], [event.method for event in result.trace.events])

    def test_gateway_rejects_timetable_for_a_different_requested_term(self):
        html = """
        <span>2026秋季学期 学生个人课表</span>
        <table>
          <tr><th>节次</th><th>星期一</th><th>星期二</th><th>星期三</th>
              <th>星期四</th><th>星期五</th><th>星期六</th><th>星期日</th></tr>
          <tr><td>上午</td><td>1-2</td><td>程序设计<br>张老师[1-16]周A101</td>
              <td></td><td></td><td></td><td></td><td></td><td></td></tr>
        </table>
        """

        class Page:
            url = "https://webvpn.example/http/abc123/home"

            def evaluate(self, _script, arguments):
                return {
                    "status": 200,
                    "url": arguments["url"],
                    "requestBody": "fhlj=kbcx%2FqueryGrkb&xnxq=2026-20271",
                    "body": html,
                }

        page = Page()
        session = SimpleNamespace(
            context=SimpleNamespace(pages=[page]),
            open_authenticated=lambda _url, **_kwargs: page,
        )
        gateway = PlaywrightAcademicGateway.__new__(PlaywrightAcademicGateway)
        gateway._session = session

        result = gateway.observe_timetable(
            TimetableObservationRequest("2026年春季学期", {"term": "2026年春季学期"}),
            lambda *_: None,
            lambda: False,
        )

        self.assertEqual("incomplete", result.status)
        self.assertIn("学期不匹配", result.error)


if __name__ == "__main__":
    unittest.main()
