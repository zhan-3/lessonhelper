import unittest
from types import SimpleNamespace

from course_selection.gateway import PlaywrightAcademicGateway
from course_selection.personal_timetable import parse_personal_timetable_html, personal_timetable_parameters


class PersonalTimetableParserTests(unittest.TestCase):
    def test_request_body_contains_only_verified_fields(self):
        self.assertEqual(
            {"fhlj": "current", "xnxq": "2026-1"},
            personal_timetable_parameters(fhlj=" current ", xnxq="2026-1"),
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
        <table><tr><th>课程代码</th><th>课程名称</th><th>星期</th><th>节次</th><th>周次</th></tr>
        <tr><td>CS101</td><td>程序设计</td><td>星期一</td><td>1-2</td><td>1-16</td></tr></table>
        """

        class Locator:
            def __init__(self, value):
                self.value = value

            @property
            def first(self):
                return self

            def count(self):
                return 1 if self.value else 0

            def input_value(self):
                return self.value

        class Page:
            url = "https://webvpn.example/http/abc123/kbcx/queryGrkb"

            def locator(self, selector):
                return Locator({"fhlj": "current", "xnxq": "2026-1"}.get(selector.split("'")[1], ""))

        class Request:
            def __init__(self):
                self.calls = []

            def post(self, endpoint, *, form, timeout):
                self.calls.append((endpoint, form, timeout))
                return SimpleNamespace(status=200, text=lambda: html)

        request = Request()
        context = SimpleNamespace(pages=[Page()], request=request)
        gateway = PlaywrightAcademicGateway.__new__(PlaywrightAcademicGateway)
        gateway._session = SimpleNamespace(context=context)

        result = gateway.refresh_timetable({"term": "2026-1"}, lambda *_: None, lambda: False)

        self.assertEqual("complete", result["status"])
        self.assertEqual("personal-timetable-api", result["source_kind"])
        self.assertEqual(1, len(result["entries"]))
        self.assertEqual("/kbcx/queryXszkb", request.calls[0][0].split("/http/abc123", 1)[1])
        self.assertEqual({"fhlj": "current", "xnxq": "2026-1"}, request.calls[0][1])


if __name__ == "__main__":
    unittest.main()
