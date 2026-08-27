import tempfile
import unittest
from pathlib import Path
from types import SimpleNamespace
from unittest.mock import patch

from course_selection.discovery import DiscoveryReport
from course_selection.gateway import PlaywrightAcademicGateway


class _Context:
    def on(self, *_args):
        return None

    def route(self, *_args):
        return None


class _Form:
    def __init__(self, count=1):
        self._count = count
        self.first = self

    def count(self):
        return self._count


class _Frame:
    url = "https://academic.example/xsxk/queryXsxk?pageXklb=szhx"

    def __init__(self, html):
        self.html = html
        self.calls = []

    def locator(self, *_args):
        return _Form()

    def evaluate(self, _script, arguments):
        self.calls.append(arguments)
        return {
            "status": 200,
            "url": "https://academic.example/xsxk/queryXsxkList",
            "requestBody": f"pageXklb={arguments['category']}&pageXnxq=2026-1",
            "body": self.html,
        }


class _Page:
    def __init__(self, html):
        self.main_frame = _Frame(html)
        self.urls = []

    def goto(self, url, **_kwargs):
        self.urls.append(url)

    def is_closed(self):
        return False


class SelectionSnapshotTests(unittest.TestCase):
    COURSE_HTML = """
    <span>选课时间：2026-08-29 08:30 至 2026-08-29 10:30</span>
    <table class="bot_line">
      <tr><th></th><th>序号</th><th>课程代码</th><th>课程名称</th>
      <th>前置课程</th><th>面向对象</th><th>校区</th><th>上课信息</th>
      <th>课程类别</th><th>开课院系</th><th>学分</th><th>学时</th>
      <th>备注信息</th><th>选课要求</th><th>已选/容量</th></tr>
      <tr><td><a onclick="saveXsxk1('TASK-9')">选择</a></td><td>1</td>
      <td>GE101</td><td>人工智能导论</td><td>无</td><td>全校本科生</td>
      <td>威海校区</td><td>教师：李老师 周一 1-2节</td><td>文化素质</td>
      <td>计算机学院</td><td>2</td><td>32</td><td></td><td></td><td>18/30</td></tr>
    </table>
    """

    def test_verified_adapter_authenticates_direct_entry_without_running_discovery(self):
        page = _Page(self.COURSE_HTML)

        class Session:
            context = SimpleNamespace(pages=[page])

            def __init__(self):
                self.authenticated_urls = []

            def open_authenticated(self, url, *, timeout_seconds, page):
                self.authenticated_urls.append((url, timeout_seconds))
                return page

        gateway = PlaywrightAcademicGateway()
        gateway._session = Session()
        gateway._academic_page = page

        with patch(
            "course_selection.discovery.InterfaceDiscovery",
            side_effect=AssertionError("discovery must not run for a valid contract"),
        ):
            result = gateway.refresh_selection(
                {
                    "allowed_categories": ("szhx",),
                    "allowed_windows": {},
                    "semester_label": "2026-1",
                },
                lambda *_args: None,
                lambda: False,
            )

        self.assertEqual("complete", result["status"])
        self.assertEqual("verified-selection-api", result["source_kind"])
        self.assertEqual("hitwh-jwts-selection-query-v1", result["contract_version"])
        self.assertEqual("TASK-9", result["sections"][0]["identity"])
        self.assertEqual([], page.urls)
        self.assertIn("pageXklb=szhx", gateway._session.authenticated_urls[0][0])
        self.assertEqual(600, gateway._session.authenticated_urls[0][1])
        self.assertEqual(1, len(page.main_frame.calls))

    def test_gateway_consumes_in_memory_payload_without_reading_diagnostic_json(self):
        payload = {
            "queries": [
                {
                    "category": "szhx",
                    "complete": True,
                    "sections": [{"identity": "section-1", "name": "Test"}],
                }
            ]
        }
        report = DiscoveryReport(
            target="selection",
            target_found=True,
            clicks=0,
            captures=0,
            blocked_requests=0,
            candidates_path=Path("candidates.json"),
            click_log_path=Path("clicks.json"),
            # Deliberately point at a file that does not exist.  The refresh
            # result must still be produced from the structured payload.
            selection_query_path=Path("missing-selection-query.json"),
            selection_query_payload=payload,
        )

        class Navigator:
            def __init__(self, **_kwargs):
                pass

            def _handle_response(self, *_args):
                return None

            def _guard_route(self, *_args):
                return None

            def run(self, *_args, **_kwargs):
                return report

        with tempfile.TemporaryDirectory() as directory:
            gateway = PlaywrightAcademicGateway(workspace_root=directory)
            gateway._session = SimpleNamespace(context=_Context())
            with patch("course_selection.discovery.InterfaceDiscovery", Navigator):
                result = gateway.refresh_selection(
                    {
                        "allowed_categories": ("szhx",),
                        "allowed_windows": {},
                        "semester_label": "2026-1",
                    },
                    lambda *_args: None,
                    lambda: False,
                )

        self.assertEqual("complete", result["status"])
        self.assertEqual("section-1", result["sections"][0]["identity"])


if __name__ == "__main__":
    unittest.main()
