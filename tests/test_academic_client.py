import unittest
from types import SimpleNamespace

from course_progress.academic_client import (
    AcademicContractError,
    AuthenticatedAcademicClient,
    resolve_academic_url,
)
from course_progress.collector import FixedGradeReader
from course_selection.deep_observation import ProgressObservationRequest
from course_selection.gateway import PlaywrightAcademicGateway


ENTRY_HTML = """
<select id="xnxqid"><option value="2025-20261">2025-2026 春季</option></select>
"""
GRADE_HTML = """
<input id="pageCount" value="1">
<table>
<tr><th>学年学期</th><th>课程代码</th><th>课程名称</th><th>课程性质</th>
<th>课程类别</th><th>学分</th><th>最终成绩</th></tr>
<tr><td>2025-20261</td><td>ART1</td><td>艺术史</td><td>选修</td>
<td>文理通识-文化素质教育课</td><td>2</td><td>通过</td></tr>
</table>
"""


class FakePage:
    def __init__(self, url="https://webvpn.example/http/abc123/kbcx/queryGrkb"):
        self.url = url
        self.evaluate_calls = []

    def content(self):
        return ENTRY_HTML

    def evaluate(self, _script, arguments):
        self.evaluate_calls.append(arguments)
        return {
            "status": 200,
            "url": arguments["url"],
            "requestBody": "pageXnxq=2025-20261&pageBkcxbj=&pageSfjg=&pageKcmc=",
            "body": GRADE_HTML,
        }


class FakeSession:
    def __init__(self, page):
        self.context = SimpleNamespace(pages=[page])
        self.opened = []

    def open_authenticated(self, url, *, timeout_seconds, page):
        self.opened.append((url, timeout_seconds))
        page.url = url
        return page


class AcademicClientTests(unittest.TestCase):
    def test_resolves_fixed_endpoint_inside_webvpn_prefix(self):
        self.assertEqual(
            "https://vpn.test/http/abc123/cjcx/queryQmcj",
            resolve_academic_url("https://vpn.test/http/abc123/kbcx/queryGrkb", "/cjcx/queryQmcj"),
        )

    def test_absolute_endpoint_is_used_verbatim_inside_webvpn_prefix(self):
        entry = "https://webvpn.hitwh.edu.cn/http/777/xsxk/queryXsxk?pageXklb=01"
        self.assertEqual(
            entry,
            resolve_academic_url("https://webvpn.hitwh.edu.cn/http/777/kbcx/queryGrkb", entry),
        )

    def test_fixed_grade_reader_uses_get_then_post_without_frames_or_clicks(self):
        page = FakePage()
        session = FakeSession(page)
        client = AuthenticatedAcademicClient(
            page,
            authenticate=lambda url, target: session.open_authenticated(
                url, timeout_seconds=15, page=target
            ),
        )
        collection = FixedGradeReader(client).collect()

        self.assertTrue(collection.complete)
        self.assertEqual("艺术史", collection.records[0].name)
        self.assertEqual(["GET", "POST"], [item["method"] for item in client.trace_requests])
        self.assertEqual(1, len(page.evaluate_calls))
        self.assertNotIn("frame", repr(page.evaluate_calls).lower())

    def test_fixed_grade_reader_rejects_missing_contract_marker(self):
        page = FakePage()
        page.content = lambda: "<html>changed</html>"
        session = FakeSession(page)
        client = AuthenticatedAcademicClient(
            page,
            authenticate=lambda url, target: session.open_authenticated(
                url, timeout_seconds=15, page=target
            ),
        )
        with self.assertRaisesRegex(AcademicContractError, "#xnxqid"):
            FixedGradeReader(client).collect()

    def test_progress_gateway_reports_contract_change_as_incomplete(self):
        page = FakePage()
        page.content = lambda: "<html>changed</html>"
        session = FakeSession(page)
        gateway = PlaywrightAcademicGateway.__new__(PlaywrightAcademicGateway)
        gateway._session = session
        gateway._academic_page = page
        result = gateway.observe_progress(
            ProgressObservationRequest({"term": "2026-1", "baseline_version": "guide-2026"}),
            lambda state, details: None,
            lambda: False,
        )
        self.assertEqual("incomplete", result.status)
        self.assertIn("#xnxqid", result.error)

    def test_progress_gateway_starts_from_direct_top_level_timetable_page(self):
        page = FakePage()
        session = FakeSession(page)
        gateway = PlaywrightAcademicGateway.__new__(PlaywrightAcademicGateway)
        gateway._session = session
        gateway._academic_page = page
        updates = []

        result = gateway.observe_progress(
            ProgressObservationRequest({"term": "2026-1", "baseline_version": "guide-2026"}),
            lambda state, details: updates.append((state, details)),
            lambda: False,
        )

        self.assertEqual("complete", result.status)
        self.assertTrue(session.opened[0][0].endswith("/cjcx/queryQmcj"))
        self.assertEqual(2, result.payload["report"]["progress"][2]["completed_credits"])
        self.assertEqual("2025-2026 春季", updates[0][1]["semester"])
        self.assertEqual(1, updates[0][1]["records"])


if __name__ == "__main__":
    unittest.main()
