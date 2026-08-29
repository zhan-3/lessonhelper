import tempfile
import unittest
from pathlib import Path
from types import SimpleNamespace
from unittest.mock import patch

from course_selection.discovery import DiscoveryReport
from course_selection.deep_observation import SelectionObservationRequest
from course_selection.gateway import PlaywrightAcademicGateway
from course_selection.selection_query import VerifiedSelectionQueryAdapter


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

    def test_gateway_recognizes_only_the_verified_academic_proxy(self):
        self.assertTrue(
            PlaywrightAcademicGateway._is_academic_application(
                "https://webvpn.hitwh.edu.cn/http/"
                "77726476706e69737468656265737421fae0558f693861446900c7a99c406d3667/"
                "xsxk/queryXsxk"
            )
        )
        self.assertFalse(
            PlaywrightAcademicGateway._is_academic_application(
                "https://webvpn.hitwh.edu.cn/http/unrelated-application/"
            )
        )

    def test_gateway_connects_through_rendered_webvpn_application(self):
        class Page:
            url = "https://webvpn.hitwh.edu.cn/"

            def is_closed(self):
                return False

            def bring_to_front(self):
                return None

        portal = Page()
        academic = Page()
        academic.url = "https://webvpn.hitwh.edu.cn/http/academic/home"

        class Session:
            context = SimpleNamespace(pages=[portal])

            def __init__(self):
                self.calls = []

            def open_portal_application(self, portal_url, name, **kwargs):
                self.calls.append((portal_url, name, kwargs["page"]))
                return academic

        gateway = PlaywrightAcademicGateway()
        gateway._session = Session()
        gateway.connect(lambda *_args: None, lambda: False)

        self.assertIs(academic, gateway._academic_page)
        self.assertEqual("新教务系统", gateway._session.calls[0][1])
        self.assertIs(portal, gateway._session.calls[0][2])

    def test_gateway_continues_from_academic_landing_through_cas_login(self):
        landing_url = (
            "https://webvpn.hitwh.edu.cn/http/"
            "77726476706e69737468656265737421fae0558f693861446900c7a99c406d3667/"
            "?wrdrecordvisit=1787895847000"
        )

        class Page:
            url = landing_url

            def __init__(self):
                self.navigations = []

            def goto(self, url, **_kwargs):
                self.navigations.append(url)
                self.url = url

            def is_closed(self):
                return False

            def bring_to_front(self):
                return None

        page = Page()

        class Session:
            context = SimpleNamespace(pages=[page])

            def __init__(self):
                self.urls = []

            def open_authenticated(self, url, *, timeout_seconds, page):
                self.urls.append(url)
                page.url = url
                return page

        gateway = PlaywrightAcademicGateway()
        gateway._session = Session()
        gateway._academic_page = page
        gateway.connect(lambda *_args: None, lambda: False)

        self.assertEqual(
            "https://webvpn.hitwh.edu.cn/http/"
            "77726476706e69737468656265737421fae0558f693861446900c7a99c406d3667/loginCAS",
            page.navigations[0],
        )
        self.assertEqual(
            "https://webvpn.hitwh.edu.cn/http/"
            "77726476706e69737468656265737421fae0558f693861446900c7a99c406d3667/"
            "xsxk/queryXsxk",
            gateway._session.urls[0],
        )

    def test_gateway_maps_webvpn_session_expiry_to_waiting_for_authentication(self):
        from course_progress.session import WebVpnSessionExpiredError

        class Page:
            url = "https://webvpn.hitwh.edu.cn/"

            def is_closed(self):
                return False

            def bring_to_front(self):
                return None

        page = Page()

        class Session:
            context = SimpleNamespace(pages=[page])

            def open_portal_application(self, portal_url, name, **kwargs):
                raise WebVpnSessionExpiredError("会话已失效（网络/IP 变化）")

        events = []
        gateway = PlaywrightAcademicGateway()
        gateway._session = Session()
        with self.assertRaises(WebVpnSessionExpiredError):
            gateway.connect(
                lambda state, details: events.append((state, details)), lambda: False
            )

        self.assertEqual("waiting_for_authentication", events[-1][0])
        self.assertIn("网络/IP", events[-1][1]["message"])
        self.assertEqual(page, gateway._academic_page)

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
            observed = gateway.observe_selection(
                SelectionObservationRequest({
                    "allowed_categories": ("szhx",),
                    "allowed_windows": {},
                    "semester_label": "2026-1",
                }),
                lambda *_args: None,
                lambda: False,
            )

        self.assertEqual("complete", observed.status)
        self.assertEqual("verified-selection-api", observed.payload["source_kind"])
        self.assertEqual("hitwh-jwts-selection-query-v1", observed.payload["contract_version"])
        section = observed.payload["sections"][0]
        self.assertEqual("TASK-9", section["identity"])
        self.assertEqual("TASK-9", section["action_rwh"])
        self.assertTrue(section["execution_ready"])
        self.assertEqual("szhx", section["query_code"])
        self.assertEqual("2026-1", section["query_term"])
        self.assertEqual(1, section["query_page"])
        self.assertEqual([], page.urls)
        self.assertIn("pageXklb=szhx", gateway._session.authenticated_urls[0][0])
        self.assertEqual(600, gateway._session.authenticated_urls[0][1])
        self.assertEqual(1, len(page.main_frame.calls))

    def test_verified_adapter_reads_every_whitelisted_category_and_page(self):
        html = self.COURSE_HTML.replace(
            '<table class="bot_line">',
            '<input name="pageCount" value="2"><table class="bot_line">',
        )
        page = _Page(html)
        result = VerifiedSelectionQueryAdapter().read(
            page,
            categories=("szhx", "xsxk"),
            semester_label="2026-1",
            allowed_windows={},
            progress=lambda *_args: None,
            cancelled=lambda: False,
        )
        self.assertEqual("complete", result["status"])
        self.assertEqual(4, len(page.main_frame.calls))
        self.assertEqual([2, 2], [query["pages_fetched"] for query in result["queries"]])
        self.assertEqual("szhx", result["queries"][0]["sections"][0]["query_code"])
        self.assertEqual("xsxk", result["queries"][1]["sections"][0]["query_code"])

    def test_gateway_keeps_discovered_payload_diagnostic_only(self):
        report = DiscoveryReport(
            target="selection",
            target_found=True,
            clicks=0,
            captures=0,
            blocked_requests=0,
            candidates_path=Path("candidates.json"),
            click_log_path=Path("clicks.json"),
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
                observed = gateway.observe_selection(
                    SelectionObservationRequest({
                        "allowed_categories": ("szhx",),
                        "allowed_windows": {},
                        "semester_label": "2026-1",
                    }),
                    lambda *_args: None,
                    lambda: False,
                )

        self.assertEqual("interface_unconfirmed", observed.status)
        self.assertEqual(0, observed.diagnostic["captures"])


if __name__ == "__main__":
    unittest.main()
