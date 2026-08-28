import json
import tempfile
import unittest
from pathlib import Path

from course_progress.capture import CaptureStore, score_candidate
from course_progress.cli import build_parser
from course_progress.credentials import CredentialStore, LoginCredentials
from course_progress.explorer import (
    DEFAULT_PORTAL_URL,
    _is_relevant_control,
    _is_safe_navigation,
    _is_safe_portal_fallback,
    resolve_profile_dir,
)
from course_progress.sanitizer import (
    REDACTED,
    sanitize_data,
    sanitize_request_body,
    sanitize_url,
)
from course_progress.session import (
    AcademicBrowserSession,
    WebVpnSessionExpiredError,
    _is_legacy_webvpn_login,
    _is_webvpn_credential_page,
)


class AcademicBrowserSessionTests(unittest.TestCase):
    def test_authentication_tracks_the_requested_academic_tab_not_the_workbench_tab(self):
        class Page:
            def __init__(self, url):
                self.url = url

            def goto(self, *_args, **_kwargs):
                return None

            def wait_for_timeout(self, *_args):
                return None

        class Context:
            def __init__(self, pages):
                self.pages = pages

            def storage_state(self, **_kwargs):
                return None

            def clear_cookies(self):
                return None

        target = Page(
            "https://webvpn.hitwh.edu.cn/https/hash/authserver/login?service=academic"
        )
        shell = Page("http://127.0.0.1:5000/")
        with tempfile.TemporaryDirectory() as directory:
            session = AcademicBrowserSession.__new__(AcademicBrowserSession)
            session.context = Context([target, shell])
            session.private_root = Path(directory)
            session.auth_state_path = Path(directory) / "state.json"
            attempts = []

            def login(page):
                attempts.append(page.url)
                page.url = "https://webvpn.hitwh.edu.cn/http/academic/xsxk/queryXsxk"
                return True

            session._fill_and_submit_login = login
            result = session.open_authenticated(
                "https://webvpn.hitwh.edu.cn/http/academic/xsxk/queryXsxk",
                timeout_seconds=1,
                page=target,
            )

        self.assertIs(target, result)
        self.assertEqual(1, len(attempts))

    def test_authentication_retries_when_first_submitted_credentials_are_ignored(self):
        class Page:
            url = "https://webvpn.hitwh.edu.cn/https/hash/authserver/login?service=academic"

            def goto(self, *_args, **_kwargs):
                return None

            def wait_for_timeout(self, *_args):
                return None

        class Context:
            pages = []

            def storage_state(self, **_kwargs):
                return None

        with tempfile.TemporaryDirectory() as directory:
            session = AcademicBrowserSession.__new__(AcademicBrowserSession)
            session.context = Context()
            session.private_root = Path(directory)
            session.auth_state_path = Path(directory) / "state.json"
            attempts = []

            def login(page):
                attempts.append(True)
                if len(attempts) == 2:
                    page.url = "https://webvpn.hitwh.edu.cn/http/academic/home"
                return True

            session._fill_and_submit_login = login
            result = session.open_authenticated(
                "https://webvpn.hitwh.edu.cn/http/academic/home",
                timeout_seconds=4,
                page=Page(),
            )

        self.assertEqual(2, len(attempts))
        self.assertIn("/http/academic/", result.url)

    def test_webvpn_health_check_rejects_ip_change_redirect(self):
        class Context:
            request = type(
                "Request", (),
                {"get": lambda *_args, **_kwargs: type("Response", (), {"status": 302, "url": "https://webvpn.hitwh.edu.cn/login?logoutByIpChange=true"})()}
            )()

        session = AcademicBrowserSession.__new__(AcademicBrowserSession)
        session.context = Context()
        with self.assertRaisesRegex(WebVpnSessionExpiredError, "网络/IP"):
            session.assert_webvpn_session()

    def test_cdp_attached_session_detaches_without_closing_browser_context(self):
        class Context:
            closed = False

            def close(self):
                self.closed = True

        class Browser:
            def __init__(self):
                self.contexts = [Context()]

        class Chromium:
            def connect_over_cdp(self, url):
                self.url = url
                return Browser()

        session = AcademicBrowserSession.__new__(AcademicBrowserSession)
        session.playwright = type("Playwright", (), {"chromium": Chromium()})()
        session.cdp_url = "http://127.0.0.1:9222"
        session._attached_over_cdp = False
        session.context = None
        session.browser = None
        session.__enter__()
        context = session.context
        session.__exit__(None, None, None)

        self.assertFalse(context.closed)
        self.assertIsNone(session.context)

    def test_portal_application_uses_rendered_resource_and_new_page(self):
        class TargetPage:
            url = "about:blank"

            def wait_for_timeout(self, *_args):
                return None

        target = TargetPage()

        class Resource:
            def is_visible(self):
                return True

            def click(self, **_kwargs):
                target.url = "https://webvpn.hitwh.edu.cn/http/academic/"
                context.page_handler(target)

        class Locator:
            def count(self):
                return 1

            def nth(self, _index):
                return Resource()

        class PortalPage:
            url = "https://webvpn.hitwh.edu.cn/"

            def get_by_text(self, text, *, exact):
                self.requested = (text, exact)
                return Locator()

            def wait_for_timeout(self, *_args):
                return None

        class Context:
            page_handler = None
            request = type(
                "Request", (),
                {"get": lambda *_args, **_kwargs: type("Response", (), {"status": 200, "url": "https://webvpn.hitwh.edu.cn/user/info"})()}
            )()

            def on(self, event, handler):
                self.page_handler = handler

            def remove_listener(self, event, handler):
                if self.page_handler is handler:
                    self.page_handler = None

        portal = PortalPage()
        context = Context()
        session = AcademicBrowserSession.__new__(AcademicBrowserSession)
        session.context = context
        opened = []

        def authenticate(url, *, timeout_seconds, page=None):
            opened.append((url, page))
            return page or portal

        session.open_authenticated = authenticate
        result = session.open_portal_application(
            "https://webvpn.hitwh.edu.cn/", "新教务系统", timeout_seconds=1
        )

        self.assertIs(target, result)
        self.assertEqual(("新教务系统", True), portal.requested)
        self.assertEqual(
            "https://webvpn.hitwh.edu.cn/http/academic/", opened[-1][0]
        )


class SanitizerTests(unittest.TestCase):
    def test_redacts_auth_and_personal_fields_but_keeps_course_fields(self):
        source = {
            "accessToken": "eyJabc.def.ghi",
            "studentName": "测试同学",
            "studentNo": "2025000000",
            "courseName": "高等数学",
            "credits": 5.0,
        }

        result = sanitize_data(source)

        self.assertEqual(result["accessToken"], REDACTED)
        self.assertEqual(result["studentName"], REDACTED)
        self.assertEqual(result["studentNo"], REDACTED)
        self.assertEqual(result["courseName"], "高等数学")
        self.assertEqual(result["credits"], 5.0)

    def test_redacts_sensitive_url_parameters(self):
        result = sanitize_url(
            "https://example.test/api?ticket=ST-secret&semester=2025-1"
        )

        self.assertIn("ticket=%5BREDACTED%5D", result)
        self.assertIn("semester=2025-1", result)
        self.assertNotIn("ST-secret", result)

    def test_redacts_legacy_academic_user_query_parameters(self):
        result = sanitize_url(
            "https://example.test/xsxk/query?yhxx=student-detail&yhid=2025000000"
        )

        self.assertNotIn("student-detail", result)
        self.assertNotIn("2025000000", result)
        self.assertEqual(result.count("%5BREDACTED%5D"), 2)

    def test_redacts_form_encoded_login_body(self):
        result = sanitize_request_body(
            "username=2025000000&password=do-not-store&execution=e1s1"
        )

        self.assertNotIn("2025000000", result)
        self.assertNotIn("do-not-store", result)
        self.assertIn("execution=e1s1", result)

    def test_redacts_generic_name_inside_user_object(self):
        result = sanitize_data(
            {
                "user": {"studentNo": "2025000000", "name": "测试同学"},
                "course": {"code": "MATH101", "name": "高等数学"},
            }
        )

        self.assertEqual(result["user"]["name"], REDACTED)
        self.assertEqual(result["course"]["name"], "高等数学")


class CredentialTests(unittest.TestCase):
    def test_encrypted_store_round_trips_without_writing_plaintext(self):
        with tempfile.TemporaryDirectory() as temp_dir:
            path = Path(temp_dir) / "login.dpapi"
            store = CredentialStore(
                path,
                protect=lambda value: b"encrypted:" + value[::-1],
                unprotect=lambda value: value.removeprefix(b"encrypted:")[::-1],
            )
            expected = LoginCredentials("2025000000", "do-not-store-plainly")

            store.save(expected)

            self.assertEqual(store.load(), expected)
            self.assertNotIn(b"do-not-store-plainly", path.read_bytes())

    def test_credentials_are_only_entered_on_exact_webvpn_cas_host(self):
        self.assertTrue(
            _is_webvpn_credential_page(
                "https://webvpn.hitwh.edu.cn/https/hash/authserver/login?service=x"
            )
        )
        self.assertFalse(
            _is_webvpn_credential_page(
                "https://evil.example/authserver/login?next=webvpn.hitwh.edu.cn"
            )
        )
        self.assertFalse(
            _is_webvpn_credential_page("http://webvpn.hitwh.edu.cn/authserver/login")
        )

    def test_legacy_easyconnect_login_is_detected_without_accepting_other_hosts(self):
        self.assertTrue(
            _is_legacy_webvpn_login(
                "https://webvpn.hitwh.edu.cn/https/hash/portal/#!/login"
            )
        )
        self.assertFalse(
            _is_legacy_webvpn_login("https://evil.example/portal/#!/login")
        )


class CandidateTests(unittest.TestCase):
    def test_graduation_progress_response_scores_above_generic_response(self):
        relevant_score, _ = score_candidate(
            "https://example.test/api/training/program",
            {"课程类别": "文化素质", "要求学分": 8},
        )
        generic_score, _ = score_candidate(
            "https://example.test/api/notice", {"message": "ok"}
        )

        self.assertGreaterEqual(relevant_score, 10)
        self.assertGreater(relevant_score, generic_score)

    def test_capture_store_writes_raw_sanitized_index_and_ranked_candidates(self):
        with tempfile.TemporaryDirectory() as temp_dir:
            store = CaptureStore(Path(temp_dir))
            candidate = store.save_json_exchange(
                url="https://example.test/api/course?token=secret",
                method="POST",
                status=200,
                content_type="application/json",
                request_body='{"studentNo":"2025000000"}',
                response_data={
                    "studentName": "测试同学",
                    "课程类别": "专业选修",
                    "学分": 3,
                },
            )
            candidates_path = store.write_candidates()

            raw = json.loads((Path(temp_dir) / "raw/000001.json").read_text("utf-8"))
            sanitized = json.loads(
                (Path(temp_dir) / "sanitized/000001.json").read_text("utf-8")
            )
            index = json.loads(
                (Path(temp_dir) / "index.jsonl").read_text("utf-8").strip()
            )
            candidates = json.loads(candidates_path.read_text("utf-8"))

            self.assertEqual(raw["studentName"], "测试同学")
            self.assertEqual(sanitized["studentName"], REDACTED)
            self.assertNotIn("secret", index["url"])
            self.assertNotIn("2025000000", index["request_body"])
            self.assertEqual(candidates[0]["capture_id"], candidate.capture_id)


class AutoNavigationTests(unittest.TestCase):
    def test_default_entry_uses_plain_portal_root(self):
        self.assertEqual(DEFAULT_PORTAL_URL, "https://webvpn.hitwh.edu.cn/")

    def test_official_webvpn_authentication_posts_are_not_treated_as_course_mutations(self):
        from course_selection.discovery import is_mutating_request

        login_body = "username=student&password=redacted&execution=e1s1&submit=login"

        self.assertFalse(
            is_mutating_request(
                "POST",
                "https://webvpn.hitwh.edu.cn/login?cas_login=true",
                login_body,
            )
        )
        self.assertFalse(
            is_mutating_request(
                "POST",
                "https://webvpn.hitwh.edu.cn/https/hash/authserver/login?service=academic",
                login_body,
            )
        )
        self.assertTrue(
            is_mutating_request(
                "POST",
                "https://webvpn.hitwh.edu.cn/api/resource/open",
                "action=query&submit=confirm",
            )
        )
        self.assertTrue(
            is_mutating_request(
                "POST",
                "http://webvpn.hitwh.edu.cn/login",
                login_body,
            )
        )
        self.assertTrue(
            is_mutating_request(
                "POST",
                "https://evil.example/authserver/login?next=webvpn.hitwh.edu.cn",
                login_body,
            )
        )

    def test_profile_ignores_profiles_from_other_browser_channels(self):
        with tempfile.TemporaryDirectory() as temp_dir:
            private_root = Path(temp_dir)
            shared = private_root / "playwright-chromium-profile"
            shared.mkdir()
            (private_root / "collector-profile").mkdir()

            self.assertEqual(
                resolve_profile_dir(private_root),
                shared,
            )

    def test_profile_does_not_reuse_legacy_collector_or_explorer(self):
        with tempfile.TemporaryDirectory() as temp_dir:
            private_root = Path(temp_dir)
            (private_root / "collector-profile").mkdir()
            (private_root / "explorer-profile").mkdir()

            self.assertEqual(
                resolve_profile_dir(private_root),
                private_root / "playwright-chromium-profile",
            )

    def test_profile_does_not_reuse_legacy_explorer(self):
        with tempfile.TemporaryDirectory() as temp_dir:
            private_root = Path(temp_dir)
            explorer = private_root / "explorer-profile"
            explorer.mkdir()

            self.assertEqual(
                resolve_profile_dir(private_root),
                private_root / "playwright-chromium-profile",
            )

    def test_profile_defaults_to_new_shared_directory(self):
        with tempfile.TemporaryDirectory() as temp_dir:
            private_root = Path(temp_dir)

            self.assertEqual(
                resolve_profile_dir(private_root),
                private_root / "playwright-chromium-profile",
            )

    def test_explore_defaults_to_automatic_navigation_without_enter_prompt(self):
        args = build_parser().parse_args(["explore"])

        self.assertEqual(args.max_pages, 12)
        self.assertEqual(args.login_timeout_seconds, 600)

    def test_collect_defaults_to_dynamic_semester_collection(self):
        args = build_parser().parse_args(["collect"])

        self.assertEqual(args.page_size, 20)
        self.assertEqual(args.login_timeout_seconds, 600)
        self.assertEqual(
            args.requirements,
            Path("docs/校园培养方案解读（2026年版）.md"),
        )

    def test_only_same_origin_http_navigation_is_allowed(self):
        self.assertTrue(
            _is_safe_navigation(
                "https://portal.test/home", "https://portal.test/course/plan"
            )
        )
        self.assertFalse(
            _is_safe_navigation(
                "https://portal.test/home", "https://other.test/course/plan"
            )
        )
        self.assertFalse(
            _is_safe_navigation(
                "https://portal.test/home", "javascript:submitCourse()"
            )
        )

    def test_read_only_course_entry_is_allowed_but_registration_is_blocked(self):
        self.assertTrue(_is_relevant_control("培养方案"))
        self.assertTrue(_is_relevant_control("已修课程"))
        self.assertFalse(_is_relevant_control("选课中心"))
        self.assertFalse(_is_relevant_control("提交课程"))
        self.assertTrue(
            _is_safe_portal_fallback("本科生综合服务", "/portal/student/academic")
        )
        self.assertFalse(_is_safe_portal_fallback("账户设置", "/user/settings"))


if __name__ == "__main__":
    unittest.main()
