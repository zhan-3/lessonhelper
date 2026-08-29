import json
import tempfile
import threading
import time
import unittest
from pathlib import Path
from types import SimpleNamespace
from unittest.mock import patch

from playwright.sync_api import Error

from course_selection.deep_observation import (
    AcademicRequestTrace,
    ManualObservationResult,
    ProgressObservationResult,
    SelectionObservationResult,
    TimetableObservationResult,
)
from course_selection.gateway import PlaywrightAcademicGateway
from course_selection.notice_discovery import (
    OfficialNoticeLink,
    candidate_from_text,
    discover_official_notice_candidates,
    parse_official_notice_article,
    parse_official_notice_links,
)
from course_selection.persistence import WorkspaceDatabase
from course_selection.tasks import ObservationService, TASK_TIMEOUT_SECONDS, TaskState
from course_selection.workbench import create_workbench_app


class FakeGateway:
    def __init__(self):
        self.connect_count = 0
        self.selection_count = 0
        self.closed = False

    def launch_shell(self, url, progress, cancelled):
        progress("connecting", {"message": "shell opened"})

    def connect(self, progress, cancelled):
        self.connect_count += 1
        progress("connecting", {})

    def refresh_selection(self, context, progress, cancelled):
        self.selection_count += 1
        progress("reading", {"category": "szhx", "page": 1})
        return {"status": "complete", "sections": [{"name": "艺术鉴赏"}]}

    def refresh_timetable(self, context, progress, cancelled):
        return {"status": "interface_unconfirmed", "entries": []}

    def observe_selection(self, request, progress, cancelled):
        result = self.refresh_selection(request.context, progress, cancelled)
        return SelectionObservationResult.complete(result, trace=AcademicRequestTrace.empty())

    def observe_timetable(self, request, progress, cancelled):
        result = self.refresh_timetable(request.context, progress, cancelled)
        if result.get("status") != "complete":
            return TimetableObservationResult.incomplete(result.get("status", "incomplete"))
        return TimetableObservationResult.complete(
            term=str(request.context.get("term", "")), entries=result.get("entries", []),
            trace=AcademicRequestTrace.empty(),
        )

    def observe_progress(self, request, progress, cancelled):
        result = self.refresh_progress(request.context, progress, cancelled)
        return ProgressObservationResult.complete(result, trace=AcademicRequestTrace.empty())

    def observe_manual(self, request, progress, cancelled, finished):
        result = self.observe_navigation(request.context, progress, cancelled, finished)
        return ManualObservationResult(status=result["status"], diagnostic=result["report"])

    def close(self):
        self.closed = True


class ProgressGateway(FakeGateway):
    def refresh_progress(self, context, progress, cancelled):
        progress("reading", {"target": "progress", "semester": "2026春", "page": 1, "page_count": 1})
        return {
            "status": "complete",
            "source_kind": "academic",
            "term": "2026年春季学期",
            "report": {
                "data_complete": True,
                "baseline_version": "guide-2026",
                "progress": [{
                    "key": "major_elective",
                    "label": "本专业选修",
                    "required_credits": 12,
                    "completed_credits": 3,
                    "remaining_credits": 9,
                    "courses": [{"code": "CS1", "name": "课程", "credits": 3}],
                }],
            },
        }


class ResetFailureGateway(FakeGateway):
    def reset_login(self):
        raise OSError("profile remains locked")


class ShellGateway(FakeGateway):
    def __init__(self):
        super().__init__()
        self.shell_urls = []
        self.timetable_count = 0

    def launch_shell(self, url, progress, cancelled):
        self.shell_urls.append(url)
        progress("connecting", {"message": "visible shell opened"})

    def refresh_timetable(self, context, progress, cancelled):
        self.timetable_count += 1
        return {"status": "complete", "entries": [{"course_name": "Current"}]}


class ManualObservationGateway(FakeGateway):
    def __init__(self):
        super().__init__()
        self.observing = threading.Event()

    def observe_navigation(self, context, progress, cancelled, finished):
        progress("observing", {"message": "manual navigation observation active"})
        self.observing.set()
        while not cancelled() and not finished():
            time.sleep(0.01)
        return {
            "status": "complete",
            "report": {
                "browser_instances": 1,
                "events": [{"method": "GET", "url": "https://example.test/academic"}],
                "blocked_requests": [],
            },
        }


class AuthenticationGateway(FakeGateway):
    def __init__(self):
        super().__init__()
        self.authentication_complete = threading.Event()

    def connect(self, progress, cancelled):
        self.connect_count += 1
        progress("waiting_for_authentication", {"message": "sign in in the browser"})
        self.authentication_complete.wait(2)
        if not self.authentication_complete.is_set():
            raise TimeoutError("authentication timed out")


class PlaywrightGatewayRouteTests(unittest.TestCase):
    def test_selection_refresh_removes_temporary_request_and_response_handlers(self):
        class FakeContext:
            def __init__(self):
                self.routes = []
                self.unroutes = []
                self.listeners = []
                self.removed_listeners = []

            def route(self, pattern, handler):
                self.routes.append((pattern, handler))

            def unroute(self, pattern, handler):
                self.unroutes.append((pattern, handler))

            def on(self, event, handler):
                self.listeners.append((event, handler))

            def remove_listener(self, event, handler):
                self.removed_listeners.append((event, handler))

        class FakeNavigator:
            def __init__(self, **_kwargs):
                pass

            def _handle_response(self, _response):
                pass

            def _guard_route(self, _route):
                pass

            def run(self, _context, **_kwargs):
                return SimpleNamespace(selection_query_payload=None)

        context = FakeContext()
        gateway = PlaywrightAcademicGateway(Path("profile"), Path("workspace"))
        gateway._session = SimpleNamespace(context=context)

        with patch("course_selection.discovery.InterfaceDiscovery", FakeNavigator):
            result = gateway.refresh_selection(
                {
                    "allowed_categories": ["allowed"],
                    "allowed_windows": {},
                    "semester_label": "2026-2027-1",
                },
                lambda _state, _details: None,
                lambda: False,
            )

        self.assertEqual("interface_unconfirmed", result["status"])
        self.assertEqual(context.routes, context.unroutes)
        self.assertEqual(context.listeners, context.removed_listeners)


    def test_manual_diagnostic_removes_handlers_between_consecutive_tasks(self):
        class FakePage:
            url = "https://academic.test/home"

        class FakeContext:
            def __init__(self):
                self.pages = [FakePage()]
                self.routes = []
                self.unroutes = []
                self.listeners = []
                self.removed_listeners = []

            def route(self, pattern, handler):
                self.routes.append((pattern, handler))

            def unroute(self, pattern, handler):
                self.unroutes.append((pattern, handler))

            def on(self, event, handler):
                self.listeners.append((event, handler))

            def remove_listener(self, event, handler):
                self.removed_listeners.append((event, handler))

        from course_selection.deep_observation import ManualObservationRequest

        context = FakeContext()
        gateway = PlaywrightAcademicGateway(Path("profile"), Path("workspace"))
        gateway._session = SimpleNamespace(context=context)
        for _ in range(2):
            result = gateway.observe_manual(
                ManualObservationRequest({}),
                lambda _state, _details: None,
                lambda: False,
                lambda: True,
            )
            self.assertEqual("complete", result.status)

        self.assertEqual(context.routes, context.unroutes)
        self.assertEqual(context.listeners, context.removed_listeners)
        self.assertEqual(2, len(context.routes))

        class FakeRoute:
            def __init__(self):
                self.request = SimpleNamespace(
                    method="POST",
                    url="https://academic.test/selection/save",
                    post_data="action=save",
                    headers={"content-type": "application/x-www-form-urlencoded"},
                    resource_type="other",
                )
                self.aborted = False

            def continue_(self):
                raise AssertionError("mutating non-XHR request must not continue")

            def abort(self, _reason):
                self.aborted = True

        route = FakeRoute()
        context.routes[0][1](route)
        self.assertTrue(route.aborted)


    def test_manual_diagnostic_reports_browser_close_from_playwright_wait(self):
        class ClosedPage:
            url = "https://academic.test/home"

            def wait_for_timeout(self, _milliseconds):
                raise Error("Target page has been closed")

        class FakeContext:
            def __init__(self):
                self.pages = [ClosedPage()]

            def route(self, _pattern, _handler):
                pass

            def unroute(self, _pattern, _handler):
                pass

            def on(self, _event, _handler):
                pass

            def remove_listener(self, _event, _handler):
                pass

        from course_selection.deep_observation import ManualObservationRequest

        gateway = PlaywrightAcademicGateway(Path("profile"), Path("workspace"))
        gateway._session = SimpleNamespace(context=FakeContext())
        result = gateway.observe_manual(
            ManualObservationRequest({"timeout_seconds": 30}),
            lambda _state, _details: None,
            lambda: False,
            lambda: False,
        )
        self.assertEqual("browser_closed", result.status)


class PollFailureGateway(FakeGateway):
    def __init__(self):
        super().__init__()
        self.polled = threading.Event()

    def poll(self):
        self.polled.set()
        raise RuntimeError("browser context failed")


class TimedOutAuthenticationGateway(FakeGateway):
    def connect(self, progress, cancelled):
        self.connect_count += 1
        progress("waiting_for_authentication", {"message": "sign in in the browser"})
        raise TimeoutError("authentication timed out")


class WorkspaceDatabaseTests(unittest.TestCase):
    def test_creates_database_and_imports_legacy_json_once(self):
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            (root / "student-profile.json").write_text(
                json.dumps({"grade": "2025", "major": "计算机"}), encoding="utf-8"
            )
            (root / "current-timetable.json").write_text(
                json.dumps({"term": "2025年秋季学期", "entries": []}), encoding="utf-8"
            )
            database = WorkspaceDatabase.open(root)
            self.assertEqual("2025", database.current_profile()["grade"])
            snapshot = database.latest_snapshot("timetable")
            self.assertEqual("2025年秋季学期", snapshot["payload"]["term"])
            self.assertEqual(database.current_profile()["version_id"], snapshot["profile_id"])
            self.assertEqual("wal", database.connection.execute("pragma journal_mode").fetchone()[0])
            database.close()
            (root / "student-profile.json").write_text(
                json.dumps({"grade": "2024"}), encoding="utf-8"
            )
            database = WorkspaceDatabase.open(root)
            self.assertEqual("2025", database.current_profile()["grade"])
            database.close()

    def test_version_three_cleanup_keeps_identity_notice_and_latest_timetable(self):
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            database = WorkspaceDatabase.open(root)
            with database.connection:
                profile_id = database._insert_profile({"grade": "2025"})
                notice_id = database._insert_notice({"status": "confirmed", "query_eligible": True})
                database.connection.execute("update notice_versions set status='confirmed' where id=?", (notice_id,))
                database.connection.execute("delete from schema_migrations where version=3")
            database.publish_snapshot("timetable", "old", {"entries": []}, source="test", profile_id=profile_id)
            latest = database.publish_snapshot("timetable", "new", {"entries": []}, source="test", profile_id=profile_id)
            database.publish_snapshot("selection", "new", {"sections": []}, source="test", profile_id=profile_id, notice_id=notice_id)
            database.save_plan({"term": "new", "profile_id": profile_id, "notice_id": notice_id, "timetable_snapshot_id": latest["id"], "selection_snapshot_id": "x"})
            database.close()

            database = WorkspaceDatabase.open(root)
            self.assertEqual("2025", database.current_profile()["grade"])
            self.assertIsNotNone(database.confirmed_notice())
            self.assertEqual(latest["id"], database.latest_snapshot("timetable")["id"])
            self.assertIsNone(database.latest_snapshot("selection"))
            self.assertIsNone(database.latest_plan())
            database.close()

    def test_failed_refresh_keeps_last_complete_snapshot(self):
        with tempfile.TemporaryDirectory() as directory:
            database = WorkspaceDatabase.open(Path(directory))
            first = database.publish_snapshot("selection", "2025-2026-2", {"sections": [1]}, source="academic")
            database.record_failed_attempt("selection", "network failure")
            self.assertEqual(first["id"], database.latest_snapshot("selection")["id"])
            database.close()


class ObservationServiceTests(unittest.TestCase):
    def test_restart_marks_abandoned_tasks_failed_before_accepting_new_work(self):
        with tempfile.TemporaryDirectory() as directory:
            database = WorkspaceDatabase.open(Path(directory))
            abandoned_service = ObservationService(database, FakeGateway, autostart=False)
            abandoned = abandoned_service.submit("connect")

            restarted_service = ObservationService(database, FakeGateway, autostart=False)

            self.assertEqual(TaskState.FAILED.value, restarted_service.inspect(abandoned.id)["state"])
            self.assertIn("restarted", restarted_service.inspect(abandoned.id)["error"])
            restarted_service.close()
            abandoned_service.close()
            database.close()

    def test_login_changes_require_an_idle_worker(self):
        with tempfile.TemporaryDirectory() as directory:
            database = WorkspaceDatabase.open(Path(directory))
            service = ObservationService(database, FakeGateway, autostart=False)
            service.submit("connect")
            with self.assertRaisesRegex(RuntimeError, "当前教务任务"):
                service.run_when_idle(lambda: None)
            service.close()
            database.close()

    def test_failed_login_reset_blocks_following_academic_tasks(self):
        with tempfile.TemporaryDirectory() as directory:
            database = WorkspaceDatabase.open(Path(directory))
            service = ObservationService(database, ResetFailureGateway)
            reset = service.submit("reset-login")
            self.assertTrue(service.wait(reset.id, 2))
            self.assertEqual(TaskState.FAILED.value, service.inspect(reset.id)["state"])
            connect = service.submit("connect")
            self.assertTrue(service.wait(connect.id, 2))
            self.assertEqual(TaskState.FAILED.value, service.inspect(connect.id)["state"])
            self.assertIn("登录重置失败", service.inspect(connect.id)["error"])
            service.close()
            database.close()

    def test_progress_refresh_publishes_score_free_snapshot(self):
        with tempfile.TemporaryDirectory() as directory:
            database = WorkspaceDatabase.open(Path(directory))
            gateway = ProgressGateway()
            service = ObservationService(database, lambda: gateway)
            task = service.submit("refresh-progress")
            self.assertTrue(service.wait(task.id, 2))
            self.assertEqual(TaskState.SUCCEEDED.value, service.inspect(task.id)["state"])
            snapshot = database.latest_snapshot("progress")
            self.assertEqual("academic", snapshot["source"])
            self.assertEqual(3, snapshot["payload"]["report"]["progress"][0]["completed_credits"])
            self.assertNotIn("score", json.dumps(snapshot, ensure_ascii=False).lower())
            service.close()
            database.close()

    def test_shell_launch_opens_workbench_without_authenticating(self):
        with tempfile.TemporaryDirectory() as directory:
            database = WorkspaceDatabase.open(Path(directory))
            gateway = ShellGateway()
            service = ObservationService(database, lambda: gateway)
            task = service.submit("launch-shell", {"workbench_url": "http://127.0.0.1:5000"})
            self.assertTrue(service.wait(task.id, 2))
            self.assertEqual(TaskState.SUCCEEDED.value, service.inspect(task.id)["state"])
            self.assertEqual(["http://127.0.0.1:5000"], gateway.shell_urls)
            self.assertEqual(0, gateway.connect_count)
            service.close()
            database.close()

    def test_connect_only_verifies_session_without_refreshing_snapshots(self):
        with tempfile.TemporaryDirectory() as directory:
            database = WorkspaceDatabase.open(Path(directory))
            gateway = ShellGateway()
            service = ObservationService(database, lambda: gateway)
            task = service.submit("connect", {
                "term": "2026-1",
                "allowed_categories": ["szhx"],
            })
            self.assertTrue(service.wait(task.id, 2))
            self.assertEqual(TaskState.SUCCEEDED.value, service.inspect(task.id)["state"])
            self.assertEqual(1, gateway.connect_count)
            self.assertEqual(0, gateway.timetable_count)
            self.assertEqual(0, gateway.selection_count)
            self.assertIsNone(database.latest_snapshot("timetable"))
            self.assertIsNone(database.latest_snapshot("selection"))
            self.assertEqual("connected", service.session_status()["state"])
            self.assertTrue(service.session_status()["last_verified_at"])
            service.close()
            database.close()

    def test_task_context_is_sanitized_before_persistence(self):
        with tempfile.TemporaryDirectory() as directory:
            database = WorkspaceDatabase.open(Path(directory))
            service = ObservationService(database, FakeGateway, autostart=False)
            task = service.submit("connect", {"token": "secret", "nested": {"authorization": "Bearer abc"}})
            context = database.connection.execute("select context from observation_tasks where id=?", (task.id,)).fetchone()[0]
            self.assertNotIn("secret", context)
            self.assertNotIn("Bearer abc", context)
            database.close()

    def test_read_timeout_reaches_terminal_state_and_releases_active_task(self):
        class SlowGateway(FakeGateway):
            observed_timeout = None

            def observe_selection(self, request, progress, cancelled):
                self.observed_timeout = request.context.get("operation_timeout_seconds")
                while not cancelled():
                    time.sleep(0.005)
                return SelectionObservationResult.incomplete("cancelled")

        with tempfile.TemporaryDirectory() as directory, patch.dict(TASK_TIMEOUT_SECONDS, {"refresh-selection": 0.03}):
            database = WorkspaceDatabase.open(Path(directory))
            service = ObservationService(database, SlowGateway)
            task = service.submit("refresh-selection", {"term": "2025-2026-2"})
            self.assertTrue(service.wait(task.id, 2))
            self.assertEqual(TaskState.FAILED.value, service.inspect(task.id)["state"])
            self.assertIn("timed out", service.inspect(task.id)["error"])
            self.assertEqual(0.03, service.gateway.observed_timeout)
            self.assertIsNone(service.active_task())
            service.close()
            database.close()

    def test_different_remote_task_is_rejected_while_one_is_active(self):
        with tempfile.TemporaryDirectory() as directory:
            database = WorkspaceDatabase.open(Path(directory))
            service = ObservationService(database, FakeGateway, autostart=False)
            first = service.submit("refresh-selection", {"term": "2025-2026-2"})
            with self.assertRaisesRegex(RuntimeError, "refresh-selection"):
                service.submit("refresh-timetable", {"term": "2025-2026-2"})
            self.assertEqual(first.id, service.active_task()["id"])
            service.close()
            database.close()

    def test_equivalent_pending_refreshes_are_coalesced(self):
        with tempfile.TemporaryDirectory() as directory:
            database = WorkspaceDatabase.open(Path(directory))
            gateway = FakeGateway()
            service = ObservationService(database, lambda: gateway, autostart=False)
            first = service.submit("refresh-selection", {"term": "2025-2026-2"})
            second = service.submit("refresh-selection", {"term": "2025-2026-2"})
            self.assertEqual(first.id, second.id)
            service.start()
            self.assertTrue(service.wait(first.id, 2))
            self.assertEqual(TaskState.SUCCEEDED.value, service.inspect(first.id)["state"])
            self.assertEqual(1, gateway.selection_count)
            service.close()
            database.close()

    def test_connection_and_consecutive_reads_reuse_one_gateway_until_shutdown(self):
        with tempfile.TemporaryDirectory() as directory:
            database = WorkspaceDatabase.open(Path(directory))
            gateways = []

            def factory():
                gateway = ShellGateway()
                gateways.append(gateway)
                return gateway

            service = ObservationService(database, factory)
            tasks = []
            for operation, context in (
                ("launch-shell", {"workbench_url": "http://127.0.0.1:5000"}),
                ("connect", {"term": "2026-1", "allowed_categories": ["szhx"]}),
                ("refresh-timetable", {"term": "2026-1"}),
                ("refresh-selection", {"term": "2026-1"}),
            ):
                task = service.submit(operation, context)
                tasks.append(task)
                self.assertTrue(service.wait(task.id, 2))

            self.assertEqual(1, len(gateways))
            self.assertFalse(gateways[0].closed)
            self.assertEqual(1, len(gateways[0].shell_urls))
            service.close()
            self.assertTrue(gateways[0].closed)
            database.close()

    def test_unrecoverable_poll_failure_closes_gateway_before_next_task(self):
        with tempfile.TemporaryDirectory() as directory:
            database = WorkspaceDatabase.open(Path(directory))
            gateways = []

            def factory():
                gateway = PollFailureGateway()
                gateways.append(gateway)
                return gateway

            service = ObservationService(database, factory)
            first = service.submit("connect")
            self.assertTrue(service.wait(first.id, 2))
            self.assertTrue(gateways[0].polled.wait(2))
            deadline = time.monotonic() + 2
            while time.monotonic() < deadline and not gateways[0].closed:
                time.sleep(0.01)
            self.assertTrue(gateways[0].closed)

            second = service.submit("connect")
            self.assertTrue(service.wait(second.id, 2))
            self.assertEqual(2, len(gateways))
            service.close()
            database.close()

    def test_authentication_wait_is_observable_and_same_task_resumes(self):
        with tempfile.TemporaryDirectory() as directory:
            database = WorkspaceDatabase.open(Path(directory))
            gateway = AuthenticationGateway()
            service = ObservationService(database, lambda: gateway)
            task = service.submit("refresh-selection", {"term": "2025-2026-2"})
            deadline = time.monotonic() + 2
            while time.monotonic() < deadline:
                if service.inspect(task.id)["state"] == TaskState.WAITING_FOR_AUTHENTICATION.value:
                    break
                time.sleep(0.01)
            self.assertEqual(TaskState.WAITING_FOR_AUTHENTICATION.value, service.inspect(task.id)["state"])
            gateway.authentication_complete.set()
            self.assertTrue(service.wait(task.id, 2))
            self.assertEqual(TaskState.SUCCEEDED.value, service.inspect(task.id)["state"])
            self.assertEqual(1, gateway.connect_count)
            service.close()
            database.close()

    def test_authentication_timeout_fails_task_but_keeps_gateway_session(self):
        with tempfile.TemporaryDirectory() as directory:
            database = WorkspaceDatabase.open(Path(directory))
            gateway = TimedOutAuthenticationGateway()
            service = ObservationService(database, lambda: gateway)
            task = service.submit("connect")
            self.assertTrue(service.wait(task.id, 2))
            self.assertEqual(TaskState.FAILED.value, service.inspect(task.id)["state"])
            self.assertEqual("unknown", service.session_state)
            self.assertEqual("unknown", service.webvpn_state)
            self.assertFalse(gateway.closed)
            service.close()
            self.assertTrue(gateway.closed)
            database.close()


class WorkbenchApiTests(unittest.TestCase):
    def test_local_login_configuration_is_dpapi_protected_and_never_returned(self):
        with tempfile.TemporaryDirectory() as directory:
            app = create_workbench_app(
                Path(directory),
                gateway_factory=FakeGateway,
                require_login_configuration=True,
            )
            client = app.test_client()
            state = client.get("/api/state").get_json()
            self.assertFalse(state["login_configuration"]["configured"])
            headers = {
                "Origin": "http://localhost",
                "Host": "localhost",
                "X-CSRF-Token": state["csrf_token"],
            }
            blocked = client.post(
                "/api/tasks",
                json={"operation": "connect"},
                headers=headers,
            )
            self.assertEqual(409, blocked.status_code)
            response = client.post(
                "/api/login-configuration",
                json={"username": "2025000000", "password": "local-secret"},
                headers=headers,
            )
            self.assertEqual(201, response.status_code)
            payload = response.get_json()
            self.assertTrue(payload["configured"])
            service = app.extensions["observation_service"]
            self.assertNotIn("connection_task", payload)
            self.assertIsNone(service.active_task())
            self.assertEqual("2025******", payload["masked_username"])
            self.assertNotIn("local-secret", json.dumps(payload))
            encrypted = Path(directory) / "course-progress" / "webvpn-login.dpapi"
            self.assertTrue(encrypted.is_file())
            self.assertNotIn(b"local-secret", encrypted.read_bytes())
            changed = client.post(
                "/api/login-configuration",
                json={"username": "2025999999", "password": "other-secret"},
                headers=headers,
            )
            self.assertEqual(400, changed.status_code)
            database = app.extensions["workspace_database"]
            with database.connection:
                database._insert_profile({"grade": "2025"})
            database.publish_snapshot("timetable", "2026-1", {"entries": []}, source="test")
            report = Path(directory) / "course-progress" / "progress-report.json"
            checkpoint = Path(directory) / "course-progress" / "collection-checkpoint.json"
            report.write_text("{}", encoding="utf-8")
            checkpoint.write_text("{}", encoding="utf-8")
            self.assertEqual(204, client.delete("/api/login-configuration", headers=headers).status_code)
            cleared = client.get("/api/state").get_json()
            self.assertFalse(cleared["login_configuration"]["configured"])
            self.assertIsNone(cleared["profile"])
            self.assertIsNone(cleared["snapshots"]["timetable"])
            self.assertFalse(report.exists())
            self.assertFalse(checkpoint.exists())
            app.extensions["observation_service"].close()
            app.extensions["workspace_database"].close()

    def test_production_frontend_is_served_from_configured_build_directory(self):
        with tempfile.TemporaryDirectory() as directory:
            frontend = Path(directory) / "frontend"
            (frontend / "assets").mkdir(parents=True)
            (frontend / "index.html").write_text(
                '<div id="root">react-build</div>', encoding="utf-8"
            )
            (frontend / "assets" / "app.js").write_text("bundle", encoding="utf-8")
            app = create_workbench_app(
                Path(directory) / "workspace",
                frontend_root=frontend,
                gateway_factory=FakeGateway,
            )
            client = app.test_client()
            self.assertIn("react-build", client.get("/").get_data(as_text=True))
            self.assertEqual("bundle", client.get("/assets/app.js").get_data(as_text=True))
            app.extensions["observation_service"].close()
            app.extensions["workspace_database"].close()

    def test_manual_notice_import_is_not_exposed(self):
        with tempfile.TemporaryDirectory() as directory:
            app = create_workbench_app(Path(directory), gateway_factory=FakeGateway)
            client = app.test_client()
            state = client.get("/api/state").get_json()
            headers = {"Origin": "http://localhost", "Host": "localhost", "X-CSRF-Token": state["csrf_token"]}

            response = client.post(
                "/api/notices/candidates",
                json={"source_url": "https://jwc.hitwh.edu.cn/a", "text": "通知正文"},
                headers=headers,
            )

            self.assertEqual(405, response.status_code)
            self.assertEqual([], client.get("/api/notices/candidates").get_json()["notices"])
            app.extensions["observation_service"].close()
            app.extensions["workspace_database"].close()

    @patch("course_selection.workbench_service.discover_official_notice_candidates")
    def test_official_notice_discovery_api_saves_static_candidates(self, discover):
        discover.return_value = [candidate_from_text(
            "https://jwc.hitwh.edu.cn/2026/notice/page.htm",
            "关于2026年秋季学期各类课程选课时间安排的通知\n"
            "2026年8月29日8:30-2026年8月29日11:00\n2025级\n"
            "文化素质教育课——选课\n新教务系统",
            official_hosts=("jwc.hitwh.edu.cn",),
        )]
        with tempfile.TemporaryDirectory() as directory:
            app = create_workbench_app(Path(directory), gateway_factory=FakeGateway)
            client = app.test_client()
            state = client.get("/api/state").get_json()
            headers = {"Origin": "http://localhost", "Host": "localhost", "X-CSRF-Token": state["csrf_token"]}

            response = client.post("/api/notices/discover", json={}, headers=headers)

            self.assertEqual(201, response.status_code)
            self.assertEqual(1, len(response.get_json()["notices"]))
            self.assertEqual(1, len(client.get("/api/notices/candidates").get_json()["notices"]))
            app.extensions["observation_service"].close()
            app.extensions["workspace_database"].close()

    def test_task_submission_validates_json(self):
        with tempfile.TemporaryDirectory() as directory:
            app = create_workbench_app(Path(directory), gateway_factory=FakeGateway)
            client = app.test_client()
            state = client.get("/api/state").get_json()
            headers = {"Origin": "http://localhost", "Host": "localhost", "X-CSRF-Token": state["csrf_token"]}
            self.assertEqual(400, client.post("/api/tasks", json={"operation": "nope"}, headers=headers).status_code)
            app.extensions["observation_service"].close()
            app.extensions["workspace_database"].close()

    def test_read_only_plan_is_persisted_without_starting_gateway(self):
        with tempfile.TemporaryDirectory() as directory:
            gateway = FakeGateway()
            app = create_workbench_app(Path(directory), gateway_factory=lambda: gateway)
            client = app.test_client()
            state = client.get("/api/state").get_json()
            headers = {"Origin": "http://localhost", "Host": "localhost", "X-CSRF-Token": state["csrf_token"]}
            response = client.post("/api/plans", json={"goals": [{"course_identity": "MATH", "rank": 1}]}, headers=headers)
            self.assertEqual(201, response.status_code)
            self.assertEqual("blocked", response.get_json()["status"])
            latest = client.get("/api/plans/latest")
            self.assertEqual(200, latest.status_code)
            self.assertEqual(response.get_json()["id"], latest.get_json()["id"])
            self.assertEqual(0, gateway.connect_count)
            app.extensions["observation_service"].close()
            app.extensions["workspace_database"].close()

    def test_manual_navigation_observation_can_be_finished_through_api(self):
        with tempfile.TemporaryDirectory() as directory:
            gateway = ManualObservationGateway()
            app = create_workbench_app(Path(directory), gateway_factory=lambda: gateway)
            client = app.test_client()
            state = client.get("/api/state").get_json()
            headers = {"Origin": "http://localhost", "Host": "localhost", "X-CSRF-Token": state["csrf_token"]}
            created = client.post("/api/tasks", json={"operation": "observe-navigation"}, headers=headers)
            self.assertEqual(202, created.status_code)
            task = created.get_json()
            self.assertEqual("observe-navigation", task["operation"])
            self.assertTrue(gateway.observing.wait(2))
            finished = client.post(f"/api/tasks/{task['id']}/finish", headers=headers)
            self.assertEqual(202, finished.status_code)
            service = app.extensions["observation_service"]
            self.assertTrue(service.wait(task["id"], 2))
            report = client.get(f"/api/tasks/{task['id']}").get_json()
            self.assertEqual(TaskState.SUCCEEDED.value, report["state"])
            self.assertEqual(1, report["progress"]["report"]["browser_instances"])
            service.close()
            app.extensions["workspace_database"].close()

    def test_loading_state_reads_sqlite_without_starting_gateway(self):
        with tempfile.TemporaryDirectory() as directory:
            gateway = FakeGateway()
            app = create_workbench_app(Path(directory), gateway_factory=lambda: gateway)
            response = app.test_client().get("/api/state")
            self.assertEqual(200, response.status_code)
            self.assertEqual(0, gateway.connect_count)
            self.assertEqual("disconnected", response.get_json()["academic_session"]["state"])
            app.extensions["observation_service"].close()
            app.extensions["workspace_database"].close()

    def test_state_changes_require_same_origin_and_csrf(self):
        with tempfile.TemporaryDirectory() as directory:
            app = create_workbench_app(Path(directory), gateway_factory=FakeGateway)
            client = app.test_client()
            self.assertEqual(403, client.post("/api/tasks", json={"operation": "connect"}).status_code)
            self.assertEqual(403, client.get("/api/state", headers={"Host": "attacker.example"}).status_code)
            state_response = client.get("/api/state")
            self.assertEqual("no-store", state_response.headers["Cache-Control"])
            self.assertIn("frame-ancestors 'none'", state_response.headers["Content-Security-Policy"])
            state = state_response.get_json()
            response = client.post("/api/tasks", json={"operation": "connect"}, headers={"Origin": "http://localhost", "Host": "localhost", "X-CSRF-Token": state["csrf_token"]})
            self.assertEqual(202, response.status_code)
            app.extensions["observation_service"].close()
            app.extensions["workspace_database"].close()


class NoticeDiscoveryTests(unittest.TestCase):
    TITLE = "关于2026年秋季学期各类课程选课时间安排的通知"

    def test_generic_selection_article_is_not_a_candidate(self):
        with self.assertRaises(ValueError):
            candidate_from_text("https://jwc.hitwh.edu.cn/a", "关于选课的通知", official_hosts=("jwc.hitwh.edu.cn",))

    def test_static_list_parser_keeps_only_arrangement_notice_and_deduplicates_nested_links(self):
        html = f'''<ul>
          <li><a href="/other/page.htm" title="关于2026年秋季学期文化素质课程选课的通知">其他</a></li>
          <li><a href="/notice/page.htm"><a href="/notice/page.htm" title="{self.TITLE}">{self.TITLE}</a></a></li>
        </ul>'''

        self.assertEqual(
            (OfficialNoticeLink(self.TITLE, "https://jwc.hitwh.edu.cn/notice/page.htm"),),
            parse_official_notice_links(html, index_url="https://jwc.hitwh.edu.cn/ks/list.htm"),
        )

    def test_static_article_parser_preserves_table_cells_and_inline_split_numbers(self):
        html = '''<html><div class="wp_articlecontent">
          <p>2026年8月29日8:30-2026年8月29日11:00</p>
          <p><span>2023级、2024级、</span><span>2</span><span>025级</span></p>
          <p>创新研修课、创新实验课、创新创业课——选课</p><p>新教务系统</p>
        </div></html>'''

        text = parse_official_notice_article(html, title=self.TITLE)
        candidate = candidate_from_text(
            "https://jwc.hitwh.edu.cn/notice/page.htm",
            text,
            official_hosts=("jwc.hitwh.edu.cn",),
        )

        self.assertEqual(("2023", "2024", "2025"), tuple(candidate["windows"][0]["grades"]))
        self.assertTrue(candidate["query_eligible"])

    @patch("course_selection.notice_discovery._download_html")
    def test_discovery_downloads_list_then_matching_static_article(self, download):
        download.side_effect = [
            f'<a href="/notice/page.htm" title="{self.TITLE}">{self.TITLE}</a>',
            '<div class="wp_articlecontent"><p>2026年8月29日8:30-2026年8月29日11:00</p><p>2025级</p><p>文化素质教育课——选课</p><p>新教务系统</p></div>',
        ]

        candidates = discover_official_notice_candidates()

        self.assertEqual(1, len(candidates))
        self.assertEqual(self.TITLE, candidates[0]["title"])
        self.assertEqual("https://jwc.hitwh.edu.cn/notice/page.htm", candidates[0]["source_url"])
        self.assertEqual(2, download.call_count)

    @patch("course_selection.notice_discovery._download_html")
    def test_discovery_checks_second_page_when_first_has_no_matching_notice(self, download):
        download.side_effect = [
            '<a href="/other.htm" title="普通教务通知">普通通知</a>',
            f'<a href="/notice/page.htm" title="{self.TITLE}">{self.TITLE}</a>',
            '<div class="wp_articlecontent"><p>2026年8月29日8:30-2026年8月29日11:00</p><p>2025级</p><p>文化素质教育课——选课</p><p>新教务系统</p></div>',
        ]

        candidates = discover_official_notice_candidates()

        self.assertEqual(1, len(candidates))
        self.assertEqual("https://jwc.hitwh.edu.cn/ks/list2.htm", download.call_args_list[1].args[0])
        self.assertEqual(3, download.call_count)


if __name__ == "__main__":
    unittest.main()
