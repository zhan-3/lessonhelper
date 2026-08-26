import json
import tempfile
import threading
import time
import unittest
from pathlib import Path

from course_selection.persistence import WorkspaceDatabase
from course_selection.notice_discovery import candidate_from_text
from course_selection.tasks import ObservationService, TaskState
from course_selection.workbench import create_workbench_app


class FakeGateway:
    def __init__(self):
        self.connect_count = 0
        self.selection_count = 0
        self.closed = False

    def connect(self, progress, cancelled):
        self.connect_count += 1
        progress("connecting", {})

    def refresh_selection(self, context, progress, cancelled):
        self.selection_count += 1
        progress("reading", {"category": "szhx", "page": 1})
        return {"status": "complete", "sections": [{"name": "艺术鉴赏"}]}

    def refresh_timetable(self, context, progress, cancelled):
        return {"status": "interface_unconfirmed", "entries": []}

    def close(self):
        self.closed = True


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

    def test_failed_refresh_keeps_last_complete_snapshot(self):
        with tempfile.TemporaryDirectory() as directory:
            database = WorkspaceDatabase.open(Path(directory))
            first = database.publish_snapshot("selection", "2025-2026-2", {"sections": [1]}, source="academic")
            database.record_failed_attempt("selection", "network failure")
            self.assertEqual(first["id"], database.latest_snapshot("selection")["id"])
            database.close()


class ObservationServiceTests(unittest.TestCase):
    def test_task_context_is_sanitized_before_persistence(self):
        with tempfile.TemporaryDirectory() as directory:
            database = WorkspaceDatabase.open(Path(directory))
            service = ObservationService(database, FakeGateway, autostart=False)
            task = service.submit("connect", {"token": "secret", "nested": {"authorization": "Bearer abc"}})
            context = database.connection.execute("select context from observation_tasks where id=?", (task.id,)).fetchone()[0]
            self.assertNotIn("secret", context)
            self.assertNotIn("Bearer abc", context)
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
            self.assertEqual(TaskState.WAITING_FOR_AUTHENTICATION.value, service.session_state)
            self.assertFalse(gateway.closed)
            service.close()
            self.assertTrue(gateway.closed)
            database.close()


class WorkbenchApiTests(unittest.TestCase):
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

    def test_notice_check_requires_csrf_and_lists_candidates(self):
        with tempfile.TemporaryDirectory() as directory:
            app = create_workbench_app(Path(directory), gateway_factory=FakeGateway)
            client = app.test_client()
            body = {"source_url": "https://jwc.hitwh.edu.cn/a", "text": "关于选课的通知"}
            self.assertEqual(403, client.post("/api/notices/candidates", json=body).status_code)
            state = client.get("/api/state").get_json()
            response = client.post("/api/notices/candidates", json=body, headers={"Origin": "http://localhost", "Host": "localhost", "X-CSRF-Token": state["csrf_token"]})
            self.assertEqual(422, response.status_code)
            self.assertEqual([], client.get("/api/notices/candidates").get_json()["notices"])
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
            state = client.get("/api/state").get_json()
            response = client.post("/api/tasks", json={"operation": "connect"}, headers={"Origin": "http://localhost", "Host": "localhost", "X-CSRF-Token": state["csrf_token"]})
            self.assertEqual(202, response.status_code)
            app.extensions["observation_service"].close()
            app.extensions["workspace_database"].close()


class NoticeDiscoveryTests(unittest.TestCase):
    def test_generic_selection_article_is_not_a_candidate(self):
        with self.assertRaises(ValueError):
            candidate_from_text("https://jwc.hitwh.edu.cn/a", "关于选课的通知", official_hosts=("jwc.hitwh.edu.cn",))


if __name__ == "__main__":
    unittest.main()
