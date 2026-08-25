import json
import tempfile
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


class WorkbenchApiTests(unittest.TestCase):
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
