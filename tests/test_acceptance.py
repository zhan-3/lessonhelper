"""Deterministic ticket-17 acceptance tests for the persistent workbench."""

from __future__ import annotations

import json
import sqlite3
import tempfile
import threading
import time
import unittest
from pathlib import Path

from course_selection.deep_observation import AcademicRequestTrace, SelectionObservationResult, TimetableObservationResult
from course_selection.persistence import WorkspaceDatabase
from course_selection.single_instance import WorkspaceLock
from course_selection.tasks import TaskState
from course_selection.workbench import create_workbench_app


class DeterministicGateway:
    def __init__(self, *, fail=False, block=False):
        self.fail = fail
        self.block = block
        self.started = threading.Event()
        self.release = threading.Event()
        self.calls: list[str] = []
        self.mutations = 0

    def connect(self, progress, cancelled):
        self.calls.append("connect")
        progress("connecting", {"message": "deterministic gateway"})

    def refresh_selection(self, context, progress, cancelled):
        self.calls.append("refresh-selection")
        self.started.set()
        while self.block and not self.release.wait(0.01):
            if cancelled():
                return {"status": "cancelled", "sections": []}
        if self.fail:
            raise OSError("upstream failed token=secret-cookie")
        progress("reading", {"category": "safe", "page": 1})
        return {"status": "complete", "sections": [{"name": "Safe course"}]}

    def refresh_timetable(self, context, progress, cancelled):
        self.calls.append("refresh-timetable")
        return {"status": "complete", "entries": [{"course_name": "Safe course"}]}

    def observe_selection(self, request, progress, cancelled):
        result = self.refresh_selection(request.context, progress, cancelled)
        if result.get("status") != "complete":
            return SelectionObservationResult.incomplete(result.get("status", "incomplete"))
        return SelectionObservationResult.complete(result, trace=AcademicRequestTrace.empty())

    def observe_timetable(self, request, progress, cancelled):
        result = self.refresh_timetable(request.context, progress, cancelled)
        return TimetableObservationResult.complete(
            term=str(request.context.get("term", "")), entries=result["entries"],
            trace=AcademicRequestTrace.empty(),
        )

    def submit_selection(self, *_args, **_kwargs):
        self.mutations += 1
        raise AssertionError("observation gateway must never expose a write path")

    def close(self):
        self.calls.append("close")


def wait_for(service, identity, expected, timeout=2):
    deadline = time.monotonic() + timeout
    while time.monotonic() < deadline:
        state = service.inspect(identity)
        if state and state["state"] in expected:
            return state
        time.sleep(0.01)
    return service.inspect(identity)


class PersistentWorkbenchAcceptanceTests(unittest.TestCase):
    def test_api_success_cancel_failure_restart_and_no_mutation(self):
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            gateway = DeterministicGateway()
            app = create_workbench_app(root, gateway_factory=lambda: gateway)
            client = app.test_client()
            state = client.get("/api/state").get_json()
            headers = {"Origin": "http://localhost", "Host": "localhost", "X-CSRF-Token": state["csrf_token"]}

            first = client.post("/api/tasks", json={"operation": "refresh-selection"}, headers=headers).get_json()
            service = app.extensions["observation_service"]
            self.assertTrue(service.wait(first["id"], 2))
            self.assertEqual(TaskState.SUCCEEDED.value, client.get(f"/api/tasks/{first['id']}").get_json()["state"])
            snapshot = client.get("/api/state").get_json()["snapshots"]["selection"]

            blocked_gateway = DeterministicGateway(block=True)
            service.gateway = blocked_gateway
            second = client.post("/api/tasks", json={"operation": "refresh-selection", "context": {"run": 2}}, headers=headers).get_json()
            self.assertTrue(blocked_gateway.started.wait(2))
            self.assertEqual(202, client.delete(f"/api/tasks/{second['id']}", headers=headers).status_code)
            self.assertTrue(service.wait(second["id"], 2))
            self.assertEqual(TaskState.CANCELLED.value, client.get(f"/api/tasks/{second['id']}").get_json()["state"])
            self.assertEqual(snapshot["id"], client.get("/api/state").get_json()["snapshots"]["selection"]["id"])

            service.gateway = DeterministicGateway(fail=True)
            third = client.post("/api/tasks", json={"operation": "refresh-selection", "context": {"run": 3}}, headers=headers).get_json()
            self.assertTrue(service.wait(third["id"], 2))
            self.assertEqual(TaskState.FAILED.value, client.get(f"/api/tasks/{third['id']}").get_json()["state"])
            self.assertEqual(snapshot["id"], client.get("/api/state").get_json()["snapshots"]["selection"]["id"])
            failed_attempt = service.database.connection.execute("select state,error from refresh_attempts where kind='selection' and state='failed' order by rowid desc limit 1").fetchone()
            self.assertEqual("failed", failed_attempt[0])
            self.assertNotIn("token=secret-cookie", failed_attempt[1])
            self.assertEqual(0, gateway.mutations)
            service.close()
            app.extensions["workspace_database"].close()

            reopened = WorkspaceDatabase.open(root)
            self.assertEqual(snapshot["id"], reopened.latest_snapshot("selection")["id"])
            reopened.close()

    def test_retention_and_sensitive_diagnostics_are_safe(self):
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            db = WorkspaceDatabase.open(root)
            for index in range(25):
                db.publish_snapshot("selection", "term", {
                    "sections": [{"name": f"course-{index}"}],
                    "raw_html": "<html>cookie=bad</html>",
                    "student_number": "20250001",
                    "nested": {"authorization": "Bearer abc"},
                }, source="deterministic")
            count = db.connection.execute("select count(*) from snapshots where kind='selection'").fetchone()[0]
            self.assertEqual(3, count)
            change_count = db.connection.execute("select count(*) from snapshot_changes where kind='selection'").fetchone()[0]
            self.assertEqual(20, change_count)
            raw = " ".join(str(row[0]) for row in db.connection.execute("select payload from snapshots"))
            self.assertNotIn("<html>", raw)
            self.assertNotIn("20250001", raw)
            self.assertNotIn("Bearer abc", raw)
            db.record_failed_attempt("selection", "<html>secret</html> token=abc")
            diagnostics = str(db.connection.execute("select error from refresh_attempts order by finished_at desc limit 1").fetchone()[0])
            self.assertNotIn("<html>", diagnostics)
            self.assertNotIn("token=abc", diagnostics)
            db.close()

    def test_successful_snapshots_record_sanitized_structured_diffs(self):
        with tempfile.TemporaryDirectory() as directory:
            db = WorkspaceDatabase.open(Path(directory))
            first = db.publish_snapshot(
                "timetable", "2026-1",
                {"entries": [{"course_code": "A", "location": "1-101"}], "token": "secret"},
                source="personal-timetable-api", profile_id="profile-1",
            )
            baseline = db.latest_snapshot_change("timetable")
            self.assertTrue(baseline["payload"]["baseline"])
            second = db.publish_snapshot(
                "timetable", "2026-1",
                {"entries": [{"course_code": "A", "location": "1-102"}], "token": "changed-secret"},
                source="personal-timetable-api", profile_id="profile-1",
            )
            change = db.latest_snapshot_change("timetable")
            self.assertEqual(first["id"], change["previous_snapshot_id"])
            self.assertEqual(second["id"], change["snapshot_id"])
            self.assertTrue(change["payload"]["changed"])
            self.assertIn("/entries/0/location", [item["path"] for item in change["payload"]["changes"]])
            self.assertNotIn("changed-secret", json.dumps(change, ensure_ascii=False))
            db.close()

    def test_single_instance_lock_recovers_stale_descriptor(self):
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            lock = WorkspaceLock(root, 8765)
            lock.path.write_text(json.dumps({"pid": 99999999, "port": 8765}), encoding="utf-8")
            self.assertTrue(lock.acquire())
            self.assertFalse(WorkspaceLock(root, 8765).acquire())
            lock.release()


if __name__ == "__main__":
    unittest.main()
