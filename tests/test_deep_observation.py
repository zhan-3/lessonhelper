import json
import tempfile
import unittest
from pathlib import Path

from course_selection.deep_observation import (
    AcademicRequestTrace,
    ReplayAcademicObserver,
    TimetableObservationRequest,
    TimetableObservationResult,
    TraceStore,
)
from course_selection.persistence import WorkspaceDatabase
from course_selection.tasks import ObservationService, TaskState


class DeepObservationTests(unittest.TestCase):
    def test_replay_complete_timetable_observation_publishes_snapshot_and_trace(self):
        trace = AcademicRequestTrace.from_requests([
            {"method": "GET", "url": "https://academic.test/home?ticket=secret", "resource_type": "document"},
            {"method": "POST", "url": "https://academic.test/kbcx/queryXszkb", "resource_type": "fetch", "post_data": "xnxq=2026-1&studentNo=2025000000"},
        ])
        observer = ReplayAcademicObserver(TimetableObservationResult.complete(
            term="2026-1", entries=[{"course_name": "程序设计"}], trace=trace,
        ))
        with tempfile.TemporaryDirectory() as directory:
            database = WorkspaceDatabase.open(Path(directory))
            service = ObservationService(database, lambda: observer)
            task = service.submit("refresh-timetable", {"term": "2026-1"})
            self.assertTrue(service.wait(task.id, 2))

            self.assertEqual(TaskState.SUCCEEDED.value, service.inspect(task.id)["state"])
            self.assertEqual("程序设计", database.latest_snapshot("timetable")["payload"]["entries"][0]["course_name"])
            trace_path = Path(directory) / "request-traces" / f"{task.id}.jsonl"
            trace_rows = [json.loads(line) for line in trace_path.read_text(encoding="utf-8").splitlines()]
            self.assertEqual([1, 2], [row["sequence"] for row in trace_rows])
            self.assertNotIn("secret", json.dumps(trace_rows))
            self.assertNotIn("2025000000", json.dumps(trace_rows))
            self.assertEqual(["studentNo", "xnxq"], trace_rows[1]["field_names"])
            service.close()
            database.close()

    def test_cancelled_or_incomplete_observation_keeps_existing_snapshot(self):
        for result in (
            TimetableObservationResult.cancelled(),
            TimetableObservationResult.incomplete("missing timetable page"),
        ):
            with self.subTest(status=result.status), tempfile.TemporaryDirectory() as directory:
                database = WorkspaceDatabase.open(Path(directory))
                existing = database.publish_snapshot("timetable", "2026-1", {"entries": [{"course_name": "Existing"}]}, source="test")
                service = ObservationService(database, lambda: ReplayAcademicObserver(result))
                task = service.submit("refresh-timetable", {"term": "2026-1"})
                self.assertTrue(service.wait(task.id, 2))
                self.assertEqual(existing["id"], database.latest_snapshot("timetable")["id"])
                self.assertNotEqual(TaskState.SUCCEEDED.value, service.inspect(task.id)["state"])
                service.close()
                database.close()

    def test_trace_write_failure_does_not_prevent_publish_and_marks_task(self):
        observer = ReplayAcademicObserver(TimetableObservationResult.complete(
            term="2026-1", entries=[{"course_name": "程序设计"}], trace=AcademicRequestTrace.empty(),
        ))
        with tempfile.TemporaryDirectory() as directory:
            database = WorkspaceDatabase.open(Path(directory))
            service = ObservationService(database, lambda: observer, trace_store=TraceStore(Path(directory) / "not-a-directory"))
            (Path(directory) / "not-a-directory").write_text("blocked", encoding="utf-8")
            task = service.submit("refresh-timetable", {"term": "2026-1"})
            self.assertTrue(service.wait(task.id, 2))
            self.assertEqual(TaskState.SUCCEEDED.value, service.inspect(task.id)["state"])
            self.assertTrue(service.inspect(task.id)["progress"]["trace_incomplete"])
            self.assertIsNotNone(database.latest_snapshot("timetable"))
            service.close()
            database.close()

    def test_replay_selection_observation_publishes_complete_sections_and_trace(self):
        from course_selection.deep_observation import SelectionObservationResult

        observer = ReplayAcademicObserver(SelectionObservationResult.complete(
            {"term": "2026-1", "source_kind": "verified-selection-api", "sections": [{"identity": "TASK-1"}]},
            trace=AcademicRequestTrace.from_requests([{"method": "POST", "url": "https://academic.test/query", "resource_type": "fetch"}]),
        ))
        with tempfile.TemporaryDirectory() as directory:
            database = WorkspaceDatabase.open(Path(directory))
            service = ObservationService(database, lambda: observer)
            task = service.submit("refresh-selection", {"term": "2026-1"})
            self.assertTrue(service.wait(task.id, 2))
            self.assertEqual(TaskState.SUCCEEDED.value, service.inspect(task.id)["state"])
            self.assertEqual("TASK-1", database.latest_snapshot("selection")["payload"]["sections"][0]["identity"])
            self.assertTrue((Path(directory) / "request-traces" / f"{task.id}.jsonl").is_file())
            service.close()
            database.close()

    def test_trace_store_keeps_only_twenty_ended_task_traces(self):
        with tempfile.TemporaryDirectory() as directory:
            store = TraceStore(Path(directory))
            for number in range(21):
                store.write(f"task-{number}", AcademicRequestTrace.empty())
            self.assertFalse((Path(directory) / "task-0.jsonl").exists())
            self.assertTrue((Path(directory) / "task-20.jsonl").exists())


if __name__ == "__main__":
    unittest.main()
