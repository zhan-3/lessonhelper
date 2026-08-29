import json
import tempfile
import unittest
from pathlib import Path

from course_selection.persistence import WorkspaceDatabase
from course_selection.workbench_service import WorkbenchService


class WorkbenchServiceTests(unittest.TestCase):
    def test_state_publishes_completed_course_progress_without_turning_missing_into_zero(self):
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            database = WorkspaceDatabase.open(root / "workspace")
            report_path = root / "progress-report.json"
            report_path.write_text(json.dumps({
                "data_complete": True,
                "progress": [{
                    "key": "cultural_quality",
                    "required_credits": 8,
                    "completed_credits": 5,
                    "courses": [{"code": "HIST", "name": "四史专题", "credits": 2}],
                }],
            }, ensure_ascii=False), encoding="utf-8")
            service = WorkbenchService(database, progress_report_path=report_path)
            progress = service.state(session_state="disconnected")["graduation_progress"]
            self.assertEqual("ready", progress["status"])
            self.assertEqual(5, progress["report"]["progress"][0]["completed_credits"])
            report_path.unlink()
            self.assertEqual("missing", service.graduation_progress()["status"])
            database.close()

    def test_progress_snapshot_must_match_current_profile_and_baseline(self):
        with tempfile.TemporaryDirectory() as directory:
            database = WorkspaceDatabase.open(Path(directory))
            with database.connection:
                current_profile = database._insert_profile({"grade": "2025"})
            database.publish_snapshot(
                "progress",
                "",
                {
                    "report": {
                        "data_complete": True,
                        "baseline_version": "guide-2026",
                        "progress": [],
                    }
                },
                source="test",
                profile_id="another-profile",
            )
            service = WorkbenchService(database)
            self.assertEqual("not_applicable", service.graduation_progress()["status"])
            with database.connection:
                database.connection.execute("update refresh_attempts set snapshot_id=null")
                database.connection.execute("delete from snapshot_changes")
                database.connection.execute("delete from snapshots")
            database.publish_snapshot(
                "progress",
                "",
                {
                    "report": {
                        "data_complete": True,
                        "baseline_version": "wrong-baseline",
                        "progress": [],
                    }
                },
                source="test",
                profile_id=current_profile,
            )
            self.assertEqual("not_applicable", service.graduation_progress()["status"])
            database.close()

    def test_refresh_context_is_computed_without_flask(self):
        with tempfile.TemporaryDirectory() as directory:
            database = WorkspaceDatabase.open(Path(directory))
            with database.connection:
                profile_id = database._insert_profile({"grade": "2025"})
            notice = database.save_notice({
                "term": "2025年春季学期",
                "windows": [{
                    "action": "selection", "method": "academic_system",
                    "grades": ["2025"], "category_codes": ["szhx", "public"],
                }],
                "query_eligible": True,
            })
            database.confirm_notice(notice["version_id"])
            service = WorkbenchService(database)
            context = service.refresh_context()
            self.assertEqual("2025年春季学期", context["term"])
            self.assertEqual(["szhx", "public"], context["allowed_categories"])
            self.assertEqual(profile_id, context["profile_id"])
            self.assertEqual(notice["version_id"], context["notice_id"])
            database.close()

    def test_candidate_use_case_returns_saved_notice_and_diff(self):
        with tempfile.TemporaryDirectory() as directory:
            database = WorkspaceDatabase.open(Path(directory))
            service = WorkbenchService(database)
            with self.assertRaises(ValueError):
                service.create_notice_candidate("https://example.com/news", "course selection")
            database.close()

    def test_grade_derives_from_student_id_prefix_when_profile_missing(self):
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            database = WorkspaceDatabase.open(root / "workspace")
            login_root = root / "course-progress"
            login_root.mkdir(parents=True)
            (login_root / "webvpn-login.dpapi").write_bytes(b"")
            (login_root / "webvpn-login-meta.json").write_text(
                json.dumps({"masked_username": "2025******"}, ensure_ascii=False),
                encoding="utf-8",
            )
            service = WorkbenchService(database, login_root=login_root)
            self.assertEqual("2025", service.effective_profile().get("grade"))
            self.assertEqual("2025", service.state(session_state="disconnected")["profile"].get("grade"))
            database.close()

    def test_refresh_context_matches_windows_from_derived_grade(self):
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            database = WorkspaceDatabase.open(root / "workspace")
            login_root = root / "course-progress"
            login_root.mkdir(parents=True)
            (login_root / "webvpn-login.dpapi").write_bytes(b"")
            (login_root / "webvpn-login-meta.json").write_text(
                json.dumps({"masked_username": "2025******"}, ensure_ascii=False),
                encoding="utf-8",
            )
            notice = database.save_notice({
                "term": "2026年春季学期",
                "windows": [{
                    "action": "selection", "method": "academic_system",
                    "grades": ["2025"], "category_codes": ["ty", "public"],
                }],
                "query_eligible": True,
            })
            database.confirm_notice(notice["version_id"])
            service = WorkbenchService(database, login_root=login_root)
            context = service.refresh_context()
            self.assertEqual(["ty", "public"], context["allowed_categories"])
            self.assertIsNone(context["profile_id"])
            database.close()


if __name__ == "__main__":
    unittest.main()
