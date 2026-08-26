import tempfile
import unittest
from pathlib import Path

from course_selection.persistence import WorkspaceDatabase
from course_selection.workbench_service import WorkbenchService


class WorkbenchServiceTests(unittest.TestCase):
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


if __name__ == "__main__":
    unittest.main()
