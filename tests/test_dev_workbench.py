import sqlite3
import tempfile
import unittest
from pathlib import Path

from watchfiles import Change, PythonFilter

from course_selection.cli import dev_workbench_cmd
from course_selection.dev_workbench import WATCHED_PACKAGES, has_active_tasks


class DevWorkbenchTests(unittest.TestCase):
    def test_cli_exposes_fixed_loopback_devtools_port(self):
        params = {param.name: param for param in dev_workbench_cmd.params}
        self.assertEqual(9222, params["debug_port"].default)
        self.assertEqual(5000, params["port"].default)

    def test_source_watch_targets_packages_and_tracks_python_only(self):
        # The watcher watches only the two package directories, so anything
        # outside them (for example .private/) can never trigger a restart.
        self.assertEqual(("course_selection", "course_progress"), WATCHED_PACKAGES)
        watch_filter = PythonFilter()
        self.assertTrue(watch_filter(Change.added, "course_selection/gateway.py"))
        self.assertTrue(watch_filter(Change.modified, "course_progress/explorer.py"))
        self.assertFalse(watch_filter(Change.modified, "course_selection/data.json"))

    def test_active_task_check_defers_restart(self):
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            database = root / "workbench.sqlite3"
            connection = sqlite3.connect(database)
            connection.execute("create table observation_tasks(state text)")
            connection.execute("insert into observation_tasks values('reading')")
            connection.commit()
            connection.close()
            self.assertTrue(has_active_tasks(root))


if __name__ == "__main__":
    unittest.main()
