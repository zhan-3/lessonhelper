import sqlite3
import tempfile
import unittest
from pathlib import Path
from unittest.mock import MagicMock, patch
from urllib.error import URLError

from watchfiles import Change, PythonFilter

from course_selection.cli import dev_workbench_cmd
from course_selection.dev_workbench import (
    WATCHED_PACKAGES,
    has_active_tasks,
    run_dev_workbench,
    wait_for_cdp,
)


class _Response:
    status = 200

    def __enter__(self):
        return self

    def __exit__(self, *_args):
        return False


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

    def test_cdp_readiness_retries_until_devtools_endpoint_answers(self):
        with patch(
            "course_selection.dev_workbench.urlopen",
            side_effect=[URLError("not ready"), _Response()],
        ) as open_url:
            self.assertTrue(wait_for_cdp("http://127.0.0.1:9222", timeout=1, retry_interval=0))
        self.assertEqual(2, open_url.call_count)

    def test_supervisor_does_not_restart_a_child_that_exited_normally(self):
        playwright = MagicMock()
        playwright.__enter__.return_value = object()
        browser_context = MagicMock()
        child = MagicMock()
        child.poll.return_value = 0

        with tempfile.TemporaryDirectory() as directory, patch(
            "course_selection.dev_workbench.sync_playwright", return_value=playwright
        ), patch(
            "course_selection.dev_workbench.launch_browser_context", return_value=browser_context
        ), patch(
            "course_selection.dev_workbench.wait_for_cdp", return_value=True
        ), patch(
            "course_selection.dev_workbench.watch", return_value=iter(())
        ), patch(
            "course_selection.dev_workbench.time.sleep"
        ), patch(
            "course_selection.dev_workbench.stop_workbench"
        ), patch(
            "course_selection.dev_workbench.start_workbench",
            side_effect=[child, AssertionError("normal child exit must not trigger a restart loop")],
        ) as start:
            self.assertEqual(0, run_dev_workbench(Path(directory), Path(directory), 5000))

        self.assertEqual(1, start.call_count)

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
