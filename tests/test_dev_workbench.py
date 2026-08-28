import sqlite3
import tempfile
import unittest
from pathlib import Path

from course_selection.cli import build_parser
from course_selection.dev_workbench import has_active_tasks, source_mtimes


class DevWorkbenchTests(unittest.TestCase):
    def test_cli_exposes_fixed_loopback_devtools_port(self):
        args = build_parser().parse_args(["dev-workbench"])
        self.assertEqual("dev-workbench", args.command)
        self.assertEqual(9222, args.debug_port)
        self.assertEqual(5000, args.port)

    def test_source_watch_ignores_private_data_and_tracks_python(self):
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            (root / "course_selection").mkdir()
            (root / "course_progress").mkdir()
            source = root / "course_selection" / "gateway.py"
            source.write_text("x = 1", encoding="utf-8")
            (root / ".private" / "course_selection").mkdir(parents=True)
            (root / ".private" / "course_selection" / "ignored.py").write_text("x = 2", encoding="utf-8")
            self.assertEqual({source}, set(source_mtimes(root)))

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
