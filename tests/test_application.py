import tempfile
import threading
import unittest
from pathlib import Path
from unittest.mock import patch

from course_selection.application import run_workbench_application


class _Lock:
    def __init__(self, *_args):
        self.released = False

    def acquire(self):
        return True

    def release(self):
        self.released = True


class _Service:
    def __init__(self, http_started):
        self.http_started = http_started
        self.shell_submitted_after_http_start = False
        self.closed = False

    def submit(self, operation, context):
        self.shell_submitted_after_http_start = self.http_started.is_set()
        return type("Task", (), {"id": "shell"})()

    def wait(self, _identity, _timeout):
        return True

    def inspect(self, _identity):
        return {"state": "succeeded"}

    def close(self):
        self.closed = True


class _Database:
    def __init__(self):
        self.closed = False

    def close(self):
        self.closed = True


class _Server:
    def __init__(self, http_started):
        self.http_started = http_started
        self.closed = False
        self.task_dispatcher = type("Dispatcher", (), {"shutdown": lambda self: None})()

    def run(self):
        self.http_started.set()

    def close(self):
        self.closed = True


class _Response:
    status = 200

    def __enter__(self):
        return self

    def __exit__(self, *_args):
        return False


class ApplicationStartupTests(unittest.TestCase):
    def test_http_server_starts_before_chromium_navigates_to_workbench(self):
        http_started = threading.Event()
        service = _Service(http_started)
        database = _Database()
        server = _Server(http_started)
        app = type("App", (), {"extensions": {
            "observation_service": service,
            "workspace_database": database,
        }})()

        with tempfile.TemporaryDirectory() as directory, (
            patch("course_selection.application.WorkspaceLock", _Lock)
        ), patch(
            "course_selection.application.create_workbench_app", return_value=app
        ), patch(
            "course_selection.application.create_server", return_value=server
        ), patch(
            "course_selection.application.urlopen",
            side_effect=lambda *_args, **_kwargs: (
                _Response() if http_started.wait(1) else self.fail("HTTP server did not start")
            ),
        ):
            run_workbench_application(Path(directory), 5000)

        self.assertTrue(service.shell_submitted_after_http_start)
        self.assertTrue(server.closed)
        self.assertTrue(service.closed)
        self.assertTrue(database.closed)


if __name__ == "__main__":
    unittest.main()
