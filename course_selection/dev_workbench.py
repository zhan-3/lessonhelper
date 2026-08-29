"""Development supervisor: persistent Chromium plus a hot-reloaded workbench."""

from __future__ import annotations

import os
import signal
import sqlite3
import subprocess
import sys
import time
from pathlib import Path

from playwright.sync_api import sync_playwright

from course_progress.explorer import launch_browser_context, resolve_profile_dir

WATCHED_PACKAGES = (Path("course_selection"), Path("course_progress"))
ACTIVE_TASK_STATES = ("queued", "connecting", "waiting_for_authentication", "reading", "observing", "cancel_requested")


def source_mtimes(root: Path) -> dict[Path, int]:
    """Return modification timestamps for Python source, excluding private data."""
    return {
        path: path.stat().st_mtime_ns
        for package in WATCHED_PACKAGES
        for path in (root / package).rglob("*.py")
        if path.is_file()
    }


def has_active_tasks(workspace_root: Path) -> bool:
    database = workspace_root / "workbench.sqlite3"
    if not database.is_file():
        return False
    try:
        connection = sqlite3.connect(f"file:{database}?mode=ro", uri=True)
        try:
            placeholders = ",".join("?" for _ in ACTIVE_TASK_STATES)
            return connection.execute(
                f"select 1 from observation_tasks where state in ({placeholders}) "
                f"union all select 1 from execution_tasks where state in ({placeholders}) limit 1",
                (*ACTIVE_TASK_STATES, *ACTIVE_TASK_STATES),
            ).fetchone() is not None
        finally:
            connection.close()
    except sqlite3.Error:
        return True


def start_workbench(root: Path, workspace_root: Path, port: int, cdp_url: str) -> subprocess.Popen:
    environment = {
        **os.environ,
        "ACADEMIC_BROWSER_CDP_URL": cdp_url,
        "ACADEMIC_WORKBENCH_DEV_DIAGNOSTICS": "1",
    }
    return subprocess.Popen(
        [
            sys.executable, "-m", "course_selection", "workbench",
            "--private-root", str(workspace_root), "--port", str(port),
        ],
        cwd=root,
        env=environment,
        creationflags=(subprocess.CREATE_NEW_PROCESS_GROUP if os.name == "nt" else 0),
    )


def stop_workbench(process: subprocess.Popen) -> None:
    if process.poll() is not None:
        return
    process.send_signal(signal.CTRL_BREAK_EVENT if os.name == "nt" else signal.SIGINT)
    try:
        process.wait(timeout=10)
    except subprocess.TimeoutExpired:
        process.kill()
        process.wait(timeout=5)


def run_dev_workbench(root: Path, workspace_root: Path, port: int, debug_port: int = 9222) -> int:
    """Keep Chromium/CDP alive while restarting only the Python workbench."""
    root, workspace_root = root.resolve(), workspace_root.resolve()
    cdp_url = f"http://127.0.0.1:{debug_port}"
    original_debug_port = os.environ.get("ACADEMIC_BROWSER_DEBUG_PORT")
    os.environ["ACADEMIC_BROWSER_DEBUG_PORT"] = str(debug_port)
    child: subprocess.Popen | None = None
    try:
        with sync_playwright() as playwright:
            # Browser Host ownership intentionally lives here, not in a child.
            with launch_browser_context(
                playwright, "chromium", resolve_profile_dir(root / ".private" / "course-progress")
            ):
                child = start_workbench(root, workspace_root, port, cdp_url)
                known = source_mtimes(root)
                restart_pending = False
                print(f"dev workbench: CDP available at {cdp_url}")
                while True:
                    time.sleep(0.5)
                    current = source_mtimes(root)
                    restart_pending = restart_pending or current != known
                    known = current
                    if child.poll() is not None:
                        child = start_workbench(root, workspace_root, port, cdp_url)
                        restart_pending = False
                    if restart_pending and not has_active_tasks(workspace_root):
                        print("dev workbench: Python sources changed; restarting workbench only")
                        stop_workbench(child)
                        child = start_workbench(root, workspace_root, port, cdp_url)
                        restart_pending = False
    except KeyboardInterrupt:
        return 0
    finally:
        if child is not None:
            stop_workbench(child)
        if original_debug_port is None:
            os.environ.pop("ACADEMIC_BROWSER_DEBUG_PORT", None)
        else:
            os.environ["ACADEMIC_BROWSER_DEBUG_PORT"] = original_debug_port
