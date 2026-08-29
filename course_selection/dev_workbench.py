"""Development supervisor: persistent Chromium plus a hot-reloaded workbench."""

from __future__ import annotations

import logging
import os
import signal
import sqlite3
import subprocess
import sys
import threading
import time
from pathlib import Path

from playwright.sync_api import sync_playwright
from watchfiles import PythonFilter, watch

from course_progress.explorer import launch_browser_context, resolve_profile_dir

logger = logging.getLogger("course-selection.dev-workbench")

WATCHED_PACKAGES = ("course_selection", "course_progress")
ACTIVE_TASK_STATES = ("queued", "connecting", "waiting_for_authentication", "reading", "observing", "cancel_requested")


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
    # Event-driven source watching: a background thread waits on OS-level file
    # notifications (watchfiles) and only signals Python-file changes.  The
    # main loop stays cheap — it just polls child status and the flag.
    restart_pending = threading.Event()
    stop_watching = threading.Event()

    def observe_changes() -> None:
        try:
            for _changes in watch(
                *(root / package for package in WATCHED_PACKAGES),
                watch_filter=PythonFilter(),
                stop_event=stop_watching,
                debounce=400,
                step=50,
            ):
                restart_pending.set()
        except OSError as error:
            logger.warning("source watcher stopped: %s", error)

    watcher = threading.Thread(target=observe_changes, name="source-watcher", daemon=True)
    watcher.start()
    try:
        with sync_playwright() as playwright, launch_browser_context(
            # Browser Host ownership intentionally lives here, not in a child.
            playwright, "chromium", resolve_profile_dir(root / ".private" / "course-progress")
        ):
            child = start_workbench(root, workspace_root, port, cdp_url)
            logger.info("dev workbench: CDP available at %s", cdp_url)
            while True:
                time.sleep(0.5)
                if child.poll() is not None:
                    # A fresh child picks up current sources anyway.
                    child = start_workbench(root, workspace_root, port, cdp_url)
                    restart_pending.clear()
                    continue
                if restart_pending.is_set() and not has_active_tasks(workspace_root):
                    restart_pending.clear()
                    logger.info("Python sources changed; restarting workbench only")
                    stop_workbench(child)
                    child = start_workbench(root, workspace_root, port, cdp_url)
    except KeyboardInterrupt:
        return 0
    finally:
        stop_watching.set()
        if child is not None:
            stop_workbench(child)
        if original_debug_port is None:
            os.environ.pop("ACADEMIC_BROWSER_DEBUG_PORT", None)
        else:
            os.environ["ACADEMIC_BROWSER_DEBUG_PORT"] = original_debug_port
