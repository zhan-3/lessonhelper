"""One-click local workbench hosted inside visible Playwright Chromium."""

from __future__ import annotations

import json
import logging
import threading
import time
from pathlib import Path
from urllib.request import Request, urlopen

from waitress import create_server

from . import config
from .gateway import PlaywrightAcademicGateway
from .single_instance import WorkspaceLock
from .workbench import create_workbench_app

logger = logging.getLogger("course-selection.application")


def activate_running_workbench(url: str) -> bool:
    """Ask an existing loopback workbench to bring its Chromium shell forward."""
    try:
        with urlopen(f"{url}/api/state", timeout=2) as response:
            token = json.loads(response.read().decode("utf-8"))["csrf_token"]
        request = Request(
            f"{url}/api/shell/activate",
            data=b"{}",
            method="POST",
            headers={
                "Content-Type": "application/json",
                "Origin": url,
                "X-CSRF-Token": token,
            },
        )
        with urlopen(request, timeout=2) as response:
            return response.status == 202
    except (OSError, ValueError, KeyError):
        return False


def _wait_until_serving(url: str, server_thread: threading.Thread, timeout: float = 5) -> bool:
    """Wait until the loopback HTTP server can answer before opening Chromium."""
    deadline = time.monotonic() + timeout
    while time.monotonic() < deadline:
        try:
            with urlopen(f"{url}/api/state", timeout=0.5) as response:
                if response.status == 200:
                    return True
        except OSError:
            if not server_thread.is_alive():
                return False
            time.sleep(0.05)
    return False


def run_workbench_application(root: Path, port: int) -> int:
    """Run Flask and its single visible bundled-Chromium application shell."""
    root = root.resolve()
    url = f"http://127.0.0.1:{port}"
    lock = WorkspaceLock(root, port)
    if not lock.acquire():
        return 0 if activate_running_workbench(url) else 1

    gateway_factory = lambda: PlaywrightAcademicGateway(
        root.parent / "course-progress",
        root,
        cdp_url=config.ACADEMIC_BROWSER_CDP_URL,
    )
    app = create_workbench_app(
        root,
        gateway_factory=gateway_factory,
        workbench_url=url,
        require_login_configuration=True,
    )
    service = app.extensions["observation_service"]
    database = app.extensions["workspace_database"]

    server = None
    server_thread = None
    try:
        server = create_server(app, host="127.0.0.1", port=port)
        server_thread = threading.Thread(
            target=server.run,
            name="local-workbench-http",
            daemon=True,
        )
        server_thread.start()
        if not _wait_until_serving(url, server_thread):
            logger.error("workbench HTTP server did not start within 5 seconds")
            return 1

        # Offline startup: opening the local shell must not authenticate or
        # contact any university endpoint. Remote work begins only after an
        # explicit user action.
        shell = service.submit("launch-shell", {"workbench_url": url})
        if not service.wait(shell.id, 30):
            logger.error("visible Chromium workbench did not start within 30 seconds")
            return 1
        result = service.inspect(shell.id) or {}
        if result.get("state") != "succeeded":
            logger.error("visible Chromium workbench failed to start: %s", result.get("error"))
            return 1

        logger.info("Workbench serving at %s (waitress)", url)
        while server_thread.is_alive():
            server_thread.join(timeout=1)
        return 0
    except KeyboardInterrupt:
        return 0
    finally:
        if server is not None:
            server.close()
            server.task_dispatcher.shutdown()
        if server_thread is not None:
            server_thread.join(timeout=5)
        service.close()
        database.close()
        lock.release()