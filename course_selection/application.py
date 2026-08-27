"""One-click local workbench hosted inside visible Playwright Chromium."""

from __future__ import annotations

import json
import threading
from pathlib import Path
from urllib.request import Request, urlopen

from werkzeug.serving import make_server

from .gateway import PlaywrightAcademicGateway
from .single_instance import WorkspaceLock
from .workbench import create_workbench_app


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


def run_workbench_application(root: Path, port: int) -> int:
    """Run Flask and its single visible bundled-Chromium application shell."""
    root = root.resolve()
    url = f"http://127.0.0.1:{port}"
    lock = WorkspaceLock(root, port)
    if not lock.acquire():
        return 0 if activate_running_workbench(url) else 1

    browser_closed = threading.Event()
    gateway_factory = lambda: PlaywrightAcademicGateway(
        root.parent / "course-progress",
        root,
        on_browser_closed=browser_closed.set,
    )
    app = create_workbench_app(
        root,
        gateway_factory=gateway_factory,
        workbench_url=url,
        require_login_configuration=True,
    )
    service = app.extensions["observation_service"]
    database = app.extensions["workspace_database"]
    server = make_server("127.0.0.1", port, app, threaded=True)
    server_thread = threading.Thread(
        target=server.serve_forever,
        name="local-workbench-http",
        daemon=True,
    )
    server_thread.start()
    try:
        shell = service.submit("launch-shell", {"workbench_url": url})
        if not service.wait(shell.id, 30):
            raise TimeoutError("visible Chromium workbench did not start within 30 seconds")
        result = service.inspect(shell.id) or {}
        if result.get("state") != "succeeded":
            raise RuntimeError(result.get("error") or "visible Chromium workbench failed to start")
        browser_closed.wait()
        return 0
    finally:
        server.shutdown()
        server_thread.join(timeout=5)
        service.close()
        database.close()
        lock.release()
