"""Loopback-only Flask adapter for the persistent workbench."""

from __future__ import annotations

import os
import secrets
import tempfile
from collections.abc import Callable
from pathlib import Path
from urllib.parse import urlsplit

from flask import Flask, abort, jsonify, request, send_from_directory

from .gateway import AcademicGateway, PlaywrightAcademicGateway
from .notice_discovery import DEFAULT_NOTICE_INDEX_URL
from .persistence import WorkspaceDatabase
from .tasks import ObservationService
from .timetable import import_timetable, timetable_snapshot_payload
from .workbench_service import NoticeReadError, WorkbenchService


def create_workbench_app(
    root: Path | str = ".private/academic-selection",
    *,
    gateway_factory: Callable[[], AcademicGateway] | None = None,
    frontend_root: Path | str | None = None,
    workbench_url: str = "http://127.0.0.1:5000",
    login_root: Path | str | None = None,
    require_login_configuration: bool = False,
) -> Flask:
    root = Path(root)
    resolved_login_root = Path(login_root) if login_root else (
        root.parent / "course-progress" if root.name == "academic-selection" else root / "course-progress"
    )
    database = WorkspaceDatabase.open(root)
    if gateway_factory is None:
        cdp_url = os.environ.get("ACADEMIC_BROWSER_CDP_URL") or None
        gateway_factory = lambda: PlaywrightAcademicGateway(
            root.parent / "course-progress", root, cdp_url=cdp_url,
        )
    service = ObservationService(database, gateway_factory)
    core = WorkbenchService(
        database,
        progress_report_path=resolved_login_root / "progress-report.json",
        login_root=resolved_login_root,
    )
    # The Vite config writes its production bundle here.  Keeping the bundle
    # beside the Python package makes the same Flask entry point work from a
    # checkout and from an installed wheel.  Tests and embedders may provide a
    # different directory without changing application code.
    frontend = Path(frontend_root) if frontend_root is not None else Path(__file__).with_name("workbench_static")
    frontend = frontend.resolve()
    app = Flask(__name__, static_folder=None)
    app.config.update(
        WORKBENCH_ROOT=root,
        WORKBENCH_FRONTEND=frontend,
        CSRF_TOKEN=secrets.token_urlsafe(24),
        WORKBENCH_URL=workbench_url.rstrip("/"),
        REQUIRE_LOGIN_CONFIGURATION=require_login_configuration,
    )
    app.extensions["workspace_database"] = database
    app.extensions["observation_service"] = service
    app.extensions["workbench_service"] = core

    @app.before_request
    def protect_local_service():
        host = urlsplit(f"http://{request.host}")
        origin = urlsplit(request.headers.get("Origin", ""))
        loopback = {"127.0.0.1", "localhost", "::1"}
        if (host.hostname or "").lower() not in loopback:
            abort(403)
        if request.method in {"POST", "PUT", "PATCH", "DELETE"}:
            if (
                origin.scheme != "http"
                or (origin.hostname or "").lower() not in loopback
                or origin.port != host.port
            ):
                abort(403)
            if request.headers.get("X-CSRF-Token") != app.config["CSRF_TOKEN"]:
                abort(403)

    @app.after_request
    def harden_local_responses(response):
        response.headers["Content-Security-Policy"] = (
            "default-src 'self'; script-src 'self'; style-src 'self'; "
            "img-src 'self' data:; connect-src 'self'; frame-ancestors 'none'; "
            "form-action 'self'; base-uri 'none'"
        )
        response.headers["Referrer-Policy"] = "no-referrer"
        response.headers["X-Content-Type-Options"] = "nosniff"
        response.headers["X-Frame-Options"] = "DENY"
        if request.path.startswith("/api/"):
            response.headers["Cache-Control"] = "no-store"
        return response

    @app.get("/api/state")
    def state():
        return jsonify(core.state(
            session_state=service.session_status(), active_task=service.active_task(),
            csrf_token=app.config["CSRF_TOKEN"],
        ))

    @app.post("/api/login-configuration")
    def configure_login():
        body = request.get_json(silent=True)
        if not isinstance(body, dict):
            return jsonify({"error": "JSON object required"}), 400
        try:
            def configure_and_restart():
                result = core.configure_login(
                    str(body.get("username", "")),
                    str(body.get("password", "")),
                )
                return result

            result = service.run_when_idle(configure_and_restart)
        except (OSError, RuntimeError, ValueError) as error:
            return jsonify({"error": str(error)}), 400
        return jsonify(result), 201

    @app.delete("/api/login-configuration")
    def clear_login():
        try:
            def clear_and_restart():
                core.clear_login()
                service.submit("reset-login")

            service.run_when_idle(clear_and_restart)
        except (OSError, RuntimeError) as error:
            return jsonify({"error": str(error)}), 409
        return "", 204

    @app.post("/api/tasks")
    def submit_task():
        body = request.get_json(silent=True)
        if not isinstance(body, dict):
            return jsonify({"error": "JSON object required"}), 400
        operation = body.get("operation", "")
        if operation not in {"connect", "refresh-selection", "refresh-timetable", "refresh-progress", "observe-navigation"}:
            return jsonify({"error": "unsupported observation operation"}), 400
        if app.config["REQUIRE_LOGIN_CONFIGURATION"] and not core.login_configuration().get("configured"):
            return jsonify({"error": "请先配置本机自动登录"}), 409
        raw_context = body.get("context", {})
        if not isinstance(raw_context, dict):
            return jsonify({"error": "context must be an object"}), 400
        context = dict(raw_context)
        if operation in {"connect", "refresh-selection", "refresh-timetable"}:
            context = core.refresh_context()
        elif operation == "refresh-progress":
            context = core.progress_context()
        try:
            task = service.submit(operation, context)
        except RuntimeError as error:
            return jsonify({"error": str(error), "active_task": service.active_task()}), 409
        return jsonify({"id": task.id, "operation": operation, "state": task.state}), 202

    @app.get("/api/tasks/<identity>")
    def inspect_task(identity: str):
        task = service.inspect(identity)
        return (jsonify(task), 200) if task else (jsonify({"error": "not found"}), 404)

    @app.post("/api/shell/activate")
    def activate_shell():
        task = service.submit(
            "launch-shell",
            {"workbench_url": app.config["WORKBENCH_URL"]},
        )
        return jsonify({"id": task.id, "operation": "launch-shell", "state": task.state}), 202

    @app.delete("/api/tasks/<identity>")
    def cancel_task(identity: str):
        return (jsonify({"cancelled": True}), 202) if service.cancel(identity) else (jsonify({"cancelled": False}), 409)

    @app.post("/api/tasks/<identity>/finish")
    def finish_task(identity: str):
        return (jsonify({"finished": True}), 202) if service.finish(identity) else (jsonify({"finished": False}), 409)

    @app.post("/api/plans")
    def save_plan():
        body = request.get_json(silent=True)
        if not isinstance(body, dict) or not isinstance(body.get("goals"), list):
            return jsonify({"error": "goals must be an array"}), 400
        if not all(isinstance(goal, dict) for goal in body["goals"]):
            return jsonify({"error": "each goal must be an object"}), 400
        return jsonify(core.save_plan(body["goals"])), 201

    @app.get("/api/plans/latest")
    def latest_plan():
        plan = database.latest_plan()
        return (jsonify(plan), 200) if plan else (jsonify({"error": "not found"}), 404)

    @app.post("/api/executions/selection")
    def execute_selection():
        body = request.get_json(silent=True)
        if not isinstance(body, dict):
            return jsonify({"error": "JSON object required"}), 400
        section_id = str(body.get("section_id") or "")
        snapshot_id = str(body.get("snapshot_id") or "")
        # Bind the one-time confirmation to the exact concrete teaching section.
        if not section_id or body.get("confirmation") != section_id:
            return jsonify({"error": "必须明确确认当前教学班"}), 409
        try:
            def validate_and_submit():
                context = core.prepare_selection_execution(section_id, snapshot_id)
                return service.submit_execution(context)

            task = service.run_when_idle(validate_and_submit)
        except (RuntimeError, ValueError) as error:
            return jsonify({"error": str(error)}), 409
        return jsonify({
            "id": task.id, "operation": "execute-selection",
            "task_kind": "execution", "state": task.state,
        }), 202

    @app.get("/api/executions")
    def execution_history():
        return jsonify({"executions": database.execution_history()})

    @app.delete("/api/executions")
    def clear_execution_history():
        database.clear_execution_history()
        return "", 204

    @app.post("/api/executions/<identity>/resolve")
    def resolve_execution(identity: str):
        return (jsonify({"resolved": True}), 200) if database.resolve_execution(identity) else (jsonify({"error": "not found"}), 404)

    @app.post("/api/notices/candidates")
    def create_notice_candidate():
        body = request.get_json(silent=True)
        if not isinstance(body, dict):
            return jsonify({"error": "JSON object required"}), 400
        source_url = str(body.get("source_url", "")).strip()
        text = str(body.get("text", "")).strip()
        try:
            saved, diff = core.create_notice_candidate(source_url, text)
        except NoticeReadError as error:
            return jsonify({"error": str(error)}), 400
        except ValueError as error:
            return jsonify({"error": str(error)}), 422
        return jsonify({"notice": saved, "diff": diff}), 201

    @app.get("/api/notices/candidates")
    def list_notice_candidates():
        return jsonify({"notices": core.list_notice_candidates()})

    @app.post("/api/notices/discover")
    def discover_notice_candidates():
        body = request.get_json(silent=True) or {}
        if not isinstance(body, dict):
            return jsonify({"error": "JSON object required"}), 400
        index_url = str(body.get("index_url", "")).strip()
        try:
            notices = core.discover_notice_candidates(index_url or DEFAULT_NOTICE_INDEX_URL)
        except NoticeReadError as error:
            return jsonify({"error": str(error)}), 400
        return jsonify({"notices": notices}), 201

    @app.post("/api/notices/<identity>/confirm")
    def confirm_notice(identity: str):
        try:
            notice = service.run_when_idle(lambda: core.confirm_notice(identity))
            return jsonify(notice)
        except (RuntimeError, ValueError) as error:
            return jsonify({"error": str(error)}), 409

    @app.post("/api/timetable/import")
    def import_timetable_snapshot():
        upload = request.files.get("timetable")
        if upload is None or Path(upload.filename or "").suffix.lower() not in {".xls", ".xlsx"}:
            return jsonify({"error": "timetable must be XLS or XLSX"}), 400
        suffix = Path(upload.filename).suffix.lower()
        with tempfile.NamedTemporaryFile(suffix=suffix, delete=False) as stream:
            temporary = Path(stream.name)
        try:
            upload.save(temporary)
            expected = request.form.get("term") or None
            entries = import_timetable(temporary, expected_term=expected)
            payload = timetable_snapshot_payload(
                entries, source_name=upload.filename or "timetable.xlsx",
            )
            profile = database.current_profile() or {}
            snapshot = database.publish_snapshot(
                "timetable", entries[0].term, payload, source="user-imported",
                profile_id=profile.get("version_id"),
            )
            return jsonify(snapshot), 201
        finally:
            temporary.unlink(missing_ok=True)

    @app.get("/")
    def index():
        if (frontend / "index.html").is_file():
            return send_from_directory(frontend, "index.html")
        return "Academic workbench frontend has not been built", 503

    @app.get("/assets/<path:name>")
    def assets(name: str):
        return send_from_directory(frontend / "assets", name)

    return app
