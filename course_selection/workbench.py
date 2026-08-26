"""Loopback-only Flask adapter for the persistent workbench."""

from __future__ import annotations

import secrets
import tempfile
from pathlib import Path
from typing import Callable

from flask import Flask, abort, jsonify, request, send_from_directory

from .gateway import AcademicGateway, PlaywrightAcademicGateway
from .persistence import WorkspaceDatabase
from .tasks import ObservationService
from .timetable import import_timetable, timetable_snapshot_payload
from .workbench_service import NoticeReadError, WorkbenchService


def create_workbench_app(
    root: Path | str = ".private/academic-selection",
    *,
    gateway_factory: Callable[[], AcademicGateway] | None = None,
    frontend_root: Path | str | None = None,
) -> Flask:
    root = Path(root)
    database = WorkspaceDatabase.open(root)
    if gateway_factory is None:
        gateway_factory = lambda: PlaywrightAcademicGateway(root.parent / "course-progress", root)
    service = ObservationService(database, gateway_factory)
    core = WorkbenchService(database)
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
    )
    app.extensions["workspace_database"] = database
    app.extensions["observation_service"] = service
    app.extensions["workbench_service"] = core

    @app.before_request
    def protect_state_changes():
        if request.method in {"POST", "PUT", "PATCH", "DELETE"}:
            if request.host_url.rstrip("/") != request.headers.get("Origin", ""):
                abort(403)
            if request.headers.get("X-CSRF-Token") != app.config["CSRF_TOKEN"]:
                abort(403)

    @app.get("/api/state")
    def state():
        return jsonify(core.state(session_state=service.session_state, csrf_token=app.config["CSRF_TOKEN"]))

    @app.post("/api/tasks")
    def submit_task():
        body = request.get_json(silent=True)
        if not isinstance(body, dict):
            return jsonify({"error": "JSON object required"}), 400
        operation = body.get("operation", "")
        if operation not in {"connect", "refresh-selection", "refresh-timetable"}:
            return jsonify({"error": "unsupported observation operation"}), 400
        raw_context = body.get("context", {})
        if not isinstance(raw_context, dict):
            return jsonify({"error": "context must be an object"}), 400
        context = dict(raw_context)
        if operation == "refresh-selection":
            context = core.refresh_context()
        task = service.submit(operation, context)
        return jsonify({"id": task.id, "state": task.state}), 202

    @app.get("/api/tasks/<identity>")
    def inspect_task(identity: str):
        task = service.inspect(identity)
        return (jsonify(task), 200) if task else (jsonify({"error": "not found"}), 404)

    @app.delete("/api/tasks/<identity>")
    def cancel_task(identity: str):
        return (jsonify({"cancelled": True}), 202) if service.cancel(identity) else (jsonify({"cancelled": False}), 409)

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

    @app.post("/api/notices/<identity>/confirm")
    def confirm_notice(identity: str):
        try:
            return jsonify(core.confirm_notice(identity))
        except ValueError as error:
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
            snapshot = database.publish_snapshot(
                "timetable", entries[0].term, payload, source="user-imported",
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
