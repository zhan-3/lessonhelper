"""Loopback-only Flask adapter for the persistent workbench."""

from __future__ import annotations

import secrets
import tempfile
from datetime import datetime, timezone
from pathlib import Path
from typing import Callable

from flask import Flask, abort, jsonify, request, send_from_directory

from .gateway import AcademicGateway, PlaywrightAcademicGateway
from .notice_discovery import candidate_from_text, notice_diff
from .persistence import WorkspaceDatabase
from .tasks import ObservationService
from .timetable import entries_to_dict, import_timetable


def _stale(snapshot: dict | None, seconds: int) -> bool:
    if not snapshot:
        return False
    source_at = datetime.fromisoformat(snapshot["source_at"])
    return (datetime.now(timezone.utc) - source_at).total_seconds() > seconds


def create_workbench_app(root: Path | str = ".private/academic-selection", *, gateway_factory: Callable[[], AcademicGateway] | None = None) -> Flask:
    root = Path(root)
    database = WorkspaceDatabase.open(root)
    if gateway_factory is None:
        gateway_factory = lambda: PlaywrightAcademicGateway(root.parent / "course-progress", root)
    service = ObservationService(database, gateway_factory)
    frontend = Path(__file__).with_name("workbench_static")
    app = Flask(__name__, static_folder=None)
    app.config.update(WORKBENCH_ROOT=root, CSRF_TOKEN=secrets.token_urlsafe(24))
    app.extensions["workspace_database"] = database
    app.extensions["observation_service"] = service

    @app.before_request
    def protect_state_changes():
        if request.method in {"POST", "PUT", "PATCH", "DELETE"}:
            if request.host_url.rstrip("/") != request.headers.get("Origin", ""):
                abort(403)
            if request.headers.get("X-CSRF-Token") != app.config["CSRF_TOKEN"]:
                abort(403)

    @app.get("/api/state")
    def state():
        selection = database.latest_snapshot("selection")
        timetable = database.latest_snapshot("timetable")
        return jsonify({"profile": database.current_profile(), "confirmed_notice": database.confirmed_notice(), "snapshots": {"selection": selection, "timetable": timetable}, "stale": {"selection": _stale(selection, 1800), "timetable": _stale(timetable, 86400)}, "academic_session": {"state": service.session_state}, "csrf_token": app.config["CSRF_TOKEN"]})

    @app.post("/api/tasks")
    def submit_task():
        body = request.get_json(force=True)
        operation = body.get("operation", "")
        context = dict(body.get("context", {}))
        if operation == "refresh-selection":
            profile, notice = database.current_profile(), database.confirmed_notice()
            windows = (notice or {}).get("windows", [])
            grade = str((profile or {}).get("grade", ""))
            allowed = [item for item in windows if item.get("action") == "selection" and item.get("method") == "academic_system" and (not grade or grade in item.get("grades", []))]
            context = {
                "term": (notice or {}).get("term", ""),
                "semester_label": str((notice or {}).get("term", "")).replace("年", "").replace("学期", "").replace(" ", ""),
                "profile_id": (profile or {}).get("version_id"),
                "notice_id": (notice or {}).get("version_id"),
                "allowed_categories": list(dict.fromkeys(code for item in allowed for code in item.get("category_codes", []))),
                "allowed_windows": {code: [item for item in allowed if code in item.get("category_codes", [])] for code in dict.fromkeys(code for item in allowed for code in item.get("category_codes", []))},
            }
        task = service.submit(operation, context)
        return jsonify({"id": task.id, "state": task.state}), 202

    @app.get("/api/tasks/<identity>")
    def inspect_task(identity: str):
        task = service.inspect(identity)
        return (jsonify(task), 200) if task else (jsonify({"error": "not found"}), 404)

    @app.delete("/api/tasks/<identity>")
    def cancel_task(identity: str):
        return (jsonify({"cancelled": True}), 202) if service.cancel(identity) else (jsonify({"cancelled": False}), 409)

    @app.post("/api/notices/candidates")
    def create_notice_candidate():
        body = request.get_json(force=True)
        candidate = candidate_from_text(
            body.get("source_url", ""), body.get("text", ""),
            official_hosts=tuple(app.config.get("OFFICIAL_NOTICE_HOSTS", ("jwc.hitwh.edu.cn",))),
        )
        previous = database.confirmed_notice()
        saved = database.save_notice(candidate)
        return jsonify({"notice": saved, "diff": notice_diff(previous, saved) if previous else ""}), 201

    @app.post("/api/notices/<identity>/confirm")
    def confirm_notice(identity: str):
        try:
            return jsonify(database.confirm_notice(identity))
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
            snapshot = database.publish_snapshot("timetable", entries[0].term, {"status": "complete", "entries": entries_to_dict(entries)}, source="user-imported")
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
