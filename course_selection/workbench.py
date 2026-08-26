"""Loopback-only Flask adapter for the persistent workbench."""

from __future__ import annotations

import secrets
import json
import tempfile
from datetime import datetime, timezone
from pathlib import Path
from typing import Callable

from flask import Flask, abort, jsonify, request, send_from_directory

from .gateway import AcademicGateway, PlaywrightAcademicGateway
from .notice_discovery import candidate_from_text, notice_diff
from .notice import fetch_notice_text
from .persistence import WorkspaceDatabase
from .tasks import ObservationService
from .timetable import import_timetable, timetable_snapshot_payload


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
        body = request.get_json(silent=True)
        if not isinstance(body, dict):
            return jsonify({"error": "JSON object required"}), 400
        source_url = str(body.get("source_url", "")).strip()
        text = str(body.get("text", "")).strip()
        if not text and source_url:
            try:
                text = fetch_notice_text(source_url)
            except (OSError, ValueError) as error:
                return jsonify({"error": f"unable to read notice: {error}"}), 400
        if not source_url or not text:
            return jsonify({"error": "source_url and text are required"}), 400
        try:
            candidate = candidate_from_text(
                source_url, text,
                official_hosts=tuple(app.config.get("OFFICIAL_NOTICE_HOSTS", ("jwc.hitwh.edu.cn",))),
            )
        except ValueError as error:
            return jsonify({"error": str(error)}), 422
        previous = database.confirmed_notice()
        saved = database.save_notice(candidate)
        return jsonify({"notice": saved, "diff": notice_diff(previous, saved) if previous else ""}), 201

    @app.get("/api/notices/candidates")
    def list_notice_candidates():
        rows = database.connection.execute(
            "select id,payload,status,created_at from notice_versions order by created_at desc"
        ).fetchall()
        notices = []
        for row in rows:
            payload = json.loads(row["payload"])
            payload.update(version_id=row["id"], status=row["status"], created_at=row["created_at"])
            notices.append(payload)
        return jsonify({"notices": notices})

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
