"""Framework-independent application services for the local workbench.

This module contains the decisions made by the workbench use cases.  Flask
routes should only translate HTTP input/output and delegate here.
"""

from __future__ import annotations

import json
from datetime import datetime, timezone
from pathlib import Path
from typing import Any

from course_progress.credentials import credential_store

from .notice import fetch_notice_text
from .notice_discovery import candidate_from_text, notice_diff
from .persistence import WorkspaceDatabase
from .planning import ReadOnlyPlan, build_read_only_plan


class NoticeReadError(ValueError):
    """The source URL could not be read before candidate parsing."""


def is_stale(snapshot: dict[str, Any] | None, seconds: int) -> bool:
    if not snapshot:
        return False
    source_at = datetime.fromisoformat(snapshot["source_at"])
    return (datetime.now(timezone.utc) - source_at).total_seconds() > seconds


class WorkbenchService:
    """Use cases shared by HTTP and other possible adapters."""

    def __init__(
        self,
        database: WorkspaceDatabase,
        *,
        official_notice_hosts: tuple[str, ...] = ("jwc.hitwh.edu.cn",),
        progress_report_path: Path | str | None = None,
        login_root: Path | str | None = None,
    ):
        self.database = database
        self.official_notice_hosts = official_notice_hosts
        self.progress_report_path = Path(progress_report_path) if progress_report_path else None
        self.login_root = Path(login_root) if login_root else database.root.parent / "course-progress"

    def login_configuration(self) -> dict[str, Any]:
        """Expose non-secret configuration state without decrypting the password."""
        store = credential_store(self.login_root)
        if not store.path.is_file():
            return {"state": "missing", "configured": False}
        metadata_path = self.login_root / "webvpn-login-meta.json"
        try:
            metadata = json.loads(metadata_path.read_text(encoding="utf-8"))
        except (OSError, UnicodeError, json.JSONDecodeError):
            metadata = {}
        result: dict[str, Any] = {"state": "configured", "configured": True}
        masked = metadata.get("masked_username")
        if isinstance(masked, str) and masked:
            result["masked_username"] = masked
        return result

    def configure_login(self, username: str, password: str) -> dict[str, Any]:
        username = username.strip()
        if not username or not password:
            raise ValueError("学号和密码不能为空")
        if len(username) > 64 or len(password) > 256:
            raise ValueError("登录信息长度无效")
        from course_progress.credentials import LoginCredentials

        store = credential_store(self.login_root)
        existing = store.load()
        if existing is not None and existing.username != username:
            raise ValueError("更换学号前请先清除当前登录并重置个人工作区")
        store.save(LoginCredentials(username, password))
        masked = f"{username[:4]}{'*' * max(2, len(username) - 4)}"
        metadata_path = self.login_root / "webvpn-login-meta.json"
        metadata_path.parent.mkdir(parents=True, exist_ok=True)
        temporary = metadata_path.with_suffix(".json.tmp")
        temporary.write_text(
            json.dumps({"masked_username": masked}, ensure_ascii=False),
            encoding="utf-8",
        )
        temporary.replace(metadata_path)
        # A storage-state export may belong to the previous credentials.  The
        # persistent browser profile is validated again by the next task.
        auth_state = self.login_root / "webvpn-auth-state.json"
        auth_state.unlink(missing_ok=True)
        return self.login_configuration()

    def clear_login(self) -> None:
        store = credential_store(self.login_root)
        store.path.unlink(missing_ok=True)
        (self.login_root / "webvpn-login-meta.json").unlink(missing_ok=True)
        (self.login_root / "webvpn-auth-state.json").unlink(missing_ok=True)
        for name in ("progress-report.json", "collection-checkpoint.json"):
            (self.login_root / name).unlink(missing_ok=True)
        self.database.reset_personal_workspace()

    def graduation_progress(self) -> dict[str, Any]:
        """Read the latest score-free progress snapshot, with legacy fallback."""
        snapshot = self.database.latest_snapshot("progress")
        if snapshot is not None:
            payload = snapshot.get("payload", {})
            report = payload.get("report") if isinstance(payload, dict) else None
            if not isinstance(report, dict) or not isinstance(report.get("progress"), list):
                return {"status": "invalid", "report": None, "snapshot": snapshot}
            profile = self.database.current_profile()
            profile_id = (profile or {}).get("version_id")
            if profile_id and snapshot.get("profile_id") != profile_id:
                return {"status": "not_applicable", "report": None, "snapshot": snapshot}
            if report.get("baseline_version") != "guide-2026":
                return {"status": "not_applicable", "report": None, "snapshot": snapshot}
            status = "ready" if report.get("data_complete") is True else "incomplete"
            return {"status": status, "report": report, "snapshot": snapshot}

        # Keep reports generated by the existing standalone collector visible
        # until the user performs the first workbench refresh.
        path = self.progress_report_path
        if path is None or not path.is_file():
            return {"status": "missing", "report": None}
        try:
            report = json.loads(path.read_text(encoding="utf-8"))
        except (OSError, UnicodeError, json.JSONDecodeError):
            return {"status": "invalid", "report": None}
        if not isinstance(report, dict) or not isinstance(report.get("progress"), list):
            return {"status": "invalid", "report": None}
        status = "ready" if report.get("data_complete") is True else "incomplete"
        return {"status": status, "report": report}

    def progress_context(self) -> dict[str, Any]:
        profile = self.database.current_profile() or {}
        notice = self.database.confirmed_notice() or {}
        timetable = self.database.latest_snapshot("timetable") or {}
        return {
            "term": str(notice.get("term") or timetable.get("term") or ""),
            "profile_id": profile.get("version_id"),
            "baseline_version": "guide-2026",
            "page_size": 20,
        }

    def state(self, *, session_state: str, csrf_token: str | None = None) -> dict[str, Any]:
        selection = self.database.latest_snapshot("selection")
        timetable = self.database.latest_snapshot("timetable")
        progress_snapshot = self.database.latest_snapshot("progress")
        result: dict[str, Any] = {
            "login_configuration": self.login_configuration(),
            "profile": self.database.current_profile(),
            "confirmed_notice": self.database.confirmed_notice(),
            "snapshots": {"selection": selection, "timetable": timetable, "progress": progress_snapshot},
            "snapshot_changes": {
                "selection": self.database.latest_snapshot_change("selection"),
                "timetable": self.database.latest_snapshot_change("timetable"),
                "progress": self.database.latest_snapshot_change("progress"),
            },
            "latest_plan": self.database.latest_plan(),
            "graduation_progress": self.graduation_progress(),
            "stale": {"selection": is_stale(selection, 1800), "timetable": is_stale(timetable, 86400), "progress": is_stale(progress_snapshot, 86400)},
            "academic_session": {"state": session_state},
        }
        if csrf_token is not None:
            result["csrf_token"] = csrf_token
        return result

    def refresh_context(self) -> dict[str, Any]:
        profile, notice = self.database.current_profile(), self.database.confirmed_notice()
        windows = (notice or {}).get("windows", [])
        grade = str((profile or {}).get("grade", ""))
        allowed = [item for item in windows if item.get("action") == "selection" and item.get("method") == "academic_system" and grade and grade in item.get("grades", [])]
        categories = list(dict.fromkeys(code for item in allowed for code in item.get("category_codes", [])))
        return {
            "term": (notice or {}).get("term", ""),
            "semester_label": str((notice or {}).get("term", "")).replace("年", "").replace("学期", "").replace(" ", ""),
            "profile_id": (profile or {}).get("version_id"),
            "notice_id": (notice or {}).get("version_id"),
            "allowed_categories": categories,
            "allowed_windows": {code: [item for item in allowed if code in item.get("category_codes", [])] for code in categories},
        }

    def create_notice_candidate(self, source_url: str, text: str) -> tuple[dict[str, Any], str]:
        if not text and source_url:
            try:
                text = fetch_notice_text(source_url)
            except (OSError, ValueError) as error:
                raise NoticeReadError(f"unable to read notice: {error}") from error
        if not source_url or not text:
            raise ValueError("source_url and text are required")
        candidate = candidate_from_text(source_url, text, official_hosts=self.official_notice_hosts)
        previous = self.database.confirmed_notice()
        saved = self.database.save_notice(candidate)
        return saved, notice_diff(previous, saved) if previous else ""

    def list_notice_candidates(self) -> list[dict[str, Any]]:
        rows = self.database.connection.execute("select id,payload,status,created_at from notice_versions order by created_at desc").fetchall()
        notices = []
        for row in rows:
            payload = json.loads(row["payload"])
            payload.update(version_id=row["id"], status=row["status"], created_at=row["created_at"])
            notices.append(payload)
        return notices

    def confirm_notice(self, identity: str) -> dict[str, Any]:
        return self.database.confirm_notice(identity)

    def build_plan(self, goals: list[dict[str, Any]]) -> ReadOnlyPlan:
        """Build a local plan from the latest applicable snapshots.

        This service method intentionally has no gateway dependency; callers
        can expose it through HTTP, CLI, or another adapter without granting
        the planner any academic write capability.
        """
        profile = self.database.current_profile() or {}
        notice = self.database.confirmed_notice() or {}
        timetable = self.database.latest_snapshot("timetable")
        selection = self.database.latest_snapshot("selection")
        term = str(notice.get("term") or (timetable or {}).get("term") or "")
        return build_read_only_plan(
            term=term, profile_id=str(profile.get("version_id", "")),
            notice_id=str(notice.get("version_id", "")), timetable_snapshot=timetable,
            selection_snapshot=selection, goals=goals,
        )

    def save_plan(self, goals: list[dict[str, Any]]) -> dict[str, Any]:
        """Persist a read-only plan; this method has no gateway or write path."""
        return self.database.save_plan(self.build_plan(goals).to_dict())
