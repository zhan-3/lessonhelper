"""Framework-independent application services for the local workbench.

This module contains the decisions made by the workbench use cases.  Flask
routes should only translate HTTP input/output and delegate here.
"""

from __future__ import annotations

import json
from datetime import datetime, timezone
from typing import Any

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

    def __init__(self, database: WorkspaceDatabase, *, official_notice_hosts: tuple[str, ...] = ("jwc.hitwh.edu.cn",)):
        self.database = database
        self.official_notice_hosts = official_notice_hosts

    def state(self, *, session_state: str, csrf_token: str | None = None) -> dict[str, Any]:
        selection = self.database.latest_snapshot("selection")
        timetable = self.database.latest_snapshot("timetable")
        result: dict[str, Any] = {
            "profile": self.database.current_profile(),
            "confirmed_notice": self.database.confirmed_notice(),
            "snapshots": {"selection": selection, "timetable": timetable},
            "latest_plan": self.database.latest_plan(),
            "stale": {"selection": is_stale(selection, 1800), "timetable": is_stale(timetable, 86400)},
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
