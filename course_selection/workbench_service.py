"""Framework-independent application services for the local workbench.

This module contains the decisions made by the workbench use cases.  Flask
routes should only translate HTTP input/output and delegate here.

The module never imports a transport: reading notices over HTTP or driving a
browser is injected as ``notice_fetcher`` / ``notice_discoverer`` so that use
cases can be exercised with plain functions instead of patched network calls.
"""

from __future__ import annotations

import json
import math
import re
from collections.abc import Callable
from copy import deepcopy
from datetime import date, datetime, timezone
from pathlib import Path
from typing import Any

from course_progress.baselines import requirement_baseline, requirement_baselines
from course_progress.credentials import credential_store
from course_progress.progress import (
    apply_course_label_estimates,
    apply_projected_course_estimates,
    apply_recognized_credit_estimates,
)

from .notice_discovery import (
    DEFAULT_NOTICE_INDEX_URL,
    candidate_from_text,
    notice_diff,
)
from .persistence import WorkspaceDatabase
from .planning import ReadOnlyPlan, build_read_only_plan

# Injected transports.  ``notice_fetcher`` turns a notice URL into its visible
# text; ``notice_discoverer`` turns an index URL into candidate payloads.
NoticeFetcher = Callable[[str], str]
NoticeDiscoverer = Callable[[str], list[dict[str, Any]]]

_RECOGNIZED_CREDIT_CATEGORIES = frozenset({
    "innovation", "social_practice", "cultural_quality",
})
_LOCAL_IDENTITY = re.compile(r"^[A-Za-z0-9][A-Za-z0-9._:-]{0,63}$")


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
        notice_fetcher: NoticeFetcher | None = None,
        notice_discoverer: NoticeDiscoverer | None = None,
    ):
        self.database = database
        self.official_notice_hosts = official_notice_hosts
        self.progress_report_path = Path(progress_report_path) if progress_report_path else None
        self.login_root = Path(login_root) if login_root else database.root.parent / "course-progress"
        self.notice_fetcher = notice_fetcher
        self.notice_discoverer = notice_discoverer

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

    def available_requirement_baselines(self) -> list[dict[str, Any]]:
        return requirement_baselines()

    def selected_requirement_baseline(self) -> dict[str, Any] | None:
        selection = self.database.requirement_baseline_selection()
        if selection is None:
            return None
        baseline = requirement_baseline(selection["version"])
        return {**baseline, **selection} if baseline is not None else None

    def select_requirement_baseline(
        self, version: str, confirmation: str
    ) -> dict[str, Any]:
        if not version or confirmation != version:
            raise ValueError("必须明确确认要求基线版本")
        baseline = requirement_baseline(version)
        if baseline is None:
            raise ValueError("unknown requirement baseline")
        selection = self.database.select_requirement_baseline(version)
        return {**baseline, **selection}

    def recognized_credits(self) -> list[dict[str, Any]]:
        selected = self.selected_requirement_baseline()
        if selected is None:
            return []
        return self.database.recognized_credit_declarations(selected["version"])

    def _validated_recognized_credit(
        self, payload: dict[str, Any], *, identity: str | None = None
    ) -> dict[str, Any]:
        selected = self.selected_requirement_baseline()
        if selected is None:
            raise ValueError("请先选择要求基线")
        declaration_identity = identity or str(payload.get("identity", ""))
        if not _LOCAL_IDENTITY.fullmatch(declaration_identity):
            raise ValueError("申报身份格式无效")
        category = str(payload.get("category", ""))
        if category not in _RECOGNIZED_CREDIT_CATEGORIES:
            raise ValueError("不支持的认定学分类别")
        credits = payload.get("credits")
        if isinstance(credits, bool) or not isinstance(credits, (int, float)):
            raise TypeError("学分必须是数字")
        credits = float(credits)
        if not math.isfinite(credits) or not 0 < credits <= 100:
            raise ValueError("学分必须是大于 0 且不超过 100 的有限数")
        note = str(payload.get("note", "")).strip()
        if not note or len(note) > 300:
            raise ValueError("说明不能为空且不能超过 300 字符")
        recognized_on = str(payload.get("recognized_on", ""))
        try:
            recognized_date = date.fromisoformat(recognized_on)
        except ValueError as error:
            raise ValueError("认定日期格式无效") from error
        if recognized_date < date(2000, 1, 1) or recognized_date > datetime.now(timezone.utc).date():
            raise ValueError("认定日期超出合理范围")
        linked_identity = str(payload.get("linked_course_identity", "")).strip()
        if linked_identity and (len(linked_identity) > 128 or any(ord(char) < 32 for char in linked_identity)):
            raise ValueError("关联课程身份格式无效")
        return {
            "identity": declaration_identity,
            "baseline_version": selected["version"],
            "category": category,
            "credits": credits,
            "note": note,
            "recognized_on": recognized_on,
            "linked_course_identity": linked_identity,
        }

    def create_recognized_credit(
        self, payload: dict[str, Any]
    ) -> tuple[dict[str, Any], bool]:
        validated = self._validated_recognized_credit(payload)
        declaration, created = self.database.create_recognized_credit(validated)
        if declaration["baseline_version"] != validated["baseline_version"]:
            raise ValueError("申报身份已用于其他要求基线")
        return declaration, created

    def update_recognized_credit(
        self, identity: str, payload: dict[str, Any]
    ) -> dict[str, Any]:
        return self.database.update_recognized_credit(
            identity, self._validated_recognized_credit(payload, identity=identity)
        )

    def delete_recognized_credit(self, identity: str) -> bool:
        if not _LOCAL_IDENTITY.fullmatch(identity):
            raise ValueError("申报身份格式无效")
        selected = self.selected_requirement_baseline()
        if selected is None:
            raise ValueError("请先选择要求基线")
        return self.database.delete_recognized_credit(identity, selected["version"])

    def _current_course_facts(self) -> dict[str, dict[str, Any]]:
        snapshot = self.database.latest_snapshot("progress") or {}
        report = (snapshot.get("payload") or {}).get("report") or {}
        selected = self.selected_requirement_baseline()
        if not selected or report.get("baseline_version") != selected["version"]:
            return {}
        courses = [
            course
            for item in report.get("progress", ())
            for course in item.get("courses", ())
            if isinstance(course, dict)
        ]
        courses.extend(
            course for course in report.get("unclassified_courses", ())
            if isinstance(course, dict)
        )
        profile_id = (self.database.current_profile() or {}).get("version_id")
        timetable = self.database.latest_snapshot("timetable") or {}
        if profile_id and timetable.get("profile_id") == profile_id:
            courses.extend(
                course for course in (timetable.get("payload") or {}).get("enrolled_courses", ())
                if isinstance(course, dict)
            )
        selection = self.database.latest_snapshot("selection") or {}
        if profile_id and selection.get("profile_id") == profile_id:
            courses.extend(
                course for course in (selection.get("payload") or {}).get("sections", ())
                if isinstance(course, dict)
            )
        facts: dict[str, dict[str, Any]] = {}
        for course in courses:
            code = str(course.get("code") or course.get("course_code") or "").strip()
            name = str(course.get("name") or course.get("course_name") or "").strip()
            identity = code or " ".join(name.lower().split())
            if identity:
                facts[identity] = {
                    **course,
                    "code": code,
                    "name": name,
                    "category": str(course.get("category", "")),
                    "credits": course.get("credits", course.get("credit", 0)),
                }
        return facts

    def course_labels(self) -> list[dict[str, Any]]:
        selected = self.selected_requirement_baseline()
        return [] if selected is None else self.database.course_labels(selected["version"])

    def save_course_label(self, course_identity: str, payload: dict[str, Any]) -> dict[str, Any]:
        selected = self.selected_requirement_baseline()
        if selected is None:
            raise ValueError("请先选择要求基线")
        if course_identity not in self._current_course_facts():
            raise ValueError("课程身份不在当前完整成绩报告中")
        d_category = payload.get("d_category", False)
        four_histories = payload.get("four_histories", False)
        if not isinstance(d_category, bool) or not isinstance(four_histories, bool):
            raise TypeError("课程标签必须是布尔值")
        outside_track = self._validated_track(payload.get("outside_track", ""), allow_empty=True)
        return self.database.save_course_label({
            "course_identity": course_identity,
            "baseline_version": selected["version"],
            "d_category": d_category,
            "four_histories": four_histories,
            "outside_track": outside_track,
        })

    def delete_course_label(self, course_identity: str) -> bool:
        selected = self.selected_requirement_baseline()
        if selected is None:
            raise ValueError("请先选择要求基线")
        return self.database.delete_course_label(course_identity, selected["version"])

    @staticmethod
    def _validated_track(value: Any, *, allow_empty: bool = False) -> str:
        if not isinstance(value, str):
            raise TypeError("外专业课程体系名称必须是字符串")
        track = value.strip()
        if not track and allow_empty:
            return ""
        if not track or len(track) > 80 or any(ord(char) < 32 for char in track):
            raise ValueError("外专业课程体系名称格式无效")
        return track

    def outside_major_track(self) -> dict[str, str] | None:
        selected = self.selected_requirement_baseline()
        track = self.database.outside_major_track_selection()
        if not selected or not track or track.get("baseline_version") != selected["version"]:
            return None
        return track

    def select_outside_major_track(self, value: Any) -> dict[str, str]:
        selected = self.selected_requirement_baseline()
        if selected is None:
            raise ValueError("请先选择要求基线")
        return self.database.select_outside_major_track(
            selected["version"], self._validated_track(value)
        )

    def clear_outside_major_track(self) -> None:
        self.database.clear_outside_major_track()

    def _with_local_estimates(
        self, classified: dict[str, Any], *, goals: list[dict[str, Any]] | None = None
    ) -> dict[str, Any]:
        report = classified.get("report")
        selected = self.selected_requirement_baseline()
        if not isinstance(report, dict) or selected is None:
            return classified
        result = deepcopy(classified)
        copied_report = result["report"]
        progress = apply_recognized_credit_estimates(
            copied_report.get("progress", []), selected, self.recognized_credits()
        )
        labels = self.course_labels()
        track = self.outside_major_track()
        selected_track = "" if track is None else track["track"]
        progress = apply_course_label_estimates(
            progress,
            copied_report.get("unclassified_courses", []),
            labels,
            selected_track,
        )
        timetable = self.database.latest_snapshot("timetable") or {}
        selection = self.database.latest_snapshot("selection") or {}
        profile_id = (self.database.current_profile() or {}).get("version_id")
        timetable_payload = timetable.get("payload") or {}
        if not profile_id or timetable.get("profile_id") != profile_id:
            timetable_payload = {}
        selection_payload = selection.get("payload") or {}
        if not profile_id or selection.get("profile_id") != profile_id:
            selection_payload = {}
        sections = selection_payload.get("sections", [])
        active_goals = goals
        if active_goals is None:
            latest_plan = self.database.latest_plan() or {}
            plan_snapshot_id = latest_plan.get("selection_snapshot_id")
            active_goals = (
                latest_plan.get("goals", [])
                if not plan_snapshot_id or plan_snapshot_id == selection.get("id")
                else []
            )
        section_by_identity = {
            str(section.get("identity") or section.get("section_id") or ""): section
            for section in sections if isinstance(section, dict)
        }
        queued_sections = [
            section_by_identity[section_id]
            for goal in active_goals if isinstance(goal, dict)
            for preferences in (goal.get("preferences"),)
            if isinstance(preferences, list)
            for preference in preferences[:1]
            if isinstance(preference, dict)
            and (section_id := str(preference.get("section_id", ""))) in section_by_identity
        ]
        copied_report["progress"] = apply_projected_course_estimates(
            progress,
            selected,
            enrolled_courses=timetable_payload.get("enrolled_courses", ()),
            queued_courses=queued_sections,
            labels=labels,
            selected_track=selected_track,
        )
        return result

    def _classified_progress(
        self, report: dict[str, Any], *, snapshot: dict[str, Any] | None = None
    ) -> dict[str, Any]:
        version = str(report.get("baseline_version", ""))
        if requirement_baseline(version) is None:
            return {"status": "not_applicable", "report": None, "snapshot": snapshot}
        selected = self.selected_requirement_baseline()
        if selected is None:
            return {
                "status": "historical", "report": None,
                "historical_report": report, "snapshot": snapshot,
                "reason": "请选择要求基线后重新同步毕业进度",
            }
        if version != selected["version"]:
            return {
                "status": "historical", "report": None,
                "historical_report": report, "snapshot": snapshot,
                "reason": "要求基线已切换，请重新同步毕业进度",
            }
        status = "ready" if report.get("data_complete") is True else "incomplete"
        return {"status": status, "report": report, "snapshot": snapshot}

    def graduation_progress(
        self, *, goals: list[dict[str, Any]] | None = None
    ) -> dict[str, Any]:
        """Read progress with unified local projections under the selected baseline."""
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
            return self._with_local_estimates(
                self._classified_progress(report, snapshot=snapshot), goals=goals
            )

        # Keep reports generated by the existing standalone collector visible
        # as history until a selected baseline is refreshed in the workbench.
        path = self.progress_report_path
        if path is None or not path.is_file():
            return {"status": "missing", "report": None}
        try:
            report = json.loads(path.read_text(encoding="utf-8"))
        except (OSError, UnicodeError, json.JSONDecodeError):
            return {"status": "invalid", "report": None}
        if not isinstance(report, dict) or not isinstance(report.get("progress"), list):
            return {"status": "invalid", "report": None}
        if not report.get("baseline_version"):
            report = {**report, "baseline_version": "guide-2026"}
        return self._with_local_estimates(
            self._classified_progress(report), goals=goals
        )

    def progress_projection(self, goals: list[dict[str, Any]]) -> dict[str, Any]:
        """Recalculate an unsaved local queue without academic-system access."""
        return self.graduation_progress(goals=goals)

    def progress_context(self) -> dict[str, Any]:
        selected = self.selected_requirement_baseline()
        if selected is None:
            raise ValueError("请先选择要求基线")
        profile = self.database.current_profile() or {}
        notice = self.database.confirmed_notice() or {}
        timetable = self.database.latest_snapshot("timetable") or {}
        return {
            "term": str(notice.get("term") or timetable.get("term") or ""),
            "profile_id": profile.get("version_id"),
            "baseline_version": selected["version"],
            "page_size": 20,
        }

    def state(self, *, session_state: str | dict[str, Any], active_task: dict[str, Any] | None = None, csrf_token: str | None = None) -> dict[str, Any]:
        selection = self.database.latest_snapshot("selection")
        timetable = self.database.latest_snapshot("timetable")
        progress_snapshot = self.database.latest_snapshot("progress")
        profile = self.database.current_profile() or {}
        notice = self.database.confirmed_notice() or {}
        selected_baseline = self.selected_requirement_baseline()

        def snapshot_status(kind: str, snapshot: dict[str, Any] | None) -> dict[str, Any]:
            if not snapshot:
                return {"status": "missing", "reason": "尚无本地快照", "source_at": ""}
            reasons: list[str] = []
            if snapshot.get("profile_id") and snapshot.get("profile_id") != profile.get("version_id"):
                reasons.append("学生画像已变化")
            if kind == "selection":
                if snapshot.get("notice_id") != notice.get("version_id"):
                    reasons.append("选课通知已变化")
                if (snapshot.get("payload") or {}).get("contract_version") != "hitwh-jwts-selection-query-v1":
                    reasons.append("查询契约已变化")
            if kind == "progress":
                report = (snapshot.get("payload") or {}).get("report") or {}
                if selected_baseline is None:
                    reasons.append("尚未选择要求基线")
                elif report.get("baseline_version") != selected_baseline["version"]:
                    reasons.append("要求基线已变化")
            return {
                "status": "historical" if reasons else "current",
                "reason": "、".join(reasons), "source_at": snapshot.get("source_at", ""),
            }

        result: dict[str, Any] = {
            "login_configuration": self.login_configuration(),
            "requirement_baselines": self.available_requirement_baselines(),
            "selected_requirement_baseline": self.selected_requirement_baseline(),
            "recognized_credits": self.recognized_credits(),
            "course_labels": self.course_labels(),
            "labelable_courses": list(self._current_course_facts().values()),
            "outside_major_track": self.outside_major_track(),
            "profile": self.effective_profile(),
            "confirmed_notice": self.database.confirmed_notice(),
            "snapshots": {"selection": selection, "timetable": timetable, "progress": progress_snapshot},
            "snapshot_changes": {
                "selection": self.database.latest_snapshot_change("selection"),
                "timetable": self.database.latest_snapshot_change("timetable"),
                "progress": self.database.latest_snapshot_change("progress"),
            },
            "latest_plan": self.database.latest_plan(),
            "graduation_progress": self.graduation_progress(),
            "snapshot_status": {
                "selection": snapshot_status("selection", selection),
                "timetable": snapshot_status("timetable", timetable),
                "progress": snapshot_status("progress", progress_snapshot),
            },
            "stale": {
                "selection": snapshot_status("selection", selection)["status"] != "current",
                "timetable": snapshot_status("timetable", timetable)["status"] != "current",
                "progress": snapshot_status("progress", progress_snapshot)["status"] != "current",
            },
            "academic_session": session_state if isinstance(session_state, dict) else {"state": session_state},
            "active_task": active_task,
            "execution_history": self.database.execution_history(),
        }
        if csrf_token is not None:
            result["csrf_token"] = csrf_token
        return result

    def effective_profile(self) -> dict[str, Any] | None:
        """Student facts with grade derived from the ID prefix when not saved."""
        profile = self.database.current_profile()
        if profile is not None:
            return dict(profile)
        masked = str((self.login_configuration() or {}).get("masked_username", ""))
        prefix = masked[:4]
        if prefix.isdigit() and len(prefix) == 4:
            return {"grade": prefix}
        return None

    def refresh_context(self) -> dict[str, Any]:
        profile = self.effective_profile()
        notice = self.database.confirmed_notice()
        timetable = self.database.latest_snapshot("timetable")
        windows = (notice or {}).get("windows", [])
        grade = str((profile or {}).get("grade", ""))
        notice_term = str((notice or {}).get("term", ""))
        timetable_term = str((timetable or {}).get("term", ""))
        terms_match = not timetable_term or (
            timetable_term.replace("年", "").replace("学期", "").replace(" ", "")
            == notice_term.replace("年", "").replace("学期", "").replace(" ", "")
        )
        allowed = [item for item in windows if terms_match and item.get("action") == "selection" and item.get("method") == "academic_system" and grade and grade in item.get("grades", [])]
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
            if self.notice_fetcher is None:
                raise NoticeReadError("no notice reader is configured for this service")
            try:
                text = self.notice_fetcher(source_url)
            except (OSError, ValueError) as error:
                raise NoticeReadError(f"unable to read notice: {error}") from error
        if not source_url or not text:
            raise ValueError("source_url and text are required")
        candidate = candidate_from_text(source_url, text, official_hosts=self.official_notice_hosts)
        previous = self.database.confirmed_notice()
        saved = self.database.save_notice(candidate)
        return saved, notice_diff(previous, saved) if previous else ""

    def discover_notice_candidates(
        self, index_url: str = DEFAULT_NOTICE_INDEX_URL
    ) -> list[dict[str, Any]]:
        if self.notice_discoverer is None:
            raise NoticeReadError("no notice discoverer is configured for this service")
        try:
            candidates = self.notice_discoverer(index_url)
        except (OSError, ValueError) as error:
            raise NoticeReadError(f"unable to discover official notices: {error}") from error
        return [self.database.save_notice(candidate) for candidate in candidates]

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

    def prepare_selection_execution(self, section_id: str, snapshot_id: str) -> dict[str, Any]:
        """Validate one concrete section against the current applicable snapshots."""
        profile = self.database.current_profile() or {}
        notice = self.database.confirmed_notice() or {}
        selection = self.database.latest_snapshot("selection")
        timetable = self.database.latest_snapshot("timetable")
        if not selection or selection.get("id") != snapshot_id:
            raise ValueError("选课快照已更新，请重新选择教学班")
        sections = (selection.get("payload") or {}).get("sections", [])
        matches = [item for item in sections if isinstance(item, dict) and str(item.get("identity")) == section_id]
        if len(matches) != 1:
            raise ValueError("当前选课快照中不存在该教学班")
        section = matches[0]
        if not section.get("execution_ready") or str(section.get("action_rwh")) != section_id:
            raise ValueError("该教学班缺少页面提供的可执行身份")
        term = str(notice.get("term") or selection.get("term") or "")
        plan = build_read_only_plan(
            term=term,
            profile_id=str(profile.get("version_id", "")),
            notice_id=str(notice.get("version_id", "")),
            timetable_snapshot=timetable,
            selection_snapshot=selection,
            goals=[{
                "goal_id": section_id,
                "course_identity": str(section.get("course_code") or section.get("name") or section_id),
                "rank": 1,
                "preferences": [{"section_id": section_id, "rank": 1}],
            }],
        )
        reasons = list(plan.blocked_reasons)
        reasons.extend(conflict.kind for conflict in plan.conflicts)
        if reasons:
            raise ValueError("当前教学班不可执行：" + "、".join(dict.fromkeys(reasons)))
        category = str(section.get("query_code") or "")
        query_term = str(section.get("query_term") or "")
        if not category or not query_term:
            raise ValueError("教学班缺少查询类别或学期来源")
        unresolved = self.database.unresolved_execution(section_id)
        if unresolved:
            raise ValueError("该教学班上次执行结果未知，请先核实并解除阻断")
        if category not in self.refresh_context().get("allowed_categories", []):
            raise ValueError("教学班类别不在当前通知白名单中")
        return {
            "section_id": section_id,
            "category": category,
            "term_value": query_term,
            "source_page": max(1, int(section.get("query_page") or 1)),
            "snapshot_id": snapshot_id,
            "profile_id": profile.get("version_id"),
            "notice_id": notice.get("version_id"),
            "course_name": str(section.get("name") or ""),
        }
