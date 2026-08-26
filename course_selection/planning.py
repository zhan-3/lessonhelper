"""Framework-independent, read-only selection planning core.

The planner consumes already-published academic snapshots.  It never talks to
the academic gateway and deliberately treats missing schedule information as
``conflict_unknown`` rather than as free time.
"""

from __future__ import annotations

from dataclasses import dataclass, field
from datetime import datetime, timezone
from typing import Any, Iterable, Mapping


class PlanningBlockedError(ValueError):
    """Raised when the inputs cannot safely drive a planning decision."""

    def __init__(self, reasons: Iterable[str]):
        self.reasons = tuple(dict.fromkeys(str(reason) for reason in reasons))
        super().__init__("planning blocked: " + ", ".join(self.reasons))


@dataclass(frozen=True)
class SectionPreference:
    section_id: str
    rank: int


@dataclass(frozen=True)
class CourseGoal:
    goal_id: str
    course_identity: str
    rank: int
    preferences: tuple[SectionPreference, ...] = ()


@dataclass(frozen=True)
class Conflict:
    section_id: str
    kind: str  # current_timetable, candidate_sections, conflict_unknown
    with_id: str | None = None
    detail: str = ""


@dataclass(frozen=True)
class ReadOnlyPlan:
    term: str
    profile_id: str
    notice_id: str
    timetable_snapshot_id: str
    selection_snapshot_id: str
    goals: tuple[CourseGoal, ...]
    conflicts: tuple[Conflict, ...] = ()
    blocked_reasons: tuple[str, ...] = ()

    @property
    def ready(self) -> bool:
        return not self.blocked_reasons

    def to_dict(self) -> dict[str, Any]:
        return {
            "status": "ready" if self.ready else "blocked",
            "term": self.term,
            "profile_id": self.profile_id,
            "notice_id": self.notice_id,
            "timetable_snapshot_id": self.timetable_snapshot_id,
            "selection_snapshot_id": self.selection_snapshot_id,
            "goals": [
                {"goal_id": g.goal_id, "course_identity": g.course_identity, "rank": g.rank,
                 "preferences": [{"section_id": p.section_id, "rank": p.rank} for p in g.preferences]}
                for g in self.goals
            ],
            "conflicts": [{"section_id": c.section_id, "kind": c.kind, "with_id": c.with_id, "detail": c.detail} for c in self.conflicts],
            "blocked_reasons": list(self.blocked_reasons),
        }


def _now(value: datetime | None) -> datetime:
    return value or datetime.now(timezone.utc)


def _parse_time(value: Any) -> tuple[int, int, int, tuple[int, ...], str] | None:
    """Read normalized fields, or the common ``meetings`` representation."""
    if not isinstance(value, Mapping):
        return None
    weekday = value.get("weekday", value.get("day"))
    start = value.get("start_period", value.get("start"))
    end = value.get("end_period", value.get("end", start))
    weeks = value.get("week_numbers", value.get("weeks", ()))
    try:
        if weekday is None or start is None or end is None:
            return None
        if isinstance(weeks, str):
            weeks = tuple(int(x) for x in weeks.replace(",", " ").split() if x.isdigit())
        return int(weekday), int(start), int(end), tuple(int(x) for x in weeks), str(value.get("week_parity", "all"))
    except (TypeError, ValueError):
        return None


def _meetings(item: Mapping[str, Any]) -> list[tuple[int, int, int, tuple[int, ...], str]] | None:
    raw = item.get("meetings", item.get("schedule_entries"))
    if isinstance(raw, list):
        parsed = [_parse_time(entry) for entry in raw]
        return None if any(entry is None for entry in parsed) else [entry for entry in parsed if entry]
    parsed = _parse_time(item)
    if parsed:
        return [parsed]
    # A non-empty textual time is not enough to prove absence of a conflict.
    return None


def _overlap(left: tuple[int, int, int, tuple[int, ...], str], right: tuple[int, int, int, tuple[int, ...], str]) -> bool:
    day_a, start_a, end_a, weeks_a, parity_a = left
    day_b, start_b, end_b, weeks_b, parity_b = right
    if day_a != day_b or max(start_a, start_b) > min(end_a, end_b):
        return False
    # End periods are represented by a second tuple item when supplied; the
    # normalized planner accepts the conservative start-period comparison.
    if weeks_a and weeks_b and not set(weeks_a).intersection(weeks_b):
        return False
    if parity_a != "all" and parity_b != "all" and parity_a != parity_b:
        return False
    return True


def _snapshot_reasons(snapshot: Mapping[str, Any] | None, *, kind: str, term: str, profile_id: str,
                      notice_id: str | None, now: datetime, max_age: int) -> list[str]:
    if not snapshot:
        return [f"{kind}_snapshot_missing"]
    reasons: list[str] = []
    if snapshot.get("kind") != kind or snapshot.get("complete", 1) not in (1, True, None):
        reasons.append(f"{kind}_snapshot_incomplete")
    if snapshot.get("term") != term:
        reasons.append(f"{kind}_term_mismatch")
    if snapshot.get("profile_id") != profile_id:
        reasons.append(f"{kind}_profile_mismatch")
    if kind == "selection" and snapshot.get("notice_id") != notice_id:
        reasons.append("selection_notice_mismatch")
    try:
        source_at = datetime.fromisoformat(str(snapshot.get("source_at")))
        if source_at.tzinfo is None:
            source_at = source_at.replace(tzinfo=timezone.utc)
        if (now - source_at).total_seconds() > max_age:
            reasons.append(f"{kind}_snapshot_stale")
    except (TypeError, ValueError):
        reasons.append(f"{kind}_source_time_invalid")
    payload = snapshot.get("payload") or {}
    if payload.get("status") in {"incomplete", "interface_unconfirmed", "failed"}:
        reasons.append(f"{kind}_snapshot_unusable")
    return reasons


def build_read_only_plan(*, term: str, profile_id: str, notice_id: str, timetable_snapshot: Mapping[str, Any] | None,
                         selection_snapshot: Mapping[str, Any] | None, goals: Iterable[Mapping[str, Any] | CourseGoal],
                         now: datetime | None = None, timetable_max_age: int = 86400,
                         selection_max_age: int = 1800) -> ReadOnlyPlan:
    """Build a local plan and conflict report from applicable snapshots."""
    current = _now(now)
    blocked = _snapshot_reasons(timetable_snapshot, kind="timetable", term=term, profile_id=profile_id,
                                notice_id=None, now=current, max_age=timetable_max_age)
    blocked += _snapshot_reasons(selection_snapshot, kind="selection", term=term, profile_id=profile_id,
                                 notice_id=notice_id, now=current, max_age=selection_max_age)
    parsed_goals: list[CourseGoal] = []
    for raw in goals:
        if isinstance(raw, CourseGoal):
            parsed_goals.append(raw)
            continue
        preferences = tuple(SectionPreference(str(p.get("section_id", p.get("identity", ""))), int(p.get("rank", i + 1)))
                            for i, p in enumerate(raw.get("preferences", raw.get("sections", []))) if p.get("section_id", p.get("identity")))
        parsed_goals.append(CourseGoal(str(raw.get("goal_id", raw.get("course_identity", ""))),
                                       str(raw.get("course_identity", raw.get("course_code", ""))), int(raw.get("rank", len(parsed_goals) + 1)), preferences))
    parsed_goals.sort(key=lambda goal: (goal.rank, goal.goal_id))
    if not parsed_goals:
        blocked.append("goals_missing")
    timetable_entries = ((timetable_snapshot or {}).get("payload") or {}).get("entries", [])
    sections = ((selection_snapshot or {}).get("payload") or {}).get("sections", [])
    current_times = [(str(entry.get("course_code") or entry.get("course_name") or "current"), _meetings(entry)) for entry in timetable_entries if isinstance(entry, Mapping)]
    section_map = {str(item.get("identity", item.get("section_id", ""))): item for item in sections if isinstance(item, Mapping)}
    conflicts: list[Conflict] = []
    selected_ids = [p.section_id for goal in parsed_goals for p in goal.preferences]
    for section_id in selected_ids:
        section = section_map.get(section_id)
        if not section:
            blocked.append(f"section_missing:{section_id}")
            continue
        section_times = _meetings(section)
        if not section_times:
            conflicts.append(Conflict(section_id, "conflict_unknown", detail="schedule missing or unparseable"))
            continue
        for other_id, other_times in current_times:
            if other_times and any(_overlap(left, right) for left in section_times for right in other_times):
                conflicts.append(Conflict(section_id, "current_timetable", other_id))
            elif other_times is None:
                conflicts.append(Conflict(section_id, "conflict_unknown", other_id))
    for index, left_id in enumerate(selected_ids):
        left_times = _meetings(section_map[left_id]) if left_id in section_map else None
        if not left_times:
            continue
        for right_id in selected_ids[index + 1:]:
            right_times = _meetings(section_map[right_id]) if right_id in section_map else None
            if right_times and any(_overlap(a, b) for a in left_times for b in right_times):
                conflicts.append(Conflict(left_id, "candidate_sections", right_id))
    return ReadOnlyPlan(term, profile_id, notice_id, str((timetable_snapshot or {}).get("id", "")),
                        str((selection_snapshot or {}).get("id", "")), tuple(parsed_goals), tuple(conflicts), tuple(dict.fromkeys(blocked)))
