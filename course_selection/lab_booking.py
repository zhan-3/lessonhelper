"""Plan and submit openlab lab bookings for one teaching-center app.

The module is split so the decision logic is testable without a browser:

* :func:`busy_intervals` turns a workbench timetable snapshot into occupied
  weekday/period/week ranges.
* :func:`plan_lab_slots` picks at most one free seat per experiment, skipping
  anything that collides with that timetable, and reports same-time collisions
  between the chosen slots as reminders for the user instead of hiding them.
* :func:`book_slots` submits each explicit target at most once, re-reads
  availability immediately before submitting, and stops on the first outcome
  that needs human verification.

Only the booking submission is a write.  Every other call is a read, and the
session token is never persisted.
"""

from __future__ import annotations

import hashlib
import json
from collections.abc import Iterable, Mapping, Sequence
from dataclasses import dataclass, field
from datetime import date
from typing import Any, Protocol

# Lab "大节" numbers map onto two teaching periods each.
LAB_PERIODS: Mapping[int, tuple[int, int]] = {
    1: (1, 2),
    2: (3, 4),
    3: (5, 6),
    4: (7, 8),
    5: (9, 10),
    6: (11, 12),
}

# Published start time of each lab session, used to map a booked lab back onto periods.
LAB_START_TIMERS: Mapping[str, int] = {
    "07:45": 1,
    "10:05": 2,
    "13:45": 3,
    "16:05": 4,
    "18:40": 5,
    "20:45": 6,
}


@dataclass(frozen=True)
class BusyInterval:
    """One occupied span from the personal timetable snapshot."""

    weekday: int
    start_period: int
    end_period: int
    weeks: frozenset[int]
    label: str = ""

    def overlaps(self, weekday: int, week: int, periods: tuple[int, int]) -> bool:
        if weekday != self.weekday or week not in self.weeks:
            return False
        return self.start_period <= periods[1] and periods[0] <= self.end_period


def week_numbers(entry: Mapping[str, Any]) -> frozenset[int]:
    """Normalize the several week encodings a timetable entry may carry."""
    explicit = entry.get("week_numbers")
    if isinstance(explicit, Sequence) and not isinstance(explicit, (str, bytes)):
        return frozenset(int(value) for value in explicit)
    start = int(entry.get("week_start") or 1)
    end = int(entry.get("week_end") or start)
    parity = str(entry.get("week_parity") or "all")
    weeks = range(start, end + 1)
    if parity == "odd":
        return frozenset(week for week in weeks if week % 2 == 1)
    if parity == "even":
        return frozenset(week for week in weeks if week % 2 == 0)
    return frozenset(weeks)


def busy_intervals(entries: Iterable[Mapping[str, Any]]) -> tuple[BusyInterval, ...]:
    """Build occupied intervals; rows with unknown times stay excluded."""
    intervals = []
    for entry in entries:
        if str(entry.get("conflict_status") or "") == "unknown":
            continue
        weekday = int(entry.get("weekday") or 0)
        start = int(entry.get("start_period") or 0)
        end = int(entry.get("end_period") or start)
        if not 1 <= weekday <= 7 or start <= 0:
            continue
        intervals.append(
            BusyInterval(
                weekday=weekday,
                start_period=start,
                end_period=max(start, end),
                weeks=week_numbers(entry),
                label=str(entry.get("course_name") or ""),
            )
        )
    return tuple(intervals)


def busy_from_bookings(rows: Iterable[Mapping[str, Any]]) -> tuple[BusyInterval, ...]:
    """Turn already-booked lab sessions into occupied intervals.

    Lab sessions live in a different system from the personal timetable, so a
    planner that only reads the timetable will happily schedule two labs at the
    same moment.  Passing these intervals in keeps that from happening.
    """
    intervals = []
    for row in rows:
        try:
            class_date = date.fromisoformat(str(row["classDate"]))
            week = int(row["eduWeek"])
        except (KeyError, ValueError):
            continue
        timer = LAB_START_TIMERS.get(str(row.get("startTime", ""))[:5])
        periods = LAB_PERIODS.get(timer) if timer else None
        if periods is None:
            continue
        intervals.append(
            BusyInterval(
                weekday=class_date.isoweekday(),
                start_period=periods[0],
                end_period=periods[1],
                weeks=frozenset({week}),
                label=str(row.get("subjectName") or "已约实验"),
            )
        )
    return tuple(intervals)


@dataclass(frozen=True)
class LabSlot:
    """One bookable seat in one experiment session."""

    subject_id: int
    subject_name: str
    class_date: str
    week: int
    weekday: int
    timer: int
    timer_name: str
    start: str
    lab: str
    table_no: str
    seat_id: str

    @property
    def key(self) -> str:
        return f"{self.subject_id}|{self.class_date}|{self.timer}"

    def to_dict(self) -> dict[str, Any]:
        return {
            "subject_id": self.subject_id,
            "subject_name": self.subject_name,
            "class_date": self.class_date,
            "week": self.week,
            "weekday": self.weekday,
            "timer": self.timer,
            "timer_name": self.timer_name,
            "start": self.start,
            "lab": self.lab,
            "table_no": self.table_no,
            "seat_id": self.seat_id,
        }

    @classmethod
    def from_dict(cls, payload: Mapping[str, Any]) -> LabSlot:
        return cls(
            subject_id=int(payload["subject_id"]),
            subject_name=str(payload["subject_name"]),
            class_date=str(payload["class_date"]),
            week=int(payload["week"]),
            weekday=int(payload["weekday"]),
            timer=int(payload["timer"]),
            timer_name=str(payload["timer_name"]),
            start=str(payload["start"]),
            lab=str(payload["lab"]),
            table_no=str(payload["table_no"]),
            seat_id=str(payload["seat_id"]),
        )


class LabSession(Protocol):
    """Read/write surface of one center app; implemented by a browser adapter or a fake."""

    def booked(self) -> list[Mapping[str, Any]]: ...

    def subjects(self) -> list[Mapping[str, Any]]: ...

    def schedule(self, subject_id: int) -> list[Mapping[str, Any]]: ...

    def free_seat(self, subject_id: int, class_date: str, timer: int) -> Mapping[str, Any] | None: ...

    def submit(self, slot: LabSlot) -> Mapping[str, Any]: ...


@dataclass
class PlanResult:
    slots: list[LabSlot] = field(default_factory=list)
    reminders: list[tuple[LabSlot, LabSlot]] = field(default_factory=list)
    unresolved: list[str] = field(default_factory=list)


def _collides(interval: BusyInterval, class_date_weekday: int, week: int, timer: int) -> bool:
    periods = LAB_PERIODS.get(timer)
    if periods is None:
        return True
    return interval.overlaps(class_date_weekday, week, periods)


class OnlySubjects:
    """Restrict a session to a subset of experiments, for targeted replanning."""

    def __init__(self, session: LabSession, subject_ids: Iterable[int]):
        self._session = session
        self._wanted = {int(value) for value in subject_ids}

    def booked(self) -> list[Mapping[str, Any]]:
        return self._session.booked()

    def subjects(self) -> list[Mapping[str, Any]]:
        return [row for row in self._session.subjects() if int(row["subjectId"]) in self._wanted]

    def schedule(self, subject_id: int) -> list[Mapping[str, Any]]:
        return self._session.schedule(subject_id)

    def free_seat(self, subject_id: int, class_date: str, timer: int) -> Mapping[str, Any] | None:
        return self._session.free_seat(subject_id, class_date, timer)

    def submit(self, slot: LabSlot) -> Mapping[str, Any]:
        return self._session.submit(slot)


def plan_lab_slots(
    session: LabSession,
    busy: Iterable[BusyInterval],
    *,
    probe_cap: int = 8,
) -> PlanResult:
    """Choose the earliest free seat per experiment, avoiding timetable clashes."""
    intervals = tuple(busy)
    done = {int(row["subjectId"]) for row in session.booked()}
    result = PlanResult()
    for subject in session.subjects():
        subject_id = int(subject["subjectId"])
        name = str(subject.get("subjectName") or subject_id)
        if subject_id in done:
            continue
        candidates = [
            (row, timer)
            for row in session.schedule(subject_id)
            for timer in row.get("timerList", [])
            if not any(
                _collides(interval, _weekday_of(row), int(row["eduWeek"]), int(timer["timer"]))
                for interval in intervals
            )
        ]
        chosen = None
        for row, timer in candidates[:probe_cap]:
            seat = session.free_seat(subject_id, str(row["classDate"]), int(timer["timer"]))
            if seat is None:
                continue
            chosen = LabSlot(
                subject_id=subject_id,
                subject_name=name,
                class_date=str(row["classDate"]),
                week=int(row["eduWeek"]),
                weekday=_weekday_of(row),
                timer=int(timer["timer"]),
                timer_name=str(timer["timerName"]),
                start=str(timer["startTime"])[:5],
                lab=str(seat.get("lab") or ""),
                table_no=str(seat.get("tableNo") or ""),
                seat_id=str(seat.get("seatId") or ""),
            )
            break
        if chosen is None:
            result.unresolved.append(name)
        else:
            result.slots.append(chosen)
    result.reminders = same_time_pairs(result.slots)
    return result


_WEEKDAYS = {"星期一": 1, "星期二": 2, "星期三": 3, "星期四": 4, "星期五": 5, "星期六": 6, "星期日": 7, "星期天": 7}


def _weekday_of(row: Mapping[str, Any]) -> int:
    value = row.get("dayWeek")
    if isinstance(value, int):
        return value
    return _WEEKDAYS.get(str(value), 0)


def same_time_pairs(slots: Iterable[LabSlot]) -> list[tuple[LabSlot, LabSlot]]:
    """Report chosen slots that share a date and a session, for manual resolution."""
    ordered = sorted(slots, key=lambda slot: (slot.class_date, slot.timer, slot.subject_id))
    pairs = []
    for index, left in enumerate(ordered):
        for right in ordered[index + 1 :]:
            if left.class_date == right.class_date and left.timer == right.timer:
                pairs.append((left, right))
    return pairs


def plan_token(slots: Iterable[LabSlot]) -> str:
    """Deterministic confirmation token binding a run to its exact targets."""
    digest = hashlib.sha256()
    for slot in sorted(slots, key=lambda item: item.key):
        digest.update(f"{slot.key}|{slot.seat_id}\n".encode())
    return digest.hexdigest()[:16]


def book_slots(
    session: LabSession, slots: Iterable[LabSlot], *, confirmation: str
) -> list[dict[str, Any]]:
    """Submit each target at most once; stop at the first unverifiable outcome."""
    targets = list(slots)
    expected = plan_token(targets)
    if confirmation != expected:
        raise ValueError(f"confirmation token mismatch: expected {expected}")
    outcomes: list[dict[str, Any]] = []
    for slot in targets:
        seat = session.free_seat(slot.subject_id, slot.class_date, slot.timer)
        if seat is None:
            outcomes.append({**slot.to_dict(), "outcome": "skipped_unavailable",
                             "detail": "no free seat at submit time"})
            break
        payload = session.submit(slot)
        if payload.get("transport") != "ok":
            outcomes.append({**slot.to_dict(), "outcome": "possibly_applied",
                             "detail": str(payload.get("detail") or payload.get("transport"))})
            break
        code, message = payload.get("code"), str(payload.get("message") or "")
        if code == 0 and payload.get("result") is True:
            outcomes.append({**slot.to_dict(), "outcome": "confirmed_success", "detail": message})
            continue
        outcomes.append({**slot.to_dict(), "outcome": "confirmed_failure", "detail": f"{code} {message}"})
        break
    return outcomes


def load_workspace_timetable(private_root: Any) -> tuple[Mapping[str, Any], ...]:
    """Read the latest timetable snapshot rows from the local workspace database."""
    from .persistence import WorkspaceDatabase

    database = WorkspaceDatabase.open(private_root)
    try:
        snapshot = database.latest_snapshot("timetable") or {}
        payload = snapshot.get("payload") or {}
        entries = payload.get("entries") or ()
        return tuple(entry for entry in entries if isinstance(entry, Mapping))
    finally:
        database.close()


def plan_to_json(result: PlanResult) -> str:
    return json.dumps(
        {
            "slots": [slot.to_dict() for slot in result.slots],
            "reminders": [[left.key, right.key] for left, right in result.reminders],
            "unresolved": result.unresolved,
            "confirmation": plan_token(result.slots),
        },
        ensure_ascii=False,
        indent=2,
    )
