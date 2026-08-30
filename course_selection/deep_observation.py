"""Typed, replayable observation boundary for academic browser reads."""

from __future__ import annotations

import json
import os
import re
import time
from collections.abc import Iterable
from dataclasses import dataclass, field
from datetime import datetime, timezone
from pathlib import Path
from typing import Any, Protocol

from course_progress.sanitizer import sanitize_url

_FIELD_NAME = re.compile(r"(?:^|[?&;])([^=&;]+)=")


def _utc_now() -> str:
    return datetime.now(timezone.utc).isoformat(timespec="milliseconds")


def _field_names(post_data: str | None) -> tuple[str, ...]:
    if not post_data:
        return ()
    return tuple(sorted(set(_FIELD_NAME.findall(post_data))))


@dataclass(frozen=True)
class AcademicRequestTraceEvent:
    sequence: int
    occurred_at: str
    method: str
    url: str
    resource_type: str
    field_names: tuple[str, ...] = ()

    @classmethod
    def from_request(cls, sequence: int, request: dict[str, Any]) -> AcademicRequestTraceEvent:
        return cls(
            sequence=sequence,
            occurred_at=str(request.get("occurred_at") or _utc_now()),
            method=str(request.get("method") or "GET").upper(),
            url=sanitize_url(str(request.get("url") or "")),
            resource_type=str(request.get("resource_type") or "other"),
            field_names=_field_names(request.get("post_data")),
        )

    def to_dict(self) -> dict[str, Any]:
        return {
            "sequence": self.sequence,
            "occurred_at": self.occurred_at,
            "method": self.method,
            "url": self.url,
            "resource_type": self.resource_type,
            "field_names": list(self.field_names),
        }


@dataclass(frozen=True)
class AcademicRequestTrace:
    events: tuple[AcademicRequestTraceEvent, ...] = ()

    @classmethod
    def empty(cls) -> AcademicRequestTrace:
        return cls()

    @classmethod
    def from_requests(cls, requests: Iterable[dict[str, Any]]) -> AcademicRequestTrace:
        return cls(tuple(AcademicRequestTraceEvent.from_request(index, request) for index, request in enumerate(requests, start=1)))


@dataclass(frozen=True)
class TimetableObservationRequest:
    term: str
    context: dict[str, Any]


@dataclass(frozen=True)
class TimetableObservationResult:
    status: str
    term: str = ""
    entries: tuple[dict[str, Any], ...] = ()
    enrolled_courses: tuple[dict[str, Any], ...] = ()
    source_kind: str = "personal-timetable-api"
    trace: AcademicRequestTrace = field(default_factory=AcademicRequestTrace.empty)
    error: str = ""

    @classmethod
    def complete(
        cls, *, term: str, entries: list[dict[str, Any]],
        enrolled_courses: list[dict[str, Any]] | None = None,
        trace: AcademicRequestTrace,
    ) -> TimetableObservationResult:
        return cls(
            status="complete", term=term, entries=tuple(entries),
            enrolled_courses=tuple(enrolled_courses or ()), trace=trace,
        )

    @classmethod
    def incomplete(cls, error: str, *, trace: AcademicRequestTrace | None = None) -> TimetableObservationResult:
        return cls(status="incomplete", error=error, trace=trace or AcademicRequestTrace.empty())

    @classmethod
    def cancelled(cls, *, trace: AcademicRequestTrace | None = None) -> TimetableObservationResult:
        return cls(status="cancelled", trace=trace or AcademicRequestTrace.empty())

    def snapshot_payload(self) -> dict[str, Any]:
        return {
            "status": self.status,
            "term": self.term,
            "entries": list(self.entries),
            "enrolled_courses": list(self.enrolled_courses),
            "source_kind": self.source_kind,
        }


@dataclass(frozen=True)
class SelectionObservationRequest:
    context: dict[str, Any]


@dataclass(frozen=True)
class SelectionObservationResult:
    status: str
    payload: dict[str, Any] = field(default_factory=dict)
    trace: AcademicRequestTrace = field(default_factory=AcademicRequestTrace.empty)
    error: str = ""

    @classmethod
    def complete(cls, payload: dict[str, Any], *, trace: AcademicRequestTrace) -> SelectionObservationResult:
        return cls(status="complete", payload=payload, trace=trace)

    @classmethod
    def incomplete(cls, error: str, *, trace: AcademicRequestTrace | None = None) -> SelectionObservationResult:
        return cls(status="incomplete", trace=trace or AcademicRequestTrace.empty(), error=error)


@dataclass(frozen=True)
class SelectionDiscoveryDiagnostic:
    """Non-publishable evidence produced when the verified contract is absent."""

    status: str = "interface_unconfirmed"
    diagnostic: dict[str, Any] = field(default_factory=dict)
    trace: AcademicRequestTrace = field(default_factory=AcademicRequestTrace.empty)
    error: str = ""


@dataclass(frozen=True)
class ProgressObservationRequest:
    context: dict[str, Any]


@dataclass(frozen=True)
class ProgressObservationResult:
    status: str
    payload: dict[str, Any] = field(default_factory=dict)
    trace: AcademicRequestTrace = field(default_factory=AcademicRequestTrace.empty)
    error: str = ""

    @classmethod
    def complete(cls, payload: dict[str, Any], *, trace: AcademicRequestTrace) -> ProgressObservationResult:
        return cls(status="complete", payload=payload, trace=trace)

    @classmethod
    def incomplete(cls, error: str, *, trace: AcademicRequestTrace | None = None) -> ProgressObservationResult:
        return cls(status="incomplete", trace=trace or AcademicRequestTrace.empty(), error=error)

    @classmethod
    def cancelled(cls, *, trace: AcademicRequestTrace | None = None) -> ProgressObservationResult:
        return cls(status="cancelled", trace=trace or AcademicRequestTrace.empty())


@dataclass(frozen=True)
class ManualObservationRequest:
    context: dict[str, Any]


@dataclass(frozen=True)
class ManualObservationResult:
    """Diagnostic-only result; deliberately has no publishable payload."""

    status: str
    diagnostic: dict[str, Any] = field(default_factory=dict)
    trace: AcademicRequestTrace = field(default_factory=AcademicRequestTrace.empty)
    error: str = ""


class TimetableObserver(Protocol):
    def observe_timetable(self, request: TimetableObservationRequest, progress, cancelled) -> TimetableObservationResult: ...


class SelectionObserver(Protocol):
    def observe_selection(
        self, request: SelectionObservationRequest, progress, cancelled
    ) -> SelectionObservationResult | SelectionDiscoveryDiagnostic: ...


class TraceStore:
    """Filesystem storage for bounded, sanitized per-task JSON Lines traces."""

    def __init__(self, root: Path | str, *, retain: int = 20):
        self.root = Path(root)
        self.retain = retain
        self._last_stamp = 0

    def write(self, task_id: str, trace: AcademicRequestTrace) -> Path:
        self.root.mkdir(parents=True, exist_ok=True)
        path = self.root / f"{task_id}.jsonl"
        temporary = path.with_suffix(".jsonl.tmp")
        with temporary.open("w", encoding="utf-8", newline="\n") as handle:
            for event in trace.events:
                handle.write(json.dumps(event.to_dict(), ensure_ascii=False, separators=(",", ":")) + "\n")
        temporary.replace(path)
        # Filesystem timestamp resolution can be one second on Windows; assign
        # a strictly increasing completion timestamp so rapid tasks retain the
        # actual newest traces.
        # FAT/NTFS timestamps may round sub-millisecond differences, so keep
        # an explicit millisecond gap between rapid completed traces.
        stamp = max(time.time_ns(), self._last_stamp + 1_000_000)
        self._last_stamp = stamp
        os.utime(path, ns=(stamp, stamp))
        traces = sorted(self.root.glob("*.jsonl"), key=lambda item: item.stat().st_mtime_ns, reverse=True)
        for old in traces[self.retain :]:
            old.unlink()
        return path


class ReplayAcademicObserver:
    """Deterministic adapter used to validate publication without Playwright."""

    def __init__(self, result: TimetableObservationResult | SelectionObservationResult | SelectionDiscoveryDiagnostic | ProgressObservationResult | ManualObservationResult):
        self.result = result
        self.requests: list[TimetableObservationRequest] = []
        self.closed = False

    def connect(self, progress, cancelled) -> None:
        progress("connecting", {})

    def observe_timetable(self, request: TimetableObservationRequest, progress, cancelled) -> TimetableObservationResult:
        self.requests.append(request)
        return self.result

    def observe_selection(self, request: SelectionObservationRequest, progress, cancelled):
        self.requests.append(request)
        if isinstance(self.result, (SelectionObservationResult, SelectionDiscoveryDiagnostic)):
            return self.result
        return SelectionObservationResult.incomplete("replay has no selection result")

    def observe_progress(self, request: ProgressObservationRequest, progress, cancelled) -> ProgressObservationResult:
        self.requests.append(request)
        if isinstance(self.result, ProgressObservationResult):
            return self.result
        return ProgressObservationResult.incomplete("replay has no progress result")

    def observe_manual(self, request: ManualObservationRequest, progress, cancelled, finished) -> ManualObservationResult:
        self.requests.append(request)
        if isinstance(self.result, ManualObservationResult):
            return self.result
        return ManualObservationResult(status="incomplete", error="replay has no manual result")

    def close(self) -> None:
        self.closed = True
