"""Serialized, recoverable observation-task orchestration."""

from __future__ import annotations

import hashlib
import json
import queue
import threading
import uuid
from dataclasses import dataclass
from enum import Enum
from typing import Any, Callable

from .gateway import AcademicGateway
from .persistence import WorkspaceDatabase, utc_now


class TaskState(str, Enum):
    QUEUED = "queued"
    CONNECTING = "connecting"
    WAITING_FOR_AUTHENTICATION = "waiting_for_authentication"
    READING = "reading"
    INTERFACE_UNCONFIRMED = "interface_unconfirmed"
    SUCCEEDED = "succeeded"
    FAILED = "failed"
    CANCEL_REQUESTED = "cancel_requested"
    CANCELLED = "cancelled"


@dataclass(frozen=True)
class SubmittedTask:
    id: str
    state: str


class ObservationService:
    def __init__(self, database: WorkspaceDatabase, gateway_factory: Callable[[], AcademicGateway], *, autostart: bool = True):
        self.database = database
        self.gateway_factory = gateway_factory
        self.gateway: AcademicGateway | None = None
        self.session_state = "disconnected"
        self._queue: queue.Queue[str | None] = queue.Queue()
        self._cancelled: set[str] = set()
        self._done: dict[str, threading.Event] = {}
        self._lock = threading.RLock()
        self._thread: threading.Thread | None = None
        if autostart:
            self.start()

    def start(self) -> None:
        if self._thread and self._thread.is_alive():
            return
        self._thread = threading.Thread(target=self._run, name="academic-observation-worker", daemon=True)
        self._thread.start()

    def submit(self, operation: str, context: dict[str, Any] | None = None) -> SubmittedTask:
        if operation not in {"connect", "refresh-selection", "refresh-timetable"}:
            raise ValueError("unsupported observation operation")
        context = context or {}
        key = hashlib.sha256(json.dumps([operation, context], sort_keys=True, ensure_ascii=False).encode()).hexdigest()
        with self._lock, self.database.connection:
            row = self.database.connection.execute(
                "select id,state from observation_tasks where coalesce_key=? and state in ('queued','connecting','waiting_for_authentication','reading','cancel_requested') order by created_at desc limit 1",
                (key,),
            ).fetchone()
            if row:
                return SubmittedTask(row["id"], row["state"])
            identity, now = uuid.uuid4().hex, utc_now()
            self.database.connection.execute(
                "insert into observation_tasks values(?,?,?,?,?,?,?,?,?)",
                (identity, operation, key, TaskState.QUEUED.value, now, now, "{}", json.dumps(context, ensure_ascii=False), ""),
            )
            self._done[identity] = threading.Event()
            self._queue.put(identity)
            return SubmittedTask(identity, TaskState.QUEUED.value)

    def inspect(self, identity: str) -> dict[str, Any] | None:
        with self._lock:
            row = self.database.connection.execute("select * from observation_tasks where id=?", (identity,)).fetchone()
        if not row:
            return None
        return {"id": row["id"], "operation": row["operation"], "state": row["state"], "created_at": row["created_at"], "updated_at": row["updated_at"], "progress": json.loads(row["progress"]), "error": row["error"]}

    def cancel(self, identity: str) -> bool:
        with self._lock, self.database.connection:
            row = self.database.connection.execute("select state from observation_tasks where id=?", (identity,)).fetchone()
            if not row or row[0] in {TaskState.SUCCEEDED.value, TaskState.FAILED.value, TaskState.CANCELLED.value}:
                return False
            self._cancelled.add(identity)
            state = TaskState.CANCELLED.value if row[0] == TaskState.QUEUED.value else TaskState.CANCEL_REQUESTED.value
            self.database.connection.execute("update observation_tasks set state=?,updated_at=? where id=?", (state, utc_now(), identity))
            if state == TaskState.CANCELLED.value:
                self._done.setdefault(identity, threading.Event()).set()
            return True

    def wait(self, identity: str, timeout: float | None = None) -> bool:
        return self._done.setdefault(identity, threading.Event()).wait(timeout)

    def _update(self, identity: str, state: str, progress: dict[str, Any] | None = None, error: str = "") -> None:
        safe_error = error[:1000]
        with self._lock, self.database.connection:
            self.database.connection.execute("update observation_tasks set state=?,updated_at=?,progress=?,error=? where id=?", (state, utc_now(), json.dumps(progress or {}, ensure_ascii=False), safe_error, identity))

    def _run(self) -> None:
        while True:
            identity = self._queue.get()
            if identity is None:
                return
            task = self.inspect(identity)
            if not task or task["state"] == TaskState.CANCELLED.value:
                continue
            try:
                self._execute(identity)
            except Exception as error:
                self._update(identity, TaskState.FAILED.value, error=str(error))
            finally:
                self._done.setdefault(identity, threading.Event()).set()

    def _execute(self, identity: str) -> None:
        with self._lock:
            row = self.database.connection.execute("select operation,context from observation_tasks where id=?", (identity,)).fetchone()
        operation, context = row[0], json.loads(row[1])
        if identity in self._cancelled:
            self._update(identity, TaskState.CANCELLED.value)
            return
        if self.gateway is None:
            self.gateway = self.gateway_factory()
        def cancelled() -> bool:
            return identity in self._cancelled
        def progress(state: str, details: dict[str, Any]) -> None:
            mapped = TaskState(state).value
            if mapped == TaskState.WAITING_FOR_AUTHENTICATION.value:
                self.session_state = TaskState.WAITING_FOR_AUTHENTICATION.value
            elif mapped == TaskState.CONNECTING.value:
                self.session_state = TaskState.CONNECTING.value
            self._update(identity, mapped, details)
        self._update(identity, TaskState.CONNECTING.value)
        self.session_state = "connecting"
        self.gateway.connect(progress, cancelled)
        if cancelled():
            self._update(identity, TaskState.CANCELLED.value)
            return
        self.session_state = "connected"
        if operation == "connect":
            self._update(identity, TaskState.SUCCEEDED.value)
            return
        self._update(identity, TaskState.READING.value)
        kind = "selection" if operation == "refresh-selection" else "timetable"
        reader = self.gateway.refresh_selection if kind == "selection" else self.gateway.refresh_timetable
        result = reader(context, progress, cancelled)
        if cancelled():
            self._update(identity, TaskState.CANCELLED.value)
        elif result.get("status") == "complete":
            self.database.publish_snapshot(kind, context.get("term", ""), result, source="academic", profile_id=context.get("profile_id"), notice_id=context.get("notice_id"))
            self._update(identity, TaskState.SUCCEEDED.value)
        else:
            error = result.get("status", "incomplete")
            self.database.record_failed_attempt(kind, error)
            self._update(identity, TaskState.FAILED.value, error=error)

    def close(self) -> None:
        self._queue.put(None)
        if self._thread:
            self._thread.join(timeout=5)
        if self.gateway:
            self.gateway.close()
        self.session_state = "disconnected"
