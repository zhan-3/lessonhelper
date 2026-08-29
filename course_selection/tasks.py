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

from .deep_observation import (
    ManualObservationRequest,
    ProgressObservationRequest,
    ProgressObservationResult,
    SelectionDiscoveryDiagnostic,
    SelectionObservationRequest,
    SelectionObservationResult,
    TimetableObservationRequest,
    TimetableObservationResult,
    TraceStore,
)
from .gateway import AcademicGateway
from .persistence import WorkspaceDatabase, sanitize_for_storage, utc_now


class TaskState(str, Enum):
    QUEUED = "queued"
    CONNECTING = "connecting"
    WAITING_FOR_AUTHENTICATION = "waiting_for_authentication"
    READING = "reading"
    OBSERVING = "observing"
    INTERFACE_UNCONFIRMED = "interface_unconfirmed"
    SUCCEEDED = "succeeded"
    FAILED = "failed"
    CANCEL_REQUESTED = "cancel_requested"
    CANCELLED = "cancelled"


@dataclass(frozen=True)
class SubmittedTask:
    id: str
    state: str


TASK_TIMEOUT_SECONDS = {
    "launch-shell": 30,
    "reset-login": 30,
    "connect": 600,
    "refresh-selection": 180,
    "refresh-timetable": 90,
    "refresh-progress": 90,
    "observe-navigation": 1800,
    "execute-selection": 30,
}


class ObservationService:
    def __init__(self, database: WorkspaceDatabase, gateway_factory: Callable[[], AcademicGateway], *, autostart: bool = True, trace_store: TraceStore | None = None):
        self.database = database
        self.gateway_factory = gateway_factory
        self.trace_store = trace_store or TraceStore(database.root / "request-traces")
        self.gateway: AcademicGateway | None = None
        self.session_state = "disconnected"
        self.browser_state = "not_started"
        self.webvpn_state = "unknown"
        self.last_verified_at = ""
        self._queue: queue.Queue[str | None] = queue.Queue()
        self._cancelled: set[str] = set()
        self._expired: set[str] = set()
        self._finished: set[str] = set()
        self._shutdown = threading.Event()
        self._login_reset_error = ""
        self._done: dict[str, threading.Event] = {}
        self._lock = threading.RLock()
        self._thread: threading.Thread | None = None
        with self.database.connection:
            now = utc_now()
            message = "workbench restarted before this task finished"
            self.database.connection.execute(
                "update observation_tasks set state='failed',updated_at=?,error=? "
                "where state in ('queued','connecting','waiting_for_authentication','reading','observing','cancel_requested')",
                (now, message),
            )
            self.database.connection.execute(
                "update execution_tasks set state='failed',updated_at=?,error=? "
                "where state in ('queued','connecting','waiting_for_authentication','reading','cancel_requested')",
                (now, message),
            )
        if autostart:
            self.start()

    def start(self) -> None:
        if self._thread and self._thread.is_alive():
            return
        self._thread = threading.Thread(target=self._run, name="academic-observation-worker", daemon=True)
        self._thread.start()

    def submit(self, operation: str, context: dict[str, Any] | None = None) -> SubmittedTask:
        if operation not in {"launch-shell", "reset-login", "connect", "refresh-selection", "refresh-timetable", "refresh-progress", "observe-navigation"}:
            raise ValueError("unsupported observation operation")
        context = sanitize_for_storage(context or {})
        key = hashlib.sha256(json.dumps([operation, context], sort_keys=True, ensure_ascii=False).encode()).hexdigest()
        with self._lock, self.database.connection:
            row = self.database.connection.execute(
                "select id,state from observation_tasks where coalesce_key=? and state in ('queued','connecting','waiting_for_authentication','reading','observing','cancel_requested') order by created_at desc limit 1",
                (key,),
            ).fetchone()
            if row:
                return SubmittedTask(row["id"], row["state"])
            active = self._active_row()
            if active:
                raise RuntimeError(f"已有教务任务正在运行：{active['operation']}")
            identity, now = uuid.uuid4().hex, utc_now()
            self.database.connection.execute(
                "insert into observation_tasks values(?,?,?,?,?,?,?,?,?)",
                (identity, operation, key, TaskState.QUEUED.value, now, now, "{}", json.dumps(context, ensure_ascii=False), ""),
            )
            self._done[identity] = threading.Event()
            self._queue.put(identity)
            return SubmittedTask(identity, TaskState.QUEUED.value)

    def submit_execution(self, context: dict[str, Any]) -> SubmittedTask:
        """Queue a non-coalesced, explicitly confirmed execution task."""
        safe_context = sanitize_for_storage(context)
        identity, now = uuid.uuid4().hex, utc_now()
        with self._lock, self.database.connection:
            active = self._active_row()
            if active:
                raise RuntimeError(f"已有教务任务正在运行：{active['operation']}")
            self.database.connection.execute(
                "insert into execution_tasks values(?,?,?,?,?,?,?,?)",
                (identity, "execute-selection", TaskState.QUEUED.value, now, now, "{}",
                 json.dumps(safe_context, ensure_ascii=False), ""),
            )
            self._done[identity] = threading.Event()
            self._queue.put("execution:" + identity)
        return SubmittedTask(identity, TaskState.QUEUED.value)

    def _active_row(self):
        return self.database.connection.execute(
            "select id,operation,state,created_at,updated_at,progress,'observation' task_kind from observation_tasks "
            "where state in ('queued','connecting','waiting_for_authentication','reading','observing','cancel_requested') "
            "union all select id,operation,state,created_at,updated_at,progress,'execution' task_kind from execution_tasks "
            "where state in ('queued','connecting','waiting_for_authentication','reading','cancel_requested') limit 1"
        ).fetchone()

    def active_task(self) -> dict[str, Any] | None:
        with self._lock:
            row = self._active_row()
        if not row:
            return None
        return {
            "id": row["id"], "operation": row["operation"], "state": row["state"],
            "task_kind": row["task_kind"], "created_at": row["created_at"],
            "updated_at": row["updated_at"], "progress": json.loads(row["progress"]),
            "timeout_seconds": TASK_TIMEOUT_SECONDS.get(row["operation"], 180),
        }

    def session_status(self) -> dict[str, Any]:
        return {
            "state": self.session_state, "browser": self.browser_state,
            "webvpn": self.webvpn_state, "last_verified_at": self.last_verified_at,
        }

    def run_when_idle(self, action: Callable[[], Any]) -> Any:
        """Serialize login/workspace changes against every academic task."""
        with self._lock:
            active = self.database.connection.execute(
                "select 1 from observation_tasks where state in ('queued','connecting','waiting_for_authentication','reading','observing','cancel_requested') "
                "union all select 1 from execution_tasks where state in ('queued','connecting','waiting_for_authentication','reading','cancel_requested') limit 1"
            ).fetchone()
            if active:
                raise RuntimeError("请等待当前教务任务结束后再更改登录")
            return action()

    def inspect(self, identity: str) -> dict[str, Any] | None:
        with self._lock:
            row = self.database.connection.execute("select *, 'observation' as task_kind from observation_tasks where id=?", (identity,)).fetchone()
            if not row:
                row = self.database.connection.execute("select *, 'execution' as task_kind from execution_tasks where id=?", (identity,)).fetchone()
        if not row:
            return None
        return {"id": row["id"], "operation": row["operation"], "task_kind": row["task_kind"], "state": row["state"], "created_at": row["created_at"], "updated_at": row["updated_at"], "progress": json.loads(row["progress"]), "error": row["error"]}

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

    def finish(self, identity: str) -> bool:
        with self._lock:
            row = self.database.connection.execute(
                "select operation,state from observation_tasks where id=?", (identity,)
            ).fetchone()
            if not row or row["operation"] != "observe-navigation":
                return False
            if row["state"] in {TaskState.SUCCEEDED.value, TaskState.FAILED.value, TaskState.CANCELLED.value}:
                return False
            self._finished.add(identity)
            return True

    def wait(self, identity: str, timeout: float | None = None) -> bool:
        return self._done.setdefault(identity, threading.Event()).wait(timeout)

    def _update(self, identity: str, state: str, progress: dict[str, Any] | None = None, error: str = "") -> None:
        safe_error = str(sanitize_for_storage(error))[:1000]
        with self._lock, self.database.connection:
            safe_progress = sanitize_for_storage(progress or {})
            table = "execution_tasks" if self.database.connection.execute("select 1 from execution_tasks where id=?", (identity,)).fetchone() else "observation_tasks"
            self.database.connection.execute(f"update {table} set state=?,updated_at=?,progress=?,error=? where id=?", (state, utc_now(), json.dumps(safe_progress, ensure_ascii=False), safe_error, identity))

    def _run(self) -> None:
        while True:
            try:
                identity = self._queue.get(timeout=0.25)
            except queue.Empty:
                if self.gateway:
                    poll = getattr(self.gateway, "poll", None)
                    if poll is not None:
                        try:
                            poll()
                        except Exception:
                            failed_gateway = self.gateway
                            self.gateway = None
                            self.session_state = "unknown"
                            self.browser_state = "closed"
                            self.webvpn_state = "unknown"
                            try:
                                failed_gateway.close()
                            except Exception:
                                pass
                continue
            if identity is None:
                if self.gateway:
                    self.gateway.close()
                    self.gateway = None
                self.session_state = "disconnected"
                return
            if identity.startswith("execution:"):
                identity = identity.split(":", 1)[1]
            task = self.inspect(identity)
            if not task or task["state"] == TaskState.CANCELLED.value:
                continue
            operation = str(task.get("operation") or "")
            timer = threading.Timer(
                TASK_TIMEOUT_SECONDS.get(operation, 180),
                lambda: self._expired.add(identity),
            )
            timer.daemon = True
            timer.start()
            try:
                self._execute(identity)
            except Exception as error:
                task = self.inspect(identity)
                # A terminal task must never leave the session badge looking
                # like authentication is still actively waiting.
                if self.session_state in {
                    TaskState.CONNECTING.value,
                    TaskState.WAITING_FOR_AUTHENTICATION.value,
                }:
                    self.session_state = "unknown"
                    self.webvpn_state = "unknown"
                if task and task["operation"] in {"refresh-selection", "refresh-timetable", "refresh-progress"}:
                    self.database.record_failed_attempt(
                        {"refresh-selection": "selection", "refresh-timetable": "timetable", "refresh-progress": "progress"}[task["operation"]],
                        str(error),
                    )
                self._update(identity, TaskState.FAILED.value, error=str(error))
            finally:
                timer.cancel()
                if identity in self._expired:
                    current = self.inspect(identity) or {}
                    if operation == "execute-selection" and current.get("state") != TaskState.SUCCEEDED.value:
                        with self._lock:
                            row = self.database.connection.execute(
                                "select context from execution_tasks where id=?", (identity,)
                            ).fetchone()
                        if row and not self.database.unresolved_execution(json.loads(row[0]).get("section_id", "")):
                            self.database.record_execution(
                                json.loads(row[0]),
                                {"status": "unknown", "message": "选课执行超时，结果需要人工核实"},
                            )
                    if current.get("state") not in {TaskState.SUCCEEDED.value, TaskState.FAILED.value}:
                        self._update(identity, TaskState.FAILED.value, error="task timed out")
                    if self.session_state in {
                        TaskState.CONNECTING.value,
                        TaskState.WAITING_FOR_AUTHENTICATION.value,
                    }:
                        self.session_state = "unknown"
                        self.webvpn_state = "unknown"
                self._done.setdefault(identity, threading.Event()).set()

    def _execute(self, identity: str) -> None:
        with self._lock:
            execution = self.database.connection.execute("select operation,context from execution_tasks where id=?", (identity,)).fetchone()
            row = execution or self.database.connection.execute("select operation,context from observation_tasks where id=?", (identity,)).fetchone()
        operation, context = row[0], json.loads(row[1])
        # Bound every downstream browser/authentication wait by this task's
        # declared deadline; a timer flag alone cannot interrupt a blocking
        # Playwright call.
        context["operation_timeout_seconds"] = TASK_TIMEOUT_SECONDS.get(operation, 180)
        if identity in self._cancelled:
            self._update(identity, TaskState.CANCELLED.value)
            return
        if self._login_reset_error and operation != "reset-login":
            raise RuntimeError(self._login_reset_error)
        if self.gateway is None:
            self.gateway = self.gateway_factory()
            self.browser_state = "running"
        def cancelled() -> bool:
            return identity in self._cancelled or identity in self._expired or self._shutdown.is_set()
        def progress(state: str, details: dict[str, Any]) -> None:
            mapped = TaskState(state).value
            if mapped == TaskState.WAITING_FOR_AUTHENTICATION.value:
                self.session_state = TaskState.WAITING_FOR_AUTHENTICATION.value
            elif mapped == TaskState.CONNECTING.value:
                self.session_state = TaskState.CONNECTING.value
            self._update(identity, mapped, details)
        if operation == "reset-login":
            self._update(identity, TaskState.CONNECTING.value, {"message": "resetting academic login"})
            reset = getattr(self.gateway, "reset_login", None)
            try:
                if reset is None:
                    self.gateway.close()
                else:
                    reset()
            except Exception as error:
                self.gateway = None
                self.session_state = "disconnected"
                self._login_reset_error = f"登录重置失败：{str(error)[:300]}"
                raise RuntimeError(self._login_reset_error) from error
            self.gateway = None
            self._login_reset_error = ""
            self.session_state = "disconnected"
            self._update(identity, TaskState.SUCCEEDED.value, {"status": "complete"})
            return
        if operation == "launch-shell":
            self._update(identity, TaskState.CONNECTING.value)
            launcher = getattr(self.gateway, "launch_shell", None)
            if launcher is None:
                raise RuntimeError("academic gateway cannot launch the workbench shell")
            launcher(str(context.get("workbench_url", "")), progress, cancelled)
            self.session_state = "disconnected"
            self._update(identity, TaskState.SUCCEEDED.value, {"status": "complete"})
            return
        self._update(identity, TaskState.CONNECTING.value)
        self.session_state = "connecting"
        self.gateway.connect(progress, cancelled)
        if cancelled():
            if self.session_state in {
                TaskState.CONNECTING.value,
                TaskState.WAITING_FOR_AUTHENTICATION.value,
            }:
                self.session_state = "unknown"
                self.webvpn_state = "unknown"
            self._update(identity, TaskState.CANCELLED.value)
            return
        self.session_state = "connected"
        self.webvpn_state = "valid"
        self.last_verified_at = utc_now()
        if operation == "execute-selection":
            self._update(identity, TaskState.READING.value, {"target": "selection-execution"})
            result = self.gateway.execute_selection(context, progress, cancelled)
            if cancelled() or result.get("status") == "cancelled":
                self._update(identity, TaskState.CANCELLED.value)
            else:
                self.database.record_execution(context, result)
                self._update(identity, TaskState.SUCCEEDED.value, {"result": result})
            return
        if operation == "connect":
            self._update(identity, TaskState.SUCCEEDED.value, {"status": "verified", "verified_at": self.last_verified_at})
            return
        if operation == "observe-navigation":
            finished = lambda: identity in self._finished
            observed = self.gateway.observe_manual(
                ManualObservationRequest(context), progress, cancelled, finished,
            )
            trace_progress: dict[str, Any] = {"report": observed.diagnostic, "trace_incomplete": False}
            try:
                trace_progress["trace_path"] = str(self.trace_store.write(identity, observed.trace))
            except OSError as error:
                trace_progress["trace_incomplete"] = True
                trace_progress["trace_error"] = str(error)[:300]
            if cancelled() or observed.status == "cancelled":
                self._update(identity, TaskState.CANCELLED.value, trace_progress)
            elif observed.status == "complete":
                self._update(identity, TaskState.SUCCEEDED.value, trace_progress)
            else:
                self._update(identity, TaskState.FAILED.value, trace_progress, observed.error or observed.status)
            return
        self._update(identity, TaskState.READING.value)
        kind = {"refresh-selection": "selection", "refresh-timetable": "timetable", "refresh-progress": "progress"}[operation]
        if kind == "selection":
            observed: SelectionObservationResult = self.gateway.observe_selection(
                SelectionObservationRequest(context=context), progress, cancelled,
            )
            trace_progress: dict[str, Any] = {"trace_incomplete": False}
            try:
                trace_progress["trace_path"] = str(self.trace_store.write(identity, observed.trace))
            except OSError as error:
                trace_progress["trace_incomplete"] = True
                trace_progress["trace_error"] = str(error)[:300]
            if cancelled() or observed.status == "cancelled":
                self._update(identity, TaskState.CANCELLED.value, trace_progress)
            elif isinstance(observed, SelectionDiscoveryDiagnostic):
                details = {**trace_progress, "diagnostic": observed.diagnostic}
                self.database.record_failed_attempt(kind, observed.status)
                self._update(identity, TaskState.INTERFACE_UNCONFIRMED.value, details, observed.status)
            elif observed.status == "complete":
                self.database.publish_snapshot(
                    kind, str(observed.payload.get("term", context.get("term", ""))), observed.payload,
                    source=str(observed.payload.get("source_kind", "academic")), profile_id=context.get("profile_id"),
                    notice_id=context.get("notice_id"),
                )
                self._update(identity, TaskState.SUCCEEDED.value, trace_progress)
            else:
                error = observed.error or observed.status
                self.database.record_failed_attempt(kind, error)
                self._update(identity, TaskState.FAILED.value, trace_progress, error)
            return
        if kind == "progress":
            observed: ProgressObservationResult = self.gateway.observe_progress(
                ProgressObservationRequest(context=context), progress, cancelled,
            )
            trace_progress: dict[str, Any] = {"trace_incomplete": False}
            try:
                trace_progress["trace_path"] = str(self.trace_store.write(identity, observed.trace))
            except OSError as error:
                trace_progress["trace_incomplete"] = True
                trace_progress["trace_error"] = str(error)[:300]
            if cancelled() or observed.status == "cancelled":
                self._update(identity, TaskState.CANCELLED.value, trace_progress)
            elif observed.status == "complete":
                report = observed.payload.get("report") or {}
                expected_baseline = str(context.get("baseline_version", "guide-2026"))
                if not report.get("data_complete") or report.get("baseline_version") != expected_baseline:
                    error = "progress result is incomplete or uses an unexpected baseline"
                    self.database.record_failed_attempt(kind, error)
                    self._update(identity, TaskState.FAILED.value, trace_progress, error)
                    return
                self.database.publish_snapshot(
                    kind, str(observed.payload.get("term", context.get("term", ""))), observed.payload,
                    source=str(observed.payload.get("source_kind", "academic")), profile_id=context.get("profile_id"),
                    notice_id=context.get("notice_id"),
                )
                self._update(identity, TaskState.SUCCEEDED.value, trace_progress)
            else:
                error = observed.error or observed.status
                self.database.record_failed_attempt(kind, error)
                self._update(identity, TaskState.FAILED.value, trace_progress, error)
            return
        if kind == "timetable":
            observed: TimetableObservationResult = self.gateway.observe_timetable(
                TimetableObservationRequest(term=str(context.get("term", "")), context=context),
                progress,
                cancelled,
            )
            trace_progress: dict[str, Any] = {"trace_incomplete": False}
            try:
                trace_progress["trace_path"] = str(self.trace_store.write(identity, observed.trace))
            except OSError as error:
                trace_progress["trace_incomplete"] = True
                trace_progress["trace_error"] = str(error)[:300]
            result = observed.snapshot_payload()
            if cancelled() or observed.status == "cancelled":
                self._update(identity, TaskState.CANCELLED.value, trace_progress)
            elif observed.status == "complete":
                self.database.publish_snapshot(
                    kind, observed.term or str(context.get("term", "")), result,
                    source=observed.source_kind, profile_id=context.get("profile_id"),
                    notice_id=context.get("notice_id"),
                )
                self._update(identity, TaskState.SUCCEEDED.value, trace_progress)
            else:
                error = observed.error or observed.status
                self.database.record_failed_attempt(kind, error)
                self._update(identity, TaskState.FAILED.value, trace_progress, error)
            return

    def close(self) -> None:
        self._shutdown.set()
        self._queue.put(None)
        if self._thread:
            self._thread.join(timeout=5)
        # The worker owns the gateway and closes it when it consumes the sentinel.
