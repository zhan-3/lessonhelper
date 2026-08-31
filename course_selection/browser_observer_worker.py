"""Single-thread owner for sync Playwright borrowed-browser operations."""

from __future__ import annotations

import queue
import threading
from dataclasses import dataclass, field
from typing import Any

from .browser_observer import BorrowedBrowserObserver, ObserverResult


@dataclass
class _Call:
    method: str
    args: tuple[Any, ...] = ()
    done: threading.Event = field(default_factory=threading.Event)
    result: ObserverResult | None = None
    error: BaseException | None = None


class BorrowedBrowserObserverWorker:
    """Marshal every observer operation onto one Playwright owner thread."""

    def __init__(self, observer: BorrowedBrowserObserver):
        self._observer = observer
        self._calls: queue.Queue[_Call | None] = queue.Queue()
        self._closed = False
        self._lock = threading.RLock()
        self._thread: threading.Thread | None = None

    def _ensure_thread(self) -> None:
        if self._thread is None:
            self._thread = threading.Thread(target=self._run, name="borrowed-browser-observer", daemon=True)
            self._thread.start()

    def _run(self) -> None:
        while True:
            try:
                call = self._calls.get(timeout=0.02)
            except queue.Empty:
                self._observer.pump_events()
                self._observer.enforce_budgets()
                continue
            if call is None:
                return
            try:
                call.result = getattr(self._observer, call.method)(*call.args)
            except BaseException as error:
                call.error = error
            finally:
                call.done.set()

    def _call(self, method: str, *args: Any) -> ObserverResult:
        with self._lock:
            if self._closed and method != "shutdown":
                return ObserverResult("failed", "observer worker is shut down", next_actions=("reload and connect explicitly",))
            self._ensure_thread()
            call = _Call(method, args)
            self._calls.put(call)
        call.done.wait()
        if call.error is not None:
            raise call.error
        assert call.result is not None
        return call.result

    def connect(self, endpoint: str) -> ObserverResult:
        return self._call("connect", endpoint)

    def inspect(self) -> ObserverResult:
        return self._call("inspect")

    def start_observation(self) -> ObserverResult:
        return self._call("start_observation")

    def checkpoint(self) -> ObserverResult:
        return self._call("checkpoint")

    def stop_observation(self) -> ObserverResult:
        return self._call("stop_observation")

    def cancel_observation(self) -> ObserverResult:
        return self._call("cancel_observation")

    def disconnect(self) -> ObserverResult:
        return self._call("disconnect")

    def shutdown(self) -> ObserverResult:
        with self._lock:
            if self._closed:
                return ObserverResult("disconnected", "observer worker is already shut down", {"connection": "borrowed"})
            if self._thread is None:
                self._closed = True
                return ObserverResult("disconnected", "observer worker shut down before use", {"connection": "borrowed"})
            result = self._call("shutdown")
            self._closed = True
            self._calls.put(None)
        assert self._thread is not None
        self._thread.join(timeout=5)
        return result
