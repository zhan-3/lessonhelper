"""Recoverable workspace-level process lock."""

from __future__ import annotations

import ctypes
import json
import os
from pathlib import Path


def _process_alive(pid: int) -> bool:
    """Return whether a PID refers to a live process (cross-platform).

    ``os.kill(pid, 0)`` is the POSIX existence check, but on Windows it raises
    instead of returning, so use the Win32 process handle API there.
    """
    if pid <= 0:
        return False
    if os.name != "nt":
        try:
            os.kill(pid, 0)
        except ProcessLookupError:
            return False
        except PermissionError:
            return True
        return True
    PROCESS_QUERY_LIMITED_INFORMATION = 0x1000
    STILL_ACTIVE = 259
    kernel32 = ctypes.windll.kernel32
    handle = kernel32.OpenProcess(PROCESS_QUERY_LIMITED_INFORMATION, False, pid)
    if not handle:
        return False
    try:
        exit_code = ctypes.c_ulong()
        if not kernel32.GetExitCodeProcess(handle, ctypes.byref(exit_code)):
            return False
        return exit_code.value == STILL_ACTIVE
    finally:
        kernel32.CloseHandle(handle)


class WorkspaceLock:
    def __init__(self, root: Path, port: int):
        self.path = root / "workbench.lock"
        self.port = port

    def acquire(self) -> bool:
        self.path.parent.mkdir(parents=True, exist_ok=True)
        try:
            descriptor = os.open(self.path, os.O_CREAT | os.O_EXCL | os.O_WRONLY)
        except FileExistsError:
            try:
                data = json.loads(self.path.read_text(encoding="utf-8"))
                if _process_alive(int(data["pid"])):
                    return False
            except (OSError, ValueError, KeyError, json.JSONDecodeError):
                pass
            # Stale lock (dead PID) or unreadable content: drop it and retry.
            self.path.unlink(missing_ok=True)
            return self.acquire()
        with os.fdopen(descriptor, "w", encoding="utf-8") as stream:
            json.dump({"pid": os.getpid(), "port": self.port}, stream)
        return True

    def release(self) -> None:
        try:
            data = json.loads(self.path.read_text(encoding="utf-8"))
            if int(data.get("pid", -1)) == os.getpid():
                self.path.unlink(missing_ok=True)
        except (OSError, ValueError, json.JSONDecodeError):
            return
