"""Recoverable workspace-level process lock."""

from __future__ import annotations

import json
import os
from pathlib import Path


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
                os.kill(int(data["pid"]), 0)
                return False
            except (OSError, ValueError, KeyError, json.JSONDecodeError):
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
