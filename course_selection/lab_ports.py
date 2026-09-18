"""Abstract seams shared by the lab booking and contract cores.

These declarations carry no I/O.  The application core depends on this module
instead of on :mod:`course_selection.lab_transport`, so the dependency direction
stays ``implementation -> abstraction``.  The concrete backends live in
``lab_transport``, which re-exports every name here for convenience.
"""

from __future__ import annotations

from collections.abc import Callable, Mapping
from typing import Any, Protocol

DEFAULT_ORIGIN = "http://openlab.hitwh.edu.cn"
TOKEN_HEADER = "vctchauthorization"
CONTENT_TYPE = "application/x-www-form-urlencoded;charset=UTF-8"

Fetch = Callable[[str, bytes, Mapping[str, str], float], tuple[int, str]]


class LabTransport(Protocol):
    """One authenticated channel to a single teaching-center app."""

    center: str

    def call(self, path: str, form: Mapping[str, Any] | None = None) -> dict[str, Any]: ...

    def close(self) -> None: ...
