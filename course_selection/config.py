"""Centralized configuration for the course-selection workbench.

Reads from environment variables first, falls back to sensible defaults.
All paths are relative to the project root unless noted otherwise.
"""

from __future__ import annotations

import os
from pathlib import Path

from dotenv import load_dotenv

# Project-root .env overrides defaults; real environment variables always win.
_load_dotenv = load_dotenv(Path(__file__).resolve().parents[1] / ".env")


def _env_path(key: str, default: str) -> Path:
    raw = os.environ.get(key, default)
    return Path(raw)


# ── Workbench ────────────────────────────────────────────────────────────────

WORKBENCH_PORT = int(os.environ.get("WORKBENCH_PORT", "5000"))
WORKBENCH_HOST = os.environ.get("WORKBENCH_HOST", "127.0.0.1")
WORKBENCH_PRIVATE_ROOT = _env_path("WORKBENCH_PRIVATE_ROOT", ".private/academic-selection")
WORKBENCH_URL = f"http://{WORKBENCH_HOST}:{WORKBENCH_PORT}"

# ── Dev workbench ────────────────────────────────────────────────────────────

DEV_DEBUG_PORT = int(os.environ.get("ACADEMIC_BROWSER_DEBUG_PORT", "9222"))
DEV_ENABLE_DIAGNOSTICS = os.environ.get("ACADEMIC_WORKBENCH_DEV_DIAGNOSTICS") == "1"

# ── Course progress ──────────────────────────────────────────────────────────

PROGRESS_PROFILE_ROOT = _env_path("PROGRESS_PROFILE_ROOT", ".private/course-progress")

# ── Lab booking (CAS book) ───────────────────────────────────────────────────

CAS_BOOK_BASE_URL = os.environ.get("CAS_BOOK_BASE_URL", "http://openlab.hitwh.edu.cn")
CAS_BOOK_LOGIN_TIMEOUT = int(os.environ.get("CAS_BOOK_LOGIN_TIMEOUT", "180"))
CAS_BOOK_POLL_INTERVAL = int(os.environ.get("CAS_BOOK_POLL_INTERVAL", "5"))
CAS_BOOK_MAX_POLL_MINUTES = int(os.environ.get("CAS_BOOK_MAX_POLL_MINUTES", "120"))

# ── Browser / CDP ────────────────────────────────────────────────────────────

ACADEMIC_BROWSER_CDP_URL = os.environ.get("ACADEMIC_BROWSER_CDP_URL") or None