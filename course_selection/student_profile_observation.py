"""Structure-only discovery of candidate student-profile read contracts."""

from __future__ import annotations

import json
import re
import time
from dataclasses import dataclass
from datetime import datetime, timedelta, timezone
from pathlib import Path
from typing import Any
from urllib.parse import urljoin, urlsplit, urlunsplit

from playwright.sync_api import Error, Response, sync_playwright

from course_progress.explorer import DEFAULT_PORTAL_URL
from course_progress.sanitizer import sanitize_url
from course_progress.session import AcademicBrowserSession, webvpn_api_get

from .discovery import find_academic_portal_redirect
from .manual_observation import ManualObservationPolicy

PROFILE_FIELD_GROUPS = {
    "student_name": {"xm", "xsxm", "xsmc", "studentname", "displayname", "fullname", "姓名"},
    "student_number": {"xh", "xsxh", "xgh", "studentno", "studentnumber", "username", "学号"},
    "grade": {"nj", "rxnj", "grade", "enrollmentyear", "入学年级", "年级"},
    "major": {"zy", "zymc", "major", "majorname", "专业"},
    "academic_level": {"pycc", "cc", "academiclevel", "degreelevel", "培养层次", "层次"},
    "campus": {"xq", "xqmc", "campus", "campusname", "校区"},
}
_MAX_RESPONSE_BYTES = 2 * 1024 * 1024
_WEBVPN_HOST = "webvpn.hitwh.edu.cn"
_WEBVPN_APP_PATH = re.compile(r"^/https?/[0-9a-f]+(?:/|$)", re.IGNORECASE)


def _normalized_field(value: str) -> str:
    return re.sub(r"[^0-9a-z\u4e00-\u9fff]", "", value.lower())


def _safe_key(value: object) -> str:
    key = str(value)
    normalized = _normalized_field(key)
    known = {alias for aliases in PROFILE_FIELD_GROUPS.values() for alias in aliases}
    if normalized in known or re.fullmatch(r"[A-Za-z_][A-Za-z0-9_-]{0,63}", key):
        return key
    return "[dynamic-key]"


def _safe_structure(value: Any, *, depth: int = 0) -> dict[str, Any]:
    if depth >= 4:
        return {"type": type(value).__name__}
    if isinstance(value, dict):
        fields: dict[str, Any] = {}
        for key, child in value.items():
            safe = _safe_key(key)
            fields.setdefault(safe, _safe_structure(child, depth=depth + 1))
        return {"type": "object", "fields": fields}
    if isinstance(value, list):
        return {
            "type": "array",
            "length": len(value),
            "item": _safe_structure(value[0], depth=depth + 1) if value else None,
        }
    return {"type": type(value).__name__}


def _safe_url(value: str) -> str:
    parts = urlsplit(sanitize_url(value))
    return urlunsplit((parts.scheme, parts.netloc, parts.path, parts.query, ""))


def _field_paths(value: Any, prefix: str = "") -> list[str]:
    paths: list[str] = []
    if isinstance(value, dict):
        for key, child in value.items():
            safe_key = _safe_key(key)
            path = f"{prefix}.{safe_key}" if prefix else safe_key
            paths.append(path)
            paths.extend(_field_paths(child, path))
    elif isinstance(value, list) and value:
        paths.extend(_field_paths(value[0], f"{prefix}[]"))
    return paths


def profile_candidate_from_payload(
    *, url: str, method: str, status: int, payload: Any
) -> dict[str, Any] | None:
    """Return value-free profile evidence when response field names are relevant."""
    matches: dict[str, list[str]] = {}
    for path in _field_paths(payload):
        field = _normalized_field(path.rsplit(".", 1)[-1].replace("[]", ""))
        for group, aliases in PROFILE_FIELD_GROUPS.items():
            if field in aliases:
                matches.setdefault(group, []).append(path)
    if not matches:
        return None
    score = sum(3 if group in {"student_name", "student_number"} else 2 for group in matches)
    if len(matches) >= 2:
        score += len(matches) * 2
    return {
        "url": _safe_url(url),
        "method": method.upper(),
        "status": status,
        "score": score,
        "matched_fields": {key: sorted(set(value)) for key, value in sorted(matches.items())},
        "response_structure": _safe_structure(payload),
    }


def cleanup_expired_analyses(
    root: Path, *, now: datetime | None = None, retention_days: int = 7
) -> int:
    """Remove old timestamped private analysis directories."""
    if retention_days <= 0:
        raise ValueError("retention_days must be positive")
    now = now or datetime.now(timezone.utc)
    cutoff = now - timedelta(days=retention_days)
    removed = 0
    if not root.is_dir():
        return removed
    for directory in root.iterdir():
        if not directory.is_dir():
            continue
        modified = datetime.fromtimestamp(directory.stat().st_mtime, timezone.utc)
        if modified >= cutoff:
            continue
        import shutil

        shutil.rmtree(directory)
        removed += 1
    return removed


def _validated_academic_entry(current_url: str, redirect: str) -> str:
    target = urljoin(current_url, redirect)
    parsed = urlsplit(target)
    if parsed.scheme != "https" or (parsed.hostname or "").lower() != _WEBVPN_HOST:
        raise ValueError("portal returned a non-WebVPN academic entry")
    if not _WEBVPN_APP_PATH.match(parsed.path):
        raise ValueError("portal returned an unrecognized academic proxy path")
    return target


@dataclass
class StudentProfileInterfaceAnalyzer:
    profile_root: Path
    output_root: Path
    portal_url: str = DEFAULT_PORTAL_URL
    login_timeout_seconds: int = 600
    wait_seconds: int = 60

    def run(self) -> Path:
        if self.login_timeout_seconds <= 0 or self.wait_seconds <= 0:
            raise ValueError("wait times must be positive")
        target_root = self.output_root / "student-profile"
        cleanup_expired_analyses(target_root)
        stamp = datetime.now(timezone.utc).strftime("%Y%m%d-%H%M%S")
        output = target_root / stamp
        output.mkdir(parents=True, exist_ok=False)
        candidates: list[dict[str, Any]] = []
        blocked: list[dict[str, Any]] = []
        policy = ManualObservationPolicy(())
        summary: dict[str, Any] = {
            "target": "student-profile",
            "started_at": datetime.now(timezone.utc).isoformat(timespec="seconds"),
            "raw_identity_saved": False,
            "webvpn_ready": False,
            "academic_entry_ready": False,
        }

        with sync_playwright() as playwright:  # noqa: SIM117 - session depends on playwright
            with AcademicBrowserSession(
                playwright,
                browser_name="chromium",
                profile_root=self.profile_root,
                persistent=True,
            ) as session:
                context = session.context
                if context is None:
                    raise RuntimeError("browser context is unavailable")

                def guard(route) -> None:
                    request = route.request
                    allowed, evidence = policy.inspect_request(
                        request.method,
                        request.url,
                        request.post_data,
                        request.headers.get("content-type", ""),
                        request.resource_type,
                    )
                    if allowed:
                        route.continue_()
                    else:
                        blocked.append(evidence)
                        route.abort("blockedbyclient")

                def capture(response: Response) -> None:
                    request = response.request
                    if request.resource_type not in {"xhr", "fetch"}:
                        return
                    content_type = response.headers.get("content-type", "").lower()
                    if "json" not in content_type:
                        return
                    try:
                        body = response.body()
                        if len(body) > _MAX_RESPONSE_BYTES:
                            return
                        payload = json.loads(body.decode("utf-8-sig"))
                    except (Error, UnicodeDecodeError, json.JSONDecodeError):
                        return
                    candidate = profile_candidate_from_payload(
                        url=response.url,
                        method=request.method,
                        status=response.status,
                        payload=payload,
                    )
                    if candidate is not None:
                        candidates.append(candidate)

                context.route("**/*", guard)
                context.on("response", capture)
                try:
                    page = session.open_authenticated(
                        self.portal_url,
                        timeout_seconds=self.login_timeout_seconds,
                    )
                    info_response = webvpn_api_get(
                        context, "https://webvpn.hitwh.edu.cn/user/info"
                    )
                    portal_response = webvpn_api_get(
                        context, "https://webvpn.hitwh.edu.cn/user/portal_groups"
                    )
                    if info_response.status != 200 or portal_response.status != 200:
                        raise RuntimeError("WebVPN capability verification failed")
                    info_payload = info_response.json()
                    portal_payload = portal_response.json()
                    summary["webvpn_ready"] = True
                    for url, payload in (
                        ("https://webvpn.hitwh.edu.cn/user/info", info_payload),
                        ("https://webvpn.hitwh.edu.cn/user/portal_groups", portal_payload),
                    ):
                        candidate = profile_candidate_from_payload(
                            url=url, method="GET", status=200, payload=payload
                        )
                        if candidate is not None:
                            candidates.append(candidate)
                    redirect = find_academic_portal_redirect(portal_payload)
                    if not redirect:
                        raise RuntimeError("verified academic entry was not found")
                    academic_entry = _validated_academic_entry(page.url, redirect)
                    summary["academic_entry"] = _safe_url(academic_entry)
                    page = session.open_authenticated(
                        academic_entry,
                        timeout_seconds=self.login_timeout_seconds,
                        page=page,
                    )
                    summary["academic_entry_ready"] = True
                    page.bring_to_front()
                    print("学生画像接口分析已启动；可在可见教务页面中进入个人/学籍信息页面。")
                    deadline = time.monotonic() + self.wait_seconds
                    while time.monotonic() < deadline:
                        page.wait_for_timeout(500)
                finally:
                    context.unroute("**/*", guard)
                    context.remove_listener("response", capture)

        unique: dict[tuple[str, str], dict[str, Any]] = {}
        for candidate in candidates:
            key = (candidate["method"], candidate["url"])
            if key not in unique or candidate["score"] > unique[key]["score"]:
                unique[key] = candidate
        ordered = sorted(unique.values(), key=lambda item: (-item["score"], item["url"]))
        summary.update(
            finished_at=datetime.now(timezone.utc).isoformat(timespec="seconds"),
            candidate_count=len(ordered),
            blocked_request_count=len(blocked),
        )
        (output / "candidate-contracts.json").write_text(
            json.dumps(ordered, ensure_ascii=False, indent=2), encoding="utf-8"
        )
        (output / "blocked-requests.json").write_text(
            json.dumps(blocked, ensure_ascii=False, indent=2), encoding="utf-8"
        )
        (output / "summary.json").write_text(
            json.dumps(summary, ensure_ascii=False, indent=2), encoding="utf-8"
        )
        return output
