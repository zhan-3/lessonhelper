"""Parse and persist the user-confirmed selection window."""

from __future__ import annotations

import re
from html.parser import HTMLParser
from urllib.parse import urlparse
from urllib.request import Request, urlopen
from dataclasses import asdict, dataclass, replace
from datetime import datetime, timezone
from pathlib import Path
from typing import Any


REQUIRED_FIELDS = ("term", "selection_type", "opens_at", "closes_at")
SELECTION_KEYWORDS = (
    "大学外语",
    "文化素质",
    "跨专业",
    "创新创业",
    "创新研修",
    "重修",
    "辅修",
    "补修",
    "选课",
)


@dataclass(frozen=True)
class SelectionNotice:
    source_kind: str
    source_url: str
    title: str
    term: str
    selection_type: str
    audience: str
    opens_at: str
    closes_at: str
    entry_hint: str
    restrictions: str
    status: str
    source_text: str
    created_at: str

    @property
    def missing_fields(self) -> tuple[str, ...]:
        return tuple(field for field in REQUIRED_FIELDS if not getattr(self, field))


def _now() -> str:
    return datetime.now(timezone.utc).isoformat(timespec="seconds")


class _TextExtractor(HTMLParser):
    def __init__(self) -> None:
        super().__init__()
        self.parts: list[str] = []

    def handle_data(self, data: str) -> None:
        value = data.strip()
        if value:
            self.parts.append(value)

    def text(self) -> str:
        return "\n".join(self.parts)


def fetch_notice_text(source_url: str, *, timeout_seconds: int = 10) -> str:
    parsed = urlparse(source_url)
    if parsed.scheme not in {"http", "https"} or not parsed.netloc:
        raise ValueError("通知链接必须是 HTTP 或 HTTPS 地址")
    request = Request(
        source_url,
        headers={"User-Agent": "academic-course-selection/0.1"},
    )
    with urlopen(request, timeout=timeout_seconds) as response:
        payload = response.read()
        charset = response.headers.get_content_charset() or "utf-8"
    parser = _TextExtractor()
    parser.feed(payload.decode(charset, errors="replace"))
    return parser.text()


def fetch_notice_text_in_browser(
    source_url: str,
    *,
    profile_root: Path = Path(".private/course-progress"),
    browser: str = "chromium",
    login_timeout_seconds: int = 600,
) -> str:
    """Read a notice in the visible, persistent academic browser session."""
    from playwright.sync_api import sync_playwright

    from course_progress.explorer import _is_login_url, launch_browser_context, resolve_profile_dir

    parsed = urlparse(source_url)
    if parsed.scheme not in {"http", "https"} or not parsed.netloc:
        raise ValueError("通知链接必须是 HTTP 或 HTTPS 地址")
    if login_timeout_seconds <= 0:
        raise ValueError("认证等待时间必须大于 0")

    profile_dir = resolve_profile_dir(profile_root.resolve())
    with sync_playwright() as playwright:
        context = launch_browser_context(playwright, browser, profile_dir)
        try:
            page = context.pages[0] if context.pages else context.new_page()
            page.goto(source_url, wait_until="domcontentloaded", timeout=60_000)
            deadline = datetime.now().timestamp() + login_timeout_seconds
            while _is_login_url(page.url) and datetime.now().timestamp() < deadline:
                page.wait_for_timeout(500)
            if _is_login_url(page.url):
                raise ValueError("等待统一身份认证超时，请先在浏览器中完成登录")
            page.wait_for_timeout(500)
            text = page.locator("body").inner_text(timeout=10_000).strip()
            if not text:
                raise ValueError("通知页面没有读取到正文")
            return text
        finally:
            context.close()


def _first_line(text: str) -> str:
    return next((line.strip() for line in text.splitlines() if line.strip()), "")


def _find_term(text: str) -> str:
    match = re.search(
        r"((?:20\d{2})(?:[-—至]\s*20\d{2})?年(?:春季|夏季|秋季|冬季)学期)",
        text,
    )
    return match.group(1).replace(" ", "") if match else ""


def _find_dates(text: str) -> tuple[str, str]:
    date_pattern = r"20\d{2}年\s*\d{1,2}月\s*\d{1,2}日(?:\s*\d{1,2}:\d{2})?"
    dates = [re.sub(r"\s+", "", value) for value in re.findall(date_pattern, text)]
    if len(dates) < 2:
        return (dates[0], "") if dates else ("", "")
    return dates[0], dates[1]


def _find_selection_type(text: str) -> str:
    for keyword in SELECTION_KEYWORDS:
        if keyword in text:
            return keyword
    return ""


def _find_line(text: str, markers: tuple[str, ...]) -> str:
    for line in text.splitlines():
        stripped = line.strip()
        if any(marker in stripped for marker in markers):
            return stripped
    return ""


def parse_notice(
    text: str,
    *,
    source_url: str = "",
    source_kind: str = "official",
    created_at: str | None = None,
) -> SelectionNotice:
    """Extract safe defaults; unresolved required fields remain empty."""

    normalized = text.strip()
    opens_at, closes_at = _find_dates(normalized)
    return SelectionNotice(
        source_kind=source_kind,
        source_url=source_url.strip(),
        title=_first_line(normalized),
        term=_find_term(normalized),
        selection_type=_find_selection_type(normalized),
        audience=_find_line(normalized, ("面向", "对象", "适用")),
        opens_at=opens_at,
        closes_at=closes_at,
        entry_hint=_find_line(normalized, ("学生选课", "选课入口", "选课系统")),
        restrictions=_find_line(normalized, ("要求", "限制", "说明")),
        status="pending_confirmation",
        source_text=normalized,
        created_at=created_at or _now(),
    )


def update_notice(notice: SelectionNotice, **fields: str) -> SelectionNotice:
    allowed = set(SelectionNotice.__dataclass_fields__) - {"status", "created_at"}
    unknown = set(fields) - allowed
    if unknown:
        raise ValueError(f"不支持的通知字段：{', '.join(sorted(unknown))}")
    return replace(notice, **{key: value.strip() for key, value in fields.items()})


def confirm_notice(notice: SelectionNotice) -> SelectionNotice:
    missing = notice.missing_fields
    if missing:
        raise ValueError(f"选课通知仍缺少：{', '.join(missing)}")
    return replace(notice, status="confirmed")


def notice_to_dict(notice: SelectionNotice) -> dict[str, Any]:
    return asdict(notice)


def save_notice(path: Path, notice: SelectionNotice) -> None:
    import json

    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(
        json.dumps(notice_to_dict(notice), ensure_ascii=False, indent=2),
        encoding="utf-8",
    )


def load_notice(path: Path) -> SelectionNotice:
    import json

    data = json.loads(path.read_text(encoding="utf-8"))
    return SelectionNotice(**data)
