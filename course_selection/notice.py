"""Parse and persist the user-confirmed selection window."""

from __future__ import annotations

import re
from dataclasses import asdict, dataclass, replace
from datetime import datetime, timezone
from html.parser import HTMLParser
from pathlib import Path
from typing import Any
from urllib.parse import urlparse
from urllib.request import Request, urlopen

from .categories import NOTICE_CATEGORY_PATTERNS

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

_WINDOW_RE = re.compile(
    r"^(20\d{2})年(\d{1,2})月(\d{1,2})日(\d{1,2}):(\d{2})"
    r"\s*[-—至]\s*"
    r"(20\d{2})年(\d{1,2})月(\d{1,2})日(\d{1,2}):(\d{2})$"
)


@dataclass(frozen=True)
class SelectionWindow:
    opens_at: str
    closes_at: str
    grades: tuple[str, ...]
    audience: str
    category_text: str
    category_codes: tuple[str, ...]
    action: str
    method: str


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
    windows: tuple[SelectionWindow, ...] = ()

    @property
    def missing_fields(self) -> tuple[str, ...]:
        if self.windows:
            return ("term",) if not self.term else ()
        return tuple(field for field in REQUIRED_FIELDS if not getattr(self, field))


def _category_codes(value: str) -> tuple[str, ...]:
    found: list[str] = []
    matches: list[tuple[int, str]] = []
    for phrase, code in NOTICE_CATEGORY_PATTERNS:
        position = value.find(phrase)
        if position >= 0:
            matches.append((position, code))
    for _, code in sorted(matches):
        if code not in found:
            found.append(code)
    return tuple(found)


def notice_selection_categories(
    notice: SelectionNotice, *, grade: str | None = None
) -> tuple[str, ...]:
    """Resolve only course categories explicitly named by a confirmed notice."""
    if notice.windows:
        found: list[str] = []
        for window in notice.windows:
            if window.action != "selection" or window.method != "academic_system":
                continue
            if grade and grade not in window.grades:
                continue
            for code in window.category_codes:
                if code not in found:
                    found.append(code)
        return tuple(found)
    primary = f"{notice.selection_type}\n{notice.title}"
    sources = (primary, notice.source_text)
    found: list[str] = []
    for source in sources:
        for code in _category_codes(source):
            if code not in found:
                found.append(code)
        if found:
            break
    return tuple(found)


def notice_selection_windows(
    notice: SelectionNotice, *, grade: str | None = None
) -> tuple[SelectionWindow, ...]:
    """Return only confirmed-notice windows usable for read-only selection queries."""
    if not notice.windows:
        return ()
    return tuple(
        window
        for window in notice.windows
        if window.action == "selection"
        and window.method == "academic_system"
        and (not grade or grade in window.grades)
    )


def notice_selection_window_map(
    notice: SelectionNotice, *, grade: str | None = None
) -> dict[str, tuple[SelectionWindow, ...]]:
    """Group confirmed, profile-matched selection windows by canonical category."""
    grouped: dict[str, list[SelectionWindow]] = {}
    for window in notice_selection_windows(notice, grade=grade):
        for code in window.category_codes:
            grouped.setdefault(code, []).append(window)
    return {code: tuple(windows) for code, windows in grouped.items()}


def notice_semester_label(notice: SelectionNotice) -> str:
    value = re.sub(r"\s+", "", notice.term)
    return value.replace("年", "").replace("学期", "")


def _window_datetime(groups: tuple[str, ...], offset: int) -> str:
    year, month, day, hour, minute = (int(value) for value in groups[offset : offset + 5])
    return datetime(year, month, day, hour, minute).astimezone().strftime("%Y-%m-%d %H:%M")


def _window_action(category_text: str, method_text: str) -> str:
    if "退课" in category_text:
        return "drop"
    if "选课" in category_text:
        return "selection"
    if any(marker in category_text for marker in ("申请", "重修", "辅修", "免听", "补修")):
        return "application"
    if "调整" in category_text:
        return "adjustment"
    if "新教务系统" in method_text and _category_codes(category_text):
        return "selection"
    return "other"


def parse_selection_windows(text: str) -> tuple[SelectionWindow, ...]:
    lines = [line.strip() for line in text.splitlines() if line.strip()]
    starts = [index for index, line in enumerate(lines) if _WINDOW_RE.fullmatch(line)]
    windows: list[SelectionWindow] = []
    for position, start in enumerate(starts):
        match = _WINDOW_RE.fullmatch(lines[start])
        if match is None:
            continue
        end = starts[position + 1] if position + 1 < len(starts) else len(lines)
        block = lines[start + 1 : end]
        audience_lines: list[str] = []
        while block and (re.search(r"20\d{2}级", block[0]) or "结业生" in block[0]):
            audience_lines.append(block.pop(0))
        if not block:
            continue
        category_text = block.pop(0)
        method_text = " ".join(block)
        grades = tuple(dict.fromkeys(re.findall(r"(20\d{2})级", "".join(audience_lines))))
        method = (
            "academic_system"
            if "新教务系统" in method_text
            else "manual"
        )
        windows.append(
            SelectionWindow(
                opens_at=_window_datetime(match.groups(), 0),
                closes_at=_window_datetime(match.groups(), 5),
                grades=grades,
                audience="".join(audience_lines),
                category_text=category_text,
                category_codes=_category_codes(category_text),
                action=_window_action(category_text, method_text),
                method=method,
            )
        )
    return tuple(windows)


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

    from course_progress.session import AcademicBrowserSession

    parsed = urlparse(source_url)
    if parsed.scheme not in {"http", "https"} or not parsed.netloc:
        raise ValueError("通知链接必须是 HTTP 或 HTTPS 地址")
    if login_timeout_seconds <= 0:
        raise ValueError("认证等待时间必须大于 0")

    with sync_playwright() as playwright, AcademicBrowserSession(
        playwright,
        browser_name=browser,
        profile_root=profile_root,
        persistent=False,
    ) as session:
        page = session.open_authenticated(
            source_url, timeout_seconds=login_timeout_seconds
        )
        page.wait_for_timeout(500)
        text = page.locator("body").inner_text(timeout=10_000).strip()
        if not text:
            raise ValueError("通知页面没有读取到正文")
        return text


def _first_line(text: str) -> str:
    return next((line.strip() for line in text.splitlines() if line.strip()), "")


def _notice_title(text: str) -> str:
    return next(
        (
            line.strip()
            for line in text.splitlines()
            if line.strip().startswith("关于") and "通知" in line
        ),
        _first_line(text),
    )


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
    windows = parse_selection_windows(normalized)
    opens_at, closes_at = _find_dates(normalized)
    if windows:
        opens_at, closes_at = windows[0].opens_at, windows[0].closes_at
    return SelectionNotice(
        source_kind=source_kind,
        source_url=source_url.strip(),
        title=_notice_title(normalized),
        term=_find_term(normalized),
        selection_type="多类别" if windows else _find_selection_type(normalized),
        audience=_find_line(normalized, ("面向", "对象", "适用")),
        opens_at=opens_at,
        closes_at=closes_at,
        entry_hint=_find_line(normalized, ("学生选课", "选课入口", "选课系统")),
        restrictions=_find_line(normalized, ("要求", "限制", "说明")),
        status="pending_confirmation",
        source_text=normalized,
        created_at=created_at or _now(),
        windows=windows,
    )


def update_notice(notice: SelectionNotice, **fields: str) -> SelectionNotice:
    allowed = set(SelectionNotice.__dataclass_fields__) - {
        "status",
        "created_at",
        "windows",
    }
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
    data["windows"] = tuple(
        SelectionWindow(
            **{
                **item,
                "grades": tuple(item.get("grades", ())),
                "category_codes": tuple(item.get("category_codes", ())),
            }
        )
        for item in data.get("windows", ())
    )
    return SelectionNotice(**data)
