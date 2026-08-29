"""Read-only observation of a confirmed academic selection entry."""

from __future__ import annotations

import json
import re
import time
from dataclasses import asdict, dataclass
from datetime import datetime, timezone
from html.parser import HTMLParser
from pathlib import Path
from typing import Any, Iterable

from course_progress.sanitizer import (
    is_sensitive_key,
    sanitize_data,
    sanitize_request_body,
    sanitize_url,
)


STATUS_LOGIN_REQUIRED = "login_required"
STATUS_ENTRY_UNREACHABLE = "entry_unreachable"
STATUS_ROUND_NOT_OPEN = "round_not_open"
STATUS_NO_MATCHING_ROUND = "no_matching_round"
STATUS_INTERFACE_UNCONFIRMED = "interface_unconfirmed"
STATUS_EMPTY = "empty"
STATUS_READY = "ready"

_ID_KEYS = ("courseId", "courseCode", "course_id", "id", "kch", "kcbh", "jxbId")
_NAME_KEYS = ("courseName", "course_name", "name", "kcmc", "kcm")
_CREDIT_KEYS = ("credits", "credit", "xf", "courseCredit")
_TEACHER_KEYS = ("teacher", "teacherName", "jsxm", "teacher_name")
_TIME_KEYS = ("time", "classTime", "sksj", "kcsj", "schedule")
_CAPACITY_KEYS = ("capacity", "remaining", "krl", "syks", "available")
_SELECTED_KEYS = ("selected", "isSelected", "yixuan", "alreadySelected")
_CATEGORY_KEYS = ("category", "courseType", "kclb", "course_category")
_SELECTION_URL_MARKERS = ("选课", "selection", "course", "xk")


@dataclass(frozen=True)
class SelectionCourseSection:
    identity: str
    name: str
    credits: str
    teacher: str
    time: str
    capacity: str
    selected: bool | None
    category: str
    course_code: str = ""
    campus: str = ""
    department: str = ""
    hours: str = ""
    prerequisites: str = ""
    audience: str = ""
    notes: str = ""
    requirements: str = ""
    selected_count: str = ""
    capacity_count: str = ""
    # Only a page-provided saveXsxk identity may drive execution.
    action_rwh: str = ""
    action_name: str = ""
    execution_ready: bool = False
    meetings: tuple[dict[str, Any], ...] = ()


@dataclass(frozen=True)
class SelectionObservation:
    status: str
    request_url: str
    method: str
    message: str
    sections: tuple[SelectionCourseSection, ...]


@dataclass(frozen=True)
class SelectionInterfaceContract:
    method: str
    url: str
    status_code: int
    request_body: str
    response_fields: tuple[str, ...]
    observed_status: str
    captured_at: str


def _now() -> str:
    return datetime.now(timezone.utc).isoformat(timespec="seconds")


def _first_value(item: dict[str, Any], keys: Iterable[str], default: str = "") -> str:
    for key in keys:
        value = item.get(key)
        if value is not None and value != "":
            return str(value).strip()
    return default


def _bool_value(item: dict[str, Any], keys: Iterable[str]) -> bool | None:
    for key in keys:
        value = item.get(key)
        if isinstance(value, bool):
            return value
        if isinstance(value, str) and value.lower() in {"true", "false"}:
            return value.lower() == "true"
    return None


def _walk_dicts(value: Any) -> Iterable[dict[str, Any]]:
    if isinstance(value, dict):
        yield value
        for child in value.values():
            yield from _walk_dicts(child)
    elif isinstance(value, list):
        for child in value:
            yield from _walk_dicts(child)


def extract_course_sections(payload: Any) -> tuple[SelectionCourseSection, ...]:
    sections: list[SelectionCourseSection] = []
    seen: set[str] = set()
    for item in _walk_dicts(payload):
        identity = _first_value(item, _ID_KEYS)
        name = _first_value(item, _NAME_KEYS)
        if not identity or not name or identity in seen:
            continue
        seen.add(identity)
        sections.append(
            SelectionCourseSection(
                identity=identity,
                name=name,
                credits=_first_value(item, _CREDIT_KEYS),
                teacher=_first_value(item, _TEACHER_KEYS),
                time=_first_value(item, _TIME_KEYS),
                capacity=_first_value(item, _CAPACITY_KEYS),
                selected=_bool_value(item, _SELECTED_KEYS),
                category=_first_value(item, _CATEGORY_KEYS),
            )
        )
    return tuple(sections)


@dataclass
class _HtmlCell:
    text: str
    attributes: str


class _SelectionTableParser(HTMLParser):
    def __init__(self):
        super().__init__(convert_charrefs=True)
        self.rows: list[list[_HtmlCell]] = []
        self._row: list[_HtmlCell] | None = None
        self._cell_text: list[str] | None = None
        self._cell_attributes: list[str] = []

    def handle_starttag(self, tag: str, attrs) -> None:
        if tag == "tr" and self._row is None:
            self._row = []
        elif tag in {"th", "td"} and self._row is not None and self._cell_text is None:
            self._cell_text = []
            self._cell_attributes = [f"{key}={value or ''}" for key, value in attrs]
        elif self._cell_text is not None:
            self._cell_attributes.extend(
                f"{key}={value or ''}" for key, value in attrs
            )

    def handle_data(self, data: str) -> None:
        if self._cell_text is not None:
            self._cell_text.append(data)

    def handle_endtag(self, tag: str) -> None:
        if tag in {"th", "td"} and self._row is not None and self._cell_text is not None:
            self._row.append(
                _HtmlCell(
                    text=" ".join("".join(self._cell_text).split()),
                    attributes=" ".join(self._cell_attributes),
                )
            )
            self._cell_text = None
            self._cell_attributes = []
        elif tag == "tr" and self._row is not None:
            if self._row:
                self.rows.append(self._row)
            self._row = None


class _PageTextParser(HTMLParser):
    def __init__(self):
        super().__init__(convert_charrefs=True)
        self.parts: list[str] = []

    def handle_data(self, data: str) -> None:
        if data.strip():
            self.parts.append(data.strip())


def _normalise_header(value: str) -> str:
    return re.sub(r"[\s↑↓]", "", value)


def _split_capacity(value: str) -> tuple[str, str]:
    matches = re.findall(r"(\d+)\s*/\s*(\d+)", value)
    return matches[0] if len(matches) == 1 else ("", "")


def _teacher_from_class_info(value: str) -> str:
    match = re.search(r"(?:教师|老师)[:：]\s*([^,，;；\s]+)", value)
    if match:
        return match.group(1)
    first_segment = value.split("◇", 1)[0].strip()
    if first_segment and not first_segment.startswith(("上课信息", "[")):
        return re.sub(r"\d+$", "", first_segment)
    return ""


def _schedule_from_class_info(value: str) -> str:
    if "◇" in value:
        value = value.split("◇", 1)[1]
    return re.sub(r"^上课信息[:：]\s*", "", value).strip("◇ ")


def _schedule_meetings(value: str) -> tuple[dict[str, Any], ...]:
    """Normalize only explicit Chinese weekday/period facts; ambiguity stays unknown."""
    weekdays = "一二三四五六日"
    meetings: list[dict[str, Any]] = []
    for match in re.finditer(
        r"(?:\[([^\]]+?)周\])?[^星期周]*?(?:星期|周)\s*([一二三四五六日天1-7])"
        r"[^第\d]*第?\s*(\d+)\s*[,，\-~至]\s*(\d+)\s*节",
        value,
    ):
        raw_weeks, raw_day, raw_start, raw_end = match.groups()
        weeks: list[int] = []
        if raw_weeks:
            parity = "odd" if "单" in raw_weeks else "even" if "双" in raw_weeks else "all"
            for token in re.split(r"[,，]", re.sub(r"[单双]", "", raw_weeks)):
                bounds = re.match(r"\s*(\d+)\s*[-~至]\s*(\d+)\s*$", token)
                values = range(int(bounds.group(1)), int(bounds.group(2)) + 1) if bounds else ([int(token)] if token.strip().isdigit() else [])
                weeks.extend(week for week in values if parity == "all" or (week % 2 == 1) == (parity == "odd"))
        day = 7 if raw_day == "天" else int(raw_day) if raw_day.isdigit() else weekdays.index(raw_day) + 1
        meetings.append({
            "day": day, "start": int(raw_start), "end": int(raw_end),
            "weeks": sorted(set(weeks)),
        })
    return tuple(meetings)


def selection_page_count(html: str) -> int:
    """Return the server-rendered result page count, defaulting to one."""
    match = re.search(
        r"<input\b(?=[^>]*\bname\s*=\s*['\"]pageCount['\"])[^>]*"
        r"\bvalue\s*=\s*['\"](\d+)['\"][^>]*>",
        html,
        flags=re.IGNORECASE,
    )
    if not match:
        match = re.search(
            r"<input\b(?=[^>]*\bvalue\s*=\s*['\"](\d+)['\"])[^>]*"
            r"\bname\s*=\s*['\"]pageCount['\"][^>]*>",
            html,
            flags=re.IGNORECASE,
        )
    return max(1, int(match.group(1))) if match else 1


def extract_course_sections_from_html(html: str) -> tuple[SelectionCourseSection, ...]:
    """Parse the legacy server-rendered candidate-course table."""
    parser = _SelectionTableParser()
    parser.feed(html)
    headers: dict[str, int] | None = None
    sections: list[SelectionCourseSection] = []
    seen: set[str] = set()
    for row in parser.rows:
        normalized = [_normalise_header(cell.text) for cell in row]
        if "课程代码" in normalized and "课程名称" in normalized:
            headers = {name: index for index, name in enumerate(normalized)}
            continue
        if headers is None or len(row) < max(headers.values()) + 1:
            continue

        def value(name: str) -> str:
            index = headers.get(name)
            return row[index].text.strip() if index is not None else ""

        course_code = value("课程代码")
        name = value("课程名称")
        if not course_code or not name:
            continue
        class_info = value("上课信息")
        selected_count, capacity_count = _split_capacity(value("已选/容量"))
        attribute_text = " ".join(cell.attributes for cell in row)
        action_match = re.search(
            r"\b(saveXsxk\d?)\s*\(\s*['\"]([^'\"]+)['\"]",
            attribute_text,
            flags=re.IGNORECASE,
        )
        action_name = action_match.group(1) if action_match else ""
        action_rwh = action_match.group(2) if action_match else ""
        identity = action_rwh or "|".join(
            part for part in (course_code, name, class_info) if part
        )
        if identity in seen:
            continue
        seen.add(identity)
        sections.append(
            SelectionCourseSection(
                identity=identity,
                name=name,
                credits=value("学分"),
                teacher=_teacher_from_class_info(class_info),
                time=_schedule_from_class_info(class_info),
                capacity=value("已选/容量"),
                selected=None,
                category=value("课程类别"),
                course_code=course_code,
                campus=value("校区"),
                department=value("开课院系"),
                hours=value("学时"),
                prerequisites=value("前置课程"),
                audience=value("面向对象"),
                notes=value("备注信息"),
                requirements=value("选课要求"),
                selected_count=selected_count,
                capacity_count=capacity_count,
                action_rwh=action_rwh,
                action_name=action_name,
                execution_ready=bool(action_rwh),
                meetings=_schedule_meetings(_schedule_from_class_info(class_info)),
            )
        )
    return tuple(sections)


def _selection_window(html: str) -> tuple[datetime, datetime] | None:
    parser = _PageTextParser()
    parser.feed(html)
    text = " ".join(parser.parts)
    match = re.search(
        r"选课时间\s*[:：]?\s*(\d{4}-\d{2}-\d{2}\s+\d{2}:\d{2})"
        r"\s*至\s*(\d{4}-\d{2}-\d{2}\s+\d{2}:\d{2})",
        text,
    )
    if not match:
        return None
    return tuple(datetime.strptime(value, "%Y-%m-%d %H:%M") for value in match.groups())


def classify_selection_html(
    status_code: int,
    html: str,
    *,
    request_url: str = "",
    now: datetime | None = None,
    expected_windows: tuple[Any, ...] = (),
) -> SelectionObservation:
    """Classify an HTML query without treating a future empty round as no courses."""
    sections = extract_course_sections_from_html(html)
    lowered = html.lower()
    window = _selection_window(html)
    current = now or datetime.now()
    expected_pairs = {
        (getattr(item, "opens_at", ""), getattr(item, "closes_at", ""))
        for item in expected_windows
    }
    observed_pair = (
        window[0].strftime("%Y-%m-%d %H:%M"), window[1].strftime("%Y-%m-%d %H:%M")
    ) if window else None
    if status_code in {401, 403} or "authserver/login" in lowered or "#!/login" in lowered:
        status = STATUS_LOGIN_REQUIRED
    elif expected_pairs and observed_pair and observed_pair not in expected_pairs:
        status = STATUS_NO_MATCHING_ROUND
    elif any(marker in html for marker in ("未开放", "尚未开始", "未到选课时间", "不在选课时间")):
        status = STATUS_ROUND_NOT_OPEN
    elif sections:
        status = STATUS_READY
    elif window and current < window[0]:
        status = STATUS_ROUND_NOT_OPEN
    elif status_code >= 400:
        status = STATUS_ENTRY_UNREACHABLE
    else:
        status = STATUS_EMPTY
    message = ""
    if window:
        message = f"选课窗口：{window[0]:%Y-%m-%d %H:%M} 至 {window[1]:%Y-%m-%d %H:%M}"
    return SelectionObservation(
        status=status,
        request_url=sanitize_url(request_url),
        method="POST",
        message=message,
        sections=sections,
    )


def _payload_text(payload: Any) -> str:
    return json.dumps(payload, ensure_ascii=False).lower()


def classify_selection_response(
    status_code: int, payload: Any, *, request_url: str = ""
) -> SelectionObservation:
    text = _payload_text(payload)
    sections = extract_course_sections(payload)
    if status_code in {401, 403} or any(
        marker in text for marker in ("登录", "未认证", "unauthorized", "forbidden")
    ):
        status = STATUS_LOGIN_REQUIRED
    elif any(marker in text for marker in ("未开放", "尚未开始", "未到选课时间", "不在选课时间")):
        status = STATUS_ROUND_NOT_OPEN
    elif sections:
        status = STATUS_READY
    elif any(marker in text for marker in ("无匹配", "无符合", "暂无相关", "没有符合")):
        status = STATUS_NO_MATCHING_ROUND
    elif status_code >= 400:
        status = STATUS_ENTRY_UNREACHABLE
    else:
        status = STATUS_EMPTY
    return SelectionObservation(
        status=status,
        request_url=request_url,
        method="",
        message="",
        sections=sections,
    )


def _field_paths(value: Any, prefix: str = "") -> set[str]:
    fields: set[str] = set()
    if isinstance(value, dict):
        for key, child in value.items():
            if is_sensitive_key(key):
                continue
            path = f"{prefix}.{key}" if prefix else str(key)
            fields.add(path)
            fields.update(_field_paths(child, path))
    elif isinstance(value, list) and value:
        fields.update(_field_paths(value[0], f"{prefix}[]"))
    return fields


def selection_page_matches_notice(text: str, notice) -> bool:
    normalized = text.lower()
    if "选课" not in normalized:
        return False
    specific_markers = tuple(
        marker.lower()
        for marker in (
            getattr(notice, "selection_type", ""),
            getattr(notice, "term", ""),
        )
        if marker
    )
    return bool(specific_markers) and all(marker in normalized for marker in specific_markers)


def observe_json_exchange(
    *,
    method: str,
    url: str,
    status_code: int,
    request_body: str | None,
    payload: Any,
) -> tuple[SelectionObservation, SelectionInterfaceContract]:
    observation = classify_selection_response(status_code, payload, request_url=url)
    observation = SelectionObservation(
        status=observation.status,
        request_url=sanitize_url(observation.request_url),
        method=method,
        message=observation.message,
        sections=observation.sections,
    )
    contract = SelectionInterfaceContract(
        method=method,
        url=sanitize_url(url),
        status_code=status_code,
        request_body=sanitize_request_body(request_body or ""),
        response_fields=tuple(sorted(_field_paths(sanitize_data(payload)))),
        observed_status=observation.status,
        captured_at=_now(),
    )
    return observation, contract


def observation_to_dict(observation: SelectionObservation) -> dict[str, Any]:
    data = asdict(observation)
    data["sections"] = [asdict(item) for item in observation.sections]
    return data


def contract_to_dict(contract: SelectionInterfaceContract) -> dict[str, Any]:
    return asdict(contract)


def save_selection_result(
    output_root, observation: SelectionObservation, contracts: Iterable[SelectionInterfaceContract] = ()
) -> Path:
    output_root.mkdir(parents=True, exist_ok=True)
    result_path = output_root / "selection-entry.json"
    result_path.write_text(
        json.dumps(observation_to_dict(observation), ensure_ascii=False, indent=2),
        encoding="utf-8",
    )
    (output_root / "selection-contracts.json").write_text(
        json.dumps([contract_to_dict(item) for item in contracts], ensure_ascii=False, indent=2),
        encoding="utf-8",
    )
    return result_path


class SelectionEntryExplorer:
    """Observe a user-opened selection page without clicking its controls."""

    def __init__(self, *, notice, output_root, max_response_bytes: int = 5 * 1024 * 1024):
        self.notice = notice
        self.output_root = output_root
        self.max_response_bytes = max_response_bytes
        self.observations: list[SelectionObservation] = []
        self.contracts: list[SelectionInterfaceContract] = []
        self._selection_pages: list[Any] = []

    def _handle_response(self, response) -> None:
        try:
            page = response.frame.page
        except Exception:
            return
        if not any(page is item for item in self._selection_pages):
            return
        request = response.request
        if request.resource_type not in {"xhr", "fetch"}:
            return
        if response.status >= 500:
            observation, contract = observe_json_exchange(
                method=request.method,
                url=response.url,
                status_code=response.status,
                request_body=request.post_data,
                payload={"message": "entry unreachable"},
            )
            self.observations.append(observation)
            self.contracts.append(contract)
            return
        content_type = response.headers.get("content-type", "")
        if "json" not in content_type.lower():
            return
        try:
            body = response.body()
            if len(body) > self.max_response_bytes:
                return
            payload = json.loads(body.decode("utf-8-sig"))
        except Exception:
            return
        observation, contract = observe_json_exchange(
            method=request.method,
            url=response.url,
            status_code=response.status,
            request_body=request.post_data,
            payload=payload,
        )
        self.observations.append(observation)
        self.contracts.append(contract)

    def _matches_selection_page(self, page) -> bool:
        try:
            text = (page.locator("body").inner_text(timeout=500) or "").lower()
        except Exception:
            return False
        selection_type = getattr(self.notice, "selection_type", "")
        return selection_page_matches_notice(text, self.notice)

    def _has_terminal_selection_observation(self) -> bool:
        terminal = {
            STATUS_READY,
            STATUS_ROUND_NOT_OPEN,
            STATUS_NO_MATCHING_ROUND,
        }
        return any(
            item.status in terminal
            and any(marker in item.request_url.lower() for marker in _SELECTION_URL_MARKERS)
            for item in self.observations
        )

    def _best_observation(self) -> SelectionObservation | None:
        priority = {
            STATUS_READY: 5,
            STATUS_ROUND_NOT_OPEN: 4,
            STATUS_NO_MATCHING_ROUND: 3,
            STATUS_INTERFACE_UNCONFIRMED: 2,
            STATUS_EMPTY: 2,
            STATUS_ENTRY_UNREACHABLE: 1,
            STATUS_LOGIN_REQUIRED: 0,
        }
        return max(
            self.observations,
            key=lambda item: priority.get(item.status, -1),
            default=None,
        )

    def run(self, context, *, wait_seconds: int = 600) -> Path:
        context.on("response", self._handle_response)
        print("请在可见浏览器中完成认证，并手动打开通知对应的学生选课页面。")
        print("安全边界：程序只监听 Fetch/XHR JSON，不点击选课、退课或提交控件。")
        deadline = time.monotonic() + wait_seconds
        matched = False
        while time.monotonic() < deadline:
            pages = tuple(context.pages)
            for page in pages:
                if self._matches_selection_page(page) and not any(
                    page is item for item in self._selection_pages
                ):
                    self._selection_pages.append(page)
            matched = matched or bool(self._selection_pages)
            if matched and self._has_terminal_selection_observation():
                break
            if pages:
                pages[-1].wait_for_timeout(500)
            else:
                time.sleep(0.5)
        if not matched:
            status = STATUS_ENTRY_UNREACHABLE
        elif not self.observations:
            status = STATUS_INTERFACE_UNCONFIRMED
        else:
            status = self.observations[-1].status
        best = self._best_observation()
        observation = SelectionObservation(
            status=best.status if best else status,
            request_url=best.request_url if best else "",
            method=best.method if best else "",
            message="",
            sections=tuple(
                section
                for item in reversed(self.observations)
                for section in item.sections
            ),
        )
        self.output_root.mkdir(parents=True, exist_ok=True)
        return save_selection_result(self.output_root, observation, self.contracts)
