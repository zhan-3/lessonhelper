"""Read-only observation of a confirmed academic selection entry."""

from __future__ import annotations

import json
import time
from dataclasses import asdict, dataclass
from datetime import datetime, timezone
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
