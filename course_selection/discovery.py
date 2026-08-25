"""Safe automatic navigation for discovering timetable and selection APIs."""

from __future__ import annotations

import json
import time
from dataclasses import asdict, dataclass
from datetime import datetime
from pathlib import Path
from typing import Any
from urllib.parse import parse_qs, urljoin, urlsplit

from playwright.sync_api import BrowserContext, Error, Playwright, Response

from course_progress.capture import CaptureStore
from course_progress.collector import resolve_academic_url
from course_progress.sanitizer import sanitize_text, sanitize_url
from course_progress.sanitizer import sanitize_request_body
from course_progress.session import AcademicBrowserSession
from course_selection.categories import CATEGORY_MENU_KEYWORDS, COURSE_CATEGORIES
from course_selection.selection_entry import (
    classify_selection_html,
    observation_to_dict,
    selection_page_count,
)


TARGET_TIMETABLE = "timetable"
TARGET_SELECTION = "selection"

TARGET_KEYWORDS = {
    TARGET_TIMETABLE: (
        "我的课表",
        "个人课表",
        "学生课表",
        "课表查询",
        "课程表",
        "课表",
    ),
    TARGET_SELECTION: (
        *CATEGORY_MENU_KEYWORDS,
        "全校任选课",
        "人文社科限选课",
        "限选课",
        "必修课",
        "学生选课",
        "选课中心",
        "课程选课",
        "网上选课",
        "选课",
    ),
}

_GENERIC_SELECTION_PRIORITIES = (
    ("全校任选课", 240),
    ("人文社科限选课", 235),
    ("限选课", 220),
    ("必修课", 215),
    ("学生选课", 180),
    ("选课中心", 175),
)

TARGET_FRAME_MARKERS = {
    TARGET_TIMETABLE: ("/kbcx/querygrkb",),
    TARGET_SELECTION: ("/xsxk/queryxsxk",),
}

INTERMEDIATE_KEYWORDS = (
    "统一身份认证登录",
    "本科生综合服务",
    "综合教务",
    "教务系统",
    "教务",
    "本科生",
    "学生服务",
    "academic",
    "student",
    "jwgl",
)

SAFE_QUERY_KEYWORDS = ("查询", "搜索", "刷新", "本学期", "当前学期")

CONTROL_BLOCKLIST = (
    "确认选课",
    "提交选课",
    "立即选课",
    "退课",
    "退选",
    "撤销",
    "保存",
    "提交",
    "删除",
    "报名",
    "缴费",
    "退出登录",
    "注销",
)

MUTATION_MARKERS = (
    "submit",
    "save",
    "delete",
    "remove",
    "withdraw",
    "dropcourse",
    "drop-course",
    "registercourse",
    "register-course",
    "selectcourse",
    "select-course",
    "choosecourse",
    "choose-course",
    "enroll",
    "退课",
    "退选",
    "提交",
    "保存",
)

CONTROL_SELECTOR = (
    "a, button, [role='button'], [role='link'], [role='menuitem'], "
    ".el-menu-item, .ant-menu-item, .layui-nav-item a"
)

_FETCH_SELECTION_PAGE = """
async ({url, category, semesterLabel, pageNo, pageCount}) => {
  const form = document.querySelector('form#queryform, form[name="queryform"]');
  if (!form) throw new Error('未找到选课查询表单');
  const parameters = new URLSearchParams(new FormData(form));
  parameters.set('rwh', '');
  parameters.set('pageXklb', category);
  if (pageNo !== null) {
    parameters.set('pageNo', String(pageNo));
    parameters.set('pageSize', '20');
    parameters.set('pageCount', String(pageCount));
  } else {
    parameters.delete('pageNo');
    parameters.delete('pageSize');
    parameters.delete('pageCount');
  }
  const semester = Array.from(form.querySelectorAll('select[name="pageXnxq"] option'))
    .find(option => option.textContent.replace(/[\s年]|学期/g, '') === semesterLabel);
  if (!semester) throw new Error(`通知学期不在页面选项中: ${semesterLabel}`);
  parameters.set('pageXnxq', semester.value);
  const response = await fetch(url, {
    method: 'POST',
    credentials: 'same-origin',
    headers: {'Content-Type': 'application/x-www-form-urlencoded'},
    body: parameters.toString(),
    redirect: 'follow',
  });
  return {
    status: response.status,
    url: response.url,
    requestBody: parameters.toString(),
    body: await response.text(),
  };
}
"""


@dataclass(frozen=True)
class DiscoveryControl:
    score: int
    identity: str
    text: str
    frame_url: str
    locator: Any
    metadata: dict[str, Any]


@dataclass(frozen=True)
class DiscoveryReport:
    target: str
    target_found: bool
    clicks: int
    captures: int
    blocked_requests: int
    candidates_path: Path
    click_log_path: Path
    selection_query_path: Path | None


def score_discovery_control(
    target: str,
    *,
    text: str,
    href: str = "",
    target_page_reached: bool = False,
) -> int:
    """Rank navigation and read-only query controls; reject mutation controls."""
    if target not in TARGET_KEYWORDS:
        raise ValueError(f"未知发现目标：{target}")
    normalized_text = "".join(text.split()).lower()
    searchable = f"{normalized_text} {href}".strip().lower()
    if not searchable or any(marker.lower() in searchable for marker in CONTROL_BLOCKLIST):
        return -1
    if target_page_reached:
        if target == TARGET_SELECTION:
            # Selection queries are issued once per notice-approved category by
            # the direct read-only fetcher after navigation completes.
            return -1
        if normalized_text in {keyword.lower() for keyword in SAFE_QUERY_KEYWORDS}:
            return 60
        return -1
    if target == TARGET_SELECTION:
        priorities = tuple(
            (alias, definition.navigation_priority)
            for definition in COURSE_CATEGORIES
            for alias in definition.menu_aliases
        ) + _GENERIC_SELECTION_PRIORITIES
        for keyword, score in priorities:
            if keyword.lower() in searchable:
                return score
    if target == TARGET_TIMETABLE:
        if any(
            keyword in searchable
            for keyword in ("我的课表", "个人课表", "学生课表", "课表查询")
        ):
            return 240
        if "学生选课" in searchable:
            # The legacy academic system exposes timetable queries below this menu.
            return 180
    target_hits = [keyword for keyword in TARGET_KEYWORDS[target] if keyword.lower() in searchable]
    if target_hits:
        return 100 + max(len(keyword) for keyword in target_hits)
    if any(keyword.lower() in searchable for keyword in INTERMEDIATE_KEYWORDS):
        return 30
    return -1


def is_mutating_request(method: str, url: str, post_data: str | None = None) -> bool:
    """Conservatively identify requests that may change course-selection state."""
    if method.upper() in {"GET", "HEAD", "OPTIONS"}:
        return False
    searchable = f"{url} {post_data or ''}".lower()
    return any(marker.lower() in searchable for marker in MUTATION_MARKERS)


def is_control_allowed_in_stage(
    target: str,
    *,
    text: str,
    href: str,
    selection_menu_expanded: bool,
    allowed_selection_categories: tuple[str, ...] = (),
) -> bool:
    """Keep top-level menus distinct from similarly named selection categories."""
    if target != TARGET_SELECTION:
        return True
    normalized = "".join(text.split())
    if not selection_menu_expanded:
        return normalized in {"统一身份认证登录", "学生选课"}
    if normalized in {"统一身份认证登录", "学生选课"}:
        return False
    category = parse_qs(urlsplit(href).query).get("pageXklb", [""])[0]
    return (
        "/xsxk/queryxsxk" in href.lower()
        and category in allowed_selection_categories
    )


def _same_origin(left: str, right: str) -> bool:
    first = urlsplit(left)
    second = urlsplit(right)
    return bool(first.netloc and first.netloc == second.netloc)


def _walk_dicts(value: Any):
    if isinstance(value, dict):
        yield value
        for child in value.values():
            yield from _walk_dicts(child)
    elif isinstance(value, list):
        for child in value:
            yield from _walk_dicts(child)


def find_academic_portal_redirect(payload: Any) -> str:
    """Find the new academic-system application route in a portal catalog."""
    best: tuple[int, str] | None = None
    for item in _walk_dicts(payload):
        name = str(item.get("name", ""))
        redirect = str(item.get("redirect", ""))
        if not redirect:
            continue
        if "新教务系统" in name:
            score = 100
        elif "教务系统" in name:
            score = 80
        elif "教务" in name:
            score = 40
        else:
            continue
        if best is None or score > best[0]:
            best = (score, redirect)
    return best[1] if best else ""


class InterfaceDiscovery:
    """Analyze visible controls, auto-navigate, and capture read-only JSON exchanges."""

    def __init__(
        self,
        *,
        target: str,
        output_root: Path,
        allowed_selection_categories: tuple[str, ...] = (),
        allowed_selection_windows: dict[str, tuple[Any, ...]] | None = None,
        notice_semester: str = "",
        max_response_bytes: int = 5 * 1024 * 1024,
    ):
        if target not in TARGET_KEYWORDS:
            raise ValueError(f"未知发现目标：{target}")
        stamp = datetime.now().strftime("%Y%m%d-%H%M%S")
        self.target = target
        self.output_root = output_root / "discovery" / target / stamp
        self.store = CaptureStore(self.output_root)
        self.max_response_bytes = max_response_bytes
        self.allowed_selection_categories = allowed_selection_categories
        self.allowed_selection_windows = allowed_selection_windows or {}
        self.notice_semester = notice_semester
        self.visited: set[str] = set()
        self.click_log: list[dict[str, str | int]] = []
        self.captured = 0
        self.target_pages: list[Any] = []
        self.portal_redirects: list[str] = []
        self.blocked_requests = 0
        self.dom_snapshots: list[dict[str, Any]] = []
        self.visual_log: list[dict[str, Any]] = []
        self.selection_menu_expanded = False

    def _guard_route(self, route) -> None:
        request = route.request
        if is_mutating_request(request.method, request.url, request.post_data):
            self.blocked_requests += 1
            print(f"安全拦截：{request.method} {request.url[:110]}")
            route.abort("blockedbyclient")
            return
        route.continue_()

    def _handle_response(self, response: Response) -> None:
        request = response.request
        if request.resource_type not in {"xhr", "fetch"}:
            return
        content_type = response.headers.get("content-type", "").lower()
        if "json" not in content_type:
            return
        try:
            body = response.body()
            if len(body) > self.max_response_bytes:
                return
            payload = json.loads(body.decode("utf-8-sig"))
        except (Error, UnicodeDecodeError, json.JSONDecodeError):
            return
        redirect = find_academic_portal_redirect(payload)
        if redirect and redirect not in self.portal_redirects:
            self.portal_redirects.append(redirect)
            print(f"门户目录：发现教务系统入口 {redirect[:100]}")
        candidate = self.store.save_json_exchange(
            url=response.url,
            method=request.method,
            status=response.status,
            content_type=content_type,
            request_body=request.post_data,
            response_data=payload,
        )
        self.captured += 1
        print(f"接口记录：{candidate.score:02d} {request.method} {candidate.url[:100]}")

    def _page_matches_target(self, page) -> bool:
        return any(self._frame_matches_target(frame) for frame in page.frames)

    def _frame_matches_target(self, frame) -> bool:
        markers = TARGET_FRAME_MARKERS[self.target]
        if markers:
            marker_matches = any(marker in frame.url.lower() for marker in markers)
            if not marker_matches:
                return False
            if self.target != TARGET_SELECTION:
                return True
            category = parse_qs(urlsplit(frame.url).query).get("pageXklb", [""])[0]
            if not category:
                try:
                    category = frame.locator("input[name='pageXklb']").get_attribute(
                        "value", timeout=500
                    ) or ""
                except Error:
                    return False
            return category in self.allowed_selection_categories
        try:
            body = (frame.locator("body").inner_text(timeout=700) or "").lower()
        except Error:
            return False
        if self.target == TARGET_SELECTION:
            return ("选择课程" in body or "备选课程" in body) and "已选课程" in body
        return False

    def _capture_click_visuals(
        self,
        context: BrowserContext,
        *,
        click_number: int,
        stage: str,
        text: str,
    ) -> None:
        files: list[str] = []
        for page_index, page in enumerate(context.pages, start=1):
            try:
                filename = f"click-{click_number:02d}-{stage}-page-{page_index}.png"
                page.screenshot(path=str(self.output_root / filename), full_page=True)
                files.append(filename)
            except Error:
                continue
        self.visual_log.append(
            {
                "click": click_number,
                "stage": stage,
                "text": sanitize_text(text),
                "files": files,
            }
        )

    def _refresh_target_pages(self, context: BrowserContext) -> None:
        for page in context.pages:
            if self._page_matches_target(page) and not any(
                page is known for known in self.target_pages
            ):
                self.target_pages.append(page)
                print(f"目标页面：{page.title()} | {page.url[:110]}")

    def _controls(
        self, context: BrowserContext
    ) -> tuple[list[DiscoveryControl], list[dict[str, Any]]]:
        found: list[DiscoveryControl] = []
        inventory: list[dict[str, Any]] = []
        target_reached = bool(self.target_pages)
        for page in context.pages:
            # Before reaching the destination, only inspect the academic shell's
            # menu tree. Content iframes contain tabs such as “备选课程”, which are
            # not navigation entries and previously caused false clicks.
            frames = page.frames if target_reached else (page.main_frame,)
            for frame in frames:
                if target_reached and not self._frame_matches_target(frame):
                    continue
                controls = frame.locator(CONTROL_SELECTOR)
                try:
                    count = min(controls.count(), 200)
                except Error:
                    continue
                for index in range(count):
                    locator = controls.nth(index)
                    try:
                        if not locator.is_visible():
                            continue
                        text = " ".join((locator.inner_text() or "").split())[:120]
                        if not text:
                            text = (locator.get_attribute("aria-label") or locator.get_attribute("title") or "")[:120]
                        href = locator.get_attribute("href") or ""
                        metadata = locator.evaluate(
                            """(element, index) => {
                              const rect = element.getBoundingClientRect();
                              const owner = element.closest(
                                "li, [role='menuitem'], .menu, .submenu, .nav-item"
                              );
                              return {
                                index,
                                tag: element.tagName.toLowerCase(),
                                role: element.getAttribute('role') || '',
                                href: element.getAttribute('href') || '',
                                onclick: element.getAttribute('onclick') || '',
                                ownerText: owner && owner !== element
                                  ? (owner.innerText || '').replace(/\\s+/g, ' ').trim().slice(0, 160)
                                  : '',
                                box: {
                                  x: Math.round(rect.x), y: Math.round(rect.y),
                                  width: Math.round(rect.width), height: Math.round(rect.height)
                                },
                                inViewport: rect.bottom > 0 && rect.right > 0
                                  && rect.top < window.innerHeight
                                  && rect.left < window.innerWidth
                              };
                            }""",
                            index,
                        )
                        score = score_discovery_control(
                            self.target,
                            text=text,
                            href=href,
                            target_page_reached=target_reached,
                        )
                        identity = f"{frame.url}|{metadata['tag']}|{text}|{href}|{index}"
                        safe_metadata = {
                            **metadata,
                            "href": sanitize_url(href) if href.startswith(("http://", "https://")) else sanitize_text(href),
                            "onclick": sanitize_text(metadata["onclick"])[:240],
                        }
                        inventory.append(
                            {
                                "identity": sanitize_text(identity),
                                "text": sanitize_text(text),
                                "frame_url": sanitize_url(frame.url),
                                "score": score,
                                **safe_metadata,
                            }
                        )
                        if score < 0:
                            continue
                        if not safe_metadata["inViewport"]:
                            continue
                        if not target_reached and not is_control_allowed_in_stage(
                            self.target,
                            text=text,
                            href=href,
                            selection_menu_expanded=self.selection_menu_expanded,
                            allowed_selection_categories=self.allowed_selection_categories,
                        ):
                            continue
                        if identity in self.visited:
                            continue
                        if href.startswith("mailto:"):
                            continue
                        if href.startswith(("http://", "https://")) and not _same_origin(frame.url, href):
                            continue
                        found.append(
                            DiscoveryControl(
                                score, identity, text, frame.url, locator, safe_metadata
                            )
                        )
                    except Error:
                        continue
        return sorted(found, key=lambda item: (-item.score, item.identity)), inventory

    def _record_dom_snapshot(
        self,
        stage: str,
        inventory: list[dict[str, Any]],
        *,
        newly_visible: list[str] | None = None,
    ) -> None:
        self.dom_snapshots.append(
            {
                "sequence": len(self.dom_snapshots) + 1,
                "stage": stage,
                "target_reached": bool(self.target_pages),
                "newly_visible": newly_visible or [],
                "controls": inventory,
            }
        )

    def _click(self, control: DiscoveryControl) -> bool:
        self.visited.add(control.identity)
        print(f"自动点击 [{control.score}]：{control.text or '未命名入口'}")
        try:
            control.locator.click(timeout=10_000)
            if (
                self.target == TARGET_SELECTION
                and "".join(control.text.split()) == "学生选课"
            ):
                self.selection_menu_expanded = True
            self.click_log.append(
                {
                    "score": control.score,
                    "text": control.text,
                    "frame_url": control.frame_url,
                    "result": "clicked",
                }
            )
            return True
        except Error as error:
            self.click_log.append(
                {
                    "score": control.score,
                    "text": control.text,
                    "frame_url": control.frame_url,
                    "result": f"failed: {str(error)[:180]}",
                }
            )
            return False

    def _query_selection_page(self) -> Path | None:
        if (
            self.target != TARGET_SELECTION
            or not self.allowed_selection_categories
            or not self.notice_semester
        ):
            return None
        for page in self.target_pages:
            for frame in page.frames:
                if not self._frame_matches_target(frame):
                    continue
                endpoint = resolve_academic_url(frame.url, "/xsxk/queryXsxkList")
                queries: list[dict[str, Any]] = []
                for category in self.allowed_selection_categories:
                    pages: list[dict[str, Any]] = []
                    sections: dict[str, Any] = {}
                    expected_pages = 1
                    complete = True
                    try:
                        for page_number in range(1, 51):
                            if page_number > expected_pages:
                                break
                            result = frame.evaluate(
                                _FETCH_SELECTION_PAGE,
                                {
                                    "url": endpoint,
                                    "category": category,
                                    "semesterLabel": self.notice_semester,
                                    # The initial search omits pagination fields. The
                                    # legacy page form sends all three fields only
                                    # when following a pagination link.
                                    "pageNo": None if page_number == 1 else page_number,
                                    "pageCount": expected_pages,
                                },
                            )
                            html = str(result["body"])
                            if page_number == 1:
                                expected_pages = min(selection_page_count(html), 50)
                            observation = classify_selection_html(
                                int(result["status"]),
                                html,
                                request_url=str(result.get("url") or endpoint),
                                expected_windows=self.allowed_selection_windows.get(category, ()),
                            )
                            for section in observation.sections:
                                sections.setdefault(section.identity, section)
                            pages.append(
                                {
                                    "page": page_number,
                                    "observation": observation_to_dict(observation),
                                }
                            )
                        queries.append(
                            {
                                "category": category,
                                "semester": self.notice_semester,
                                "method": "POST",
                                "url": sanitize_url(str(result.get("url") or endpoint)),
                                "request_body": sanitize_request_body(
                                    str(result.get("requestBody") or "")
                                ),
                                "page_count": expected_pages,
                                "pages_fetched": len(pages),
                                "complete": complete and len(pages) == expected_pages,
                                "record_count": len(sections),
                                "sections": [asdict(section) for section in sections.values()],
                                "pages": pages,
                            }
                        )
                        print(
                            f"课程查询：{category} 共 {len(pages)}/{expected_pages} 页，"
                            f"解析 {len(sections)} 条课程记录"
                        )
                    except (Error, KeyError, TypeError, ValueError) as error:
                        complete = False
                        queries.append(
                            {
                                "category": category,
                                "semester": self.notice_semester,
                                "complete": False,
                                "page_count": expected_pages,
                                "pages_fetched": len(pages),
                                "record_count": len(sections),
                                "sections": [asdict(section) for section in sections.values()],
                                "pages": pages,
                                "error": str(error)[:160],
                            }
                        )
                        print(f"课程查询失败 [{category}]：{str(error)[:160]}")
                query_path = self.output_root / "selection-query.json"
                query_path.write_text(
                    json.dumps({"queries": queries}, ensure_ascii=False, indent=2),
                    encoding="utf-8",
                )
                return query_path
        return None

    def run(
        self,
        context: BrowserContext,
        *,
        max_clicks: int = 8,
        wait_seconds: int = 30,
    ) -> DiscoveryReport:
        deadline = time.monotonic() + wait_seconds
        clicks = 0
        idle_rounds = 0
        while time.monotonic() < deadline and clicks < max_clicks:
            self._refresh_target_pages(context)
            if not self.target_pages and self.portal_redirects:
                redirect = self.portal_redirects.pop(0)
                pages = context.pages
                if pages:
                    target_url = urljoin(pages[-1].url, redirect)
                    identity = f"portal-catalog|{target_url}"
                    if identity not in self.visited:
                        self.visited.add(identity)
                        print(f"自动导航：门户目录 -> {target_url[:110]}")
                        self.click_log.append(
                            {
                                "score": 120,
                                "text": "门户目录：新教务系统",
                                "frame_url": pages[-1].url,
                                "result": "navigated",
                            }
                        )
                        pages[-1].goto(
                            target_url,
                            wait_until="domcontentloaded",
                            timeout=60_000,
                        )
                        pages[-1].wait_for_timeout(1_500)
                        clicks += 1
                        continue
            controls, inventory = self._controls(context)
            self._record_dom_snapshot("before-click", inventory)
            if not controls:
                idle_rounds += 1
                if self.target_pages and idle_rounds >= 3:
                    break
                pages = context.pages
                (pages[-1].wait_for_timeout(500) if pages else time.sleep(0.5))
                continue
            idle_rounds = 0
            selected = controls[0]
            before_identities = {item["identity"] for item in inventory}
            self.output_root.mkdir(parents=True, exist_ok=True)
            self._capture_click_visuals(
                context,
                click_number=clicks + 1,
                stage="before",
                text=selected.text,
            )
            if self._click(selected):
                clicks += 1
            pages = context.pages
            (pages[-1].wait_for_timeout(1200) if pages else time.sleep(1.2))
            self._refresh_target_pages(context)
            _, after_inventory = self._controls(context)
            newly_visible = [
                item["text"]
                for item in after_inventory
                if item["identity"] not in before_identities
            ]
            self._record_dom_snapshot(
                f"after-click: {sanitize_text(selected.text)}",
                after_inventory,
                newly_visible=newly_visible,
            )
            self._capture_click_visuals(
                context,
                click_number=clicks,
                stage="after",
                text=selected.text,
            )
            if newly_visible:
                print(f"菜单展开：新增 {len(newly_visible)} 个可见控件")

        self._refresh_target_pages(context)
        self.output_root.mkdir(parents=True, exist_ok=True)
        target_contracts: list[dict[str, Any]] = []
        for page in self.target_pages:
            for frame in page.frames:
                if not self._frame_matches_target(frame):
                    continue
                try:
                    html_path = self.output_root / "target-frame.html"
                    html_path.write_text(frame.content(), encoding="utf-8")
                    forms = frame.locator("form").evaluate_all(
                        """forms => forms.map(form => ({
                          method: (form.method || 'GET').toUpperCase(),
                          action: form.action,
                          fields: Array.from(form.elements).map(field => ({
                            tag: field.tagName.toLowerCase(),
                            type: field.type || '',
                            name: field.name || ''
                          })),
                          selects: Array.from(form.querySelectorAll('select')).map(select => ({
                            name: select.name || '',
                            options: Array.from(select.options).map(option => ({
                              text: option.textContent.trim(), value: option.value
                            }))
                          }))
                        }))"""
                    )
                    for form in forms:
                        form["action"] = sanitize_url(form.get("action", ""))
                    target_contracts.append(
                        {
                            "url": sanitize_url(frame.url),
                            "html": html_path.name,
                            "forms": forms,
                        }
                    )
                except Error as error:
                    target_contracts.append({"error": str(error)[:180]})
        (self.output_root / "target-contract.json").write_text(
            json.dumps(target_contracts, ensure_ascii=False, indent=2),
            encoding="utf-8",
        )
        page_diagnostics: list[dict[str, Any]] = []
        for index, page in enumerate(context.pages):
            try:
                screenshot_path = self.output_root / f"page-{index + 1}.png"
                page.screenshot(path=str(screenshot_path), full_page=True)
                page_diagnostics.append(
                    {
                        "title": page.title(),
                        "url": sanitize_url(page.url),
                        "frames": [sanitize_url(frame.url) for frame in page.frames],
                        "screenshot": screenshot_path.name,
                    }
                )
            except Error as error:
                page_diagnostics.append({"error": str(error)[:180]})
        (self.output_root / "pages.json").write_text(
            json.dumps(page_diagnostics, ensure_ascii=False, indent=2),
            encoding="utf-8",
        )
        (self.output_root / "clicks.json").write_text(
            json.dumps(self.click_log, ensure_ascii=False, indent=2), encoding="utf-8"
        )
        (self.output_root / "dom-snapshots.json").write_text(
            json.dumps(self.dom_snapshots, ensure_ascii=False, indent=2),
            encoding="utf-8",
        )
        (self.output_root / "visuals.json").write_text(
            json.dumps(self.visual_log, ensure_ascii=False, indent=2),
            encoding="utf-8",
        )
        selection_query_path = self._query_selection_page()
        candidates = self.store.write_candidates()
        print(f"自动点击：{clicks}；捕获 JSON：{self.captured}")
        print(f"点击审计：{self.output_root / 'clicks.json'}")
        print(f"接口候选：{candidates}")
        return DiscoveryReport(
            target=self.target,
            target_found=bool(self.target_pages),
            clicks=clicks,
            captures=self.captured,
            blocked_requests=self.blocked_requests,
            candidates_path=candidates,
            click_log_path=self.output_root / "clicks.json",
            selection_query_path=selection_query_path,
        )


class AcademicInterfaceDiscovery:
    """Deep module: discover one academic interface from an authenticated portal."""

    def __init__(
        self,
        playwright: Playwright,
        *,
        browser_name: str,
        profile_root: Path,
        output_root: Path,
        persistent_session: bool = False,
    ):
        self.playwright = playwright
        self.browser_name = browser_name
        self.profile_root = profile_root
        self.output_root = output_root
        self.persistent_session = persistent_session

    def discover(
        self,
        target: str,
        *,
        portal_url: str,
        login_timeout_seconds: int = 600,
        wait_seconds: int = 30,
        max_clicks: int = 8,
        allowed_selection_categories: tuple[str, ...] = (),
        allowed_selection_windows: dict[str, tuple[Any, ...]] | None = None,
        notice_semester: str = "",
    ) -> DiscoveryReport:
        navigator = InterfaceDiscovery(
            target=target,
            output_root=self.output_root,
            allowed_selection_categories=allowed_selection_categories,
            allowed_selection_windows=allowed_selection_windows,
            notice_semester=notice_semester,
        )
        with AcademicBrowserSession(
            self.playwright,
            browser_name=self.browser_name,
            profile_root=self.profile_root,
            persistent=self.persistent_session,
        ) as session:
            if session.context is None:
                raise RuntimeError("浏览器会话未初始化")
            session.context.on("response", navigator._handle_response)
            session.open_authenticated(
                portal_url, timeout_seconds=login_timeout_seconds
            )
            session.context.route("**/*", navigator._guard_route)
            return navigator.run(
                session.context,
                max_clicks=max_clicks,
                wait_seconds=wait_seconds,
            )
