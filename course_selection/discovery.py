"""Safe automatic navigation for discovering timetable and selection APIs."""

from __future__ import annotations

import json
import time
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path
from typing import Any
from urllib.parse import urljoin, urlsplit

from playwright.sync_api import BrowserContext, Error, Playwright, Response

from course_progress.capture import CaptureStore
from course_progress.session import AcademicBrowserSession


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
        "学生选课",
        "选课中心",
        "课程选课",
        "网上选课",
        "选课",
    ),
}

INTERMEDIATE_KEYWORDS = (
    "本科生综合服务",
    "综合教务",
    "教务系统",
    "教务",
    "本科生",
    "学生服务",
    "学生",
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


@dataclass(frozen=True)
class DiscoveryControl:
    score: int
    identity: str
    text: str
    frame_url: str
    locator: Any


@dataclass(frozen=True)
class DiscoveryReport:
    target: str
    target_found: bool
    clicks: int
    captures: int
    blocked_requests: int
    candidates_path: Path
    click_log_path: Path


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
    searchable = f"{text} {href}".strip().lower()
    if not searchable or any(marker.lower() in searchable for marker in CONTROL_BLOCKLIST):
        return -1
    target_hits = [
        keyword for keyword in TARGET_KEYWORDS[target] if keyword.lower() in searchable
    ]
    if target_hits:
        return 100 + max(len(keyword) for keyword in target_hits)
    if target_page_reached and any(
        keyword.lower() in searchable for keyword in SAFE_QUERY_KEYWORDS
    ):
        return 60
    if any(keyword.lower() in searchable for keyword in INTERMEDIATE_KEYWORDS):
        return 30
    return -1


def is_mutating_request(method: str, url: str, post_data: str | None = None) -> bool:
    """Conservatively identify requests that may change course-selection state."""
    if method.upper() in {"GET", "HEAD", "OPTIONS"}:
        return False
    searchable = f"{url} {post_data or ''}".lower()
    return any(marker.lower() in searchable for marker in MUTATION_MARKERS)


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
        max_response_bytes: int = 5 * 1024 * 1024,
    ):
        if target not in TARGET_KEYWORDS:
            raise ValueError(f"未知发现目标：{target}")
        stamp = datetime.now().strftime("%Y%m%d-%H%M%S")
        self.target = target
        self.output_root = output_root / "discovery" / target / stamp
        self.store = CaptureStore(self.output_root)
        self.max_response_bytes = max_response_bytes
        self.visited: set[str] = set()
        self.click_log: list[dict[str, str | int]] = []
        self.captured = 0
        self.target_pages: list[Any] = []
        self.portal_redirects: list[str] = []
        self.blocked_requests = 0

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
        try:
            body = (page.locator("body").inner_text(timeout=700) or "").lower()
        except Error:
            return False
        return any(keyword.lower() in body for keyword in TARGET_KEYWORDS[self.target])

    def _refresh_target_pages(self, context: BrowserContext) -> None:
        for page in context.pages:
            if self._page_matches_target(page) and not any(
                page is known for known in self.target_pages
            ):
                self.target_pages.append(page)
                print(f"目标页面：{page.title()} | {page.url[:110]}")

    def _controls(self, context: BrowserContext) -> list[DiscoveryControl]:
        found: list[DiscoveryControl] = []
        target_reached = bool(self.target_pages)
        for page in context.pages:
            for frame in page.frames:
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
                        score = score_discovery_control(
                            self.target,
                            text=text,
                            href=href,
                            target_page_reached=target_reached,
                        )
                        if score < 0:
                            continue
                        identity = f"{frame.url}|{text}|{href}"
                        if identity in self.visited:
                            continue
                        if href.startswith(("javascript:", "mailto:")):
                            continue
                        if href.startswith(("http://", "https://")) and not _same_origin(frame.url, href):
                            continue
                        found.append(
                            DiscoveryControl(score, identity, text, frame.url, locator)
                        )
                    except Error:
                        continue
        return sorted(found, key=lambda item: (-item.score, item.identity))

    def _click(self, control: DiscoveryControl) -> bool:
        self.visited.add(control.identity)
        print(f"自动点击 [{control.score}]：{control.text or '未命名入口'}")
        try:
            control.locator.click(timeout=10_000)
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
                        pages[-1].goto(
                            target_url,
                            wait_until="domcontentloaded",
                            timeout=60_000,
                        )
                        pages[-1].wait_for_timeout(1_500)
                        clicks += 1
                        continue
            controls = self._controls(context)
            if not controls:
                idle_rounds += 1
                if self.target_pages and idle_rounds >= 3:
                    break
                pages = context.pages
                (pages[-1].wait_for_timeout(500) if pages else time.sleep(0.5))
                continue
            idle_rounds = 0
            if self._click(controls[0]):
                clicks += 1
            pages = context.pages
            (pages[-1].wait_for_timeout(1200) if pages else time.sleep(1.2))

        self._refresh_target_pages(context)
        self.output_root.mkdir(parents=True, exist_ok=True)
        (self.output_root / "clicks.json").write_text(
            json.dumps(self.click_log, ensure_ascii=False, indent=2), encoding="utf-8"
        )
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
    ):
        self.playwright = playwright
        self.browser_name = browser_name
        self.profile_root = profile_root
        self.output_root = output_root

    def discover(
        self,
        target: str,
        *,
        portal_url: str,
        login_timeout_seconds: int = 600,
        wait_seconds: int = 30,
        max_clicks: int = 8,
    ) -> DiscoveryReport:
        navigator = InterfaceDiscovery(target=target, output_root=self.output_root)
        with AcademicBrowserSession(
            self.playwright,
            browser_name=self.browser_name,
            profile_root=self.profile_root,
        ) as session:
            if session.context is None:
                raise RuntimeError("浏览器会话未初始化")
            session.context.route("**/*", navigator._guard_route)
            session.context.on("response", navigator._handle_response)
            session.open_authenticated(
                portal_url, timeout_seconds=login_timeout_seconds
            )
            return navigator.run(
                session.context,
                max_clicks=max_clicks,
                wait_seconds=wait_seconds,
            )
