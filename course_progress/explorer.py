"""Headed Playwright explorer for discovering HITWH course progress APIs."""

from __future__ import annotations

import json
import os
import time
from datetime import datetime, timezone
from pathlib import Path
from typing import Any
from urllib.parse import urljoin, urlsplit

from playwright.sync_api import BrowserContext, Error, Playwright, Response

from .capture import CaptureStore

DEFAULT_PORTAL_URL = "https://webvpn.hitwh.edu.cn/"

SHARED_PROFILE_NAME = "playwright-chromium-profile"
PROFILE_LOCK_MARKERS = (
    "processsingleton",
    "user data directory is already in use",
    "profile is in use",
)


def resolve_profile_dir(private_root: Path) -> Path:
    """Use one dedicated profile compatible with bundled Playwright Chromium."""
    return private_root / SHARED_PROFILE_NAME


def _is_profile_lock_error(error: Error) -> bool:
    message = str(error).lower()
    return any(marker in message for marker in PROFILE_LOCK_MARKERS)


def _is_login_url(url: str) -> bool:
    lowered = url.lower()
    return (
        "authserver/login" in lowered
        or "logincas" in lowered
        or "loginnocas" in lowered
        or "#!/login" in lowered
    )


AUTO_NAV_KEYWORDS = (
    "教务",
    "综合教务",
    "信息门户",
    "学生事务",
    "本科生",
    "学籍",
    "培养方案",
    "教学计划",
    "学业完成",
    "毕业审核",
    "毕业要求",
    "已修课程",
    "课程完成",
    "课程信息",
    "课程",
    "学分",
    "成绩",
    "training",
    "curriculum",
    "graduat",
    "credit",
    "course",
)

AUTO_NAV_BLOCKLIST = (
    "选课",
    "退课",
    "退选",
    "抢课",
    "提交",
    "保存",
    "删除",
    "修改",
    "缴费",
    "报名",
    "确认",
    "新增",
    "导入",
    "退出",
    "注销",
    "修改密码",
    "个人设置",
)


def _is_safe_navigation(current_url: str, target_url: str) -> bool:
    """Allow only same-origin GET navigation during automatic exploration."""
    current = urlsplit(current_url)
    target = urlsplit(target_url)
    return (
        target.scheme in {"http", "https"}
        and target.netloc == current.netloc
        and not target_url.lower().startswith(("javascript:", "mailto:"))
    )


def _is_relevant_control(text: str, href: str | None = None) -> bool:
    searchable = f"{text} {href or ''}".lower()
    return (
        any(keyword.lower() in searchable for keyword in AUTO_NAV_KEYWORDS)
        and not any(keyword.lower() in searchable for keyword in AUTO_NAV_BLOCKLIST)
    )


def _is_safe_portal_fallback(text: str, href: str | None) -> bool:
    """Allow likely application tiles when the portal uses generic labels."""
    if not href or not text:
        return False
    searchable = f"{text} {href}".lower()
    if any(keyword.lower() in searchable for keyword in AUTO_NAV_BLOCKLIST):
        return False
    if any(keyword.lower() in searchable for keyword in AUTO_NAV_KEYWORDS):
        return True
    return any(
        keyword in searchable
        for keyword in (
            "student",
            "academic",
            "教务",
            "本科",
            "学生",
            "jwc",
            "jwgl",
            "edu",
        )
    )


def launch_browser_context(
    playwright: Playwright, browser_name: str, profile_dir: Path
) -> BrowserContext:
    profile_dir.mkdir(parents=True, exist_ok=True)
    browser_args = ["--start-maximized"]
    debug_port = os.environ.get("ACADEMIC_BROWSER_DEBUG_PORT", "").strip()
    if debug_port:
        if not debug_port.isdigit() or not 1024 <= int(debug_port) <= 65535:
            raise ValueError("ACADEMIC_BROWSER_DEBUG_PORT must be a port from 1024 to 65535")
        browser_args.extend([
            "--remote-debugging-address=127.0.0.1",
            f"--remote-debugging-port={debug_port}",
        ])
    options: dict[str, Any] = {
        "user_data_dir": str(profile_dir),
        "headless": False,
        "no_viewport": True,
        "args": browser_args,
    }
    if browser_name == "chrome":
        options["channel"] = "chrome"

    try:
        return playwright.chromium.launch_persistent_context(**options)
    except Error as exc:
        if _is_profile_lock_error(exc):
            raise RuntimeError(
                f"浏览器 profile 正被其他进程使用：{profile_dir}。"
                "请关闭占用它的浏览器或采集命令后重试。"
            ) from exc
        if "Executable doesn't exist" in str(exc):
            raise RuntimeError(
                "Playwright Chromium 尚未安装。请运行："
                "uv run playwright install chromium"
            ) from exc
        raise


class PortalExplorer:
    def __init__(
        self,
        *,
        playwright: Playwright,
        browser_name: str,
        profile_dir: Path,
        capture_dir: Path,
        max_response_bytes: int,
    ):
        self.playwright = playwright
        self.browser_name = browser_name
        self.profile_dir = profile_dir
        self.capture_dir = capture_dir
        self.max_response_bytes = max_response_bytes
        self.store = CaptureStore(capture_dir)
        self.context: BrowserContext | None = None
        self.captured = 0
        self.skipped_large = 0
        self.auto_visited: set[str] = set()

    def _launch(self) -> BrowserContext:
        return launch_browser_context(
            self.playwright, self.browser_name, self.profile_dir
        )

    def _handle_response(self, response: Response) -> None:
        request = response.request
        if request.resource_type not in {"xhr", "fetch"}:
            return

        content_type = response.headers.get("content-type", "")
        content_length = response.headers.get("content-length")
        if content_length and content_length.isdigit() and int(content_length) > self.max_response_bytes:
            self.skipped_large += 1
            return

        try:
            body = response.body()
        except Error:
            return

        if len(body) > self.max_response_bytes:
            self.skipped_large += 1
            return

        try:
            response_data = json.loads(body.decode("utf-8-sig"))
        except (UnicodeDecodeError, json.JSONDecodeError):
            return

        candidate = self.store.save_json_exchange(
            url=response.url,
            method=request.method,
            status=response.status,
            content_type=content_type,
            request_body=request.post_data,
            response_data=response_data,
        )
        self.captured += 1
        marker = "候选" if candidate.score >= 5 else "记录"
        print(
            f"[{marker} {candidate.score:02d}] "
            f"{request.method} {candidate.url[:110]}"
        )

    def _find_auto_controls(self, page) -> list[tuple[str, Any, str, str]]:
        found: list[tuple[int, str, Any, str, str]] = []
        for frame in page.frames:
            controls = frame.locator("a, button, [role='button']")
            self._collect_frame_controls(frame, controls, found)
        return [
            (identity, control, text, frame_url)
            for _, identity, control, text, frame_url in sorted(
                found, key=lambda item: (-item[0], item[1])
            )
        ]

    def _collect_frame_controls(
        self, frame, controls, found: list[tuple[int, str, Any, str, str]]
    ) -> None:
        for index in range(min(controls.count(), 150)):
            control = controls.nth(index)
            try:
                if not control.is_visible():
                    continue
                text = " ".join((control.inner_text() or "").split())[:120]
                href = control.get_attribute("href")
                if not _is_relevant_control(text, href) and not _is_safe_portal_fallback(
                    text, href
                ):
                    continue
                identity = f"{frame.url}|{text}|{href or ''}"
                if identity in self.auto_visited:
                    continue
                priority = 2 if _is_relevant_control(text, href) else 1
                found.append((priority, identity, control, text, frame.url))
            except Error:
                continue

    def _auto_navigate(self, page, max_pages: int) -> None:
        """Visit relevant read-only controls without submitting any forms."""
        pages_visited = 0
        while pages_visited < max_pages:
            controls = self._find_auto_controls(page)
            if not controls:
                print("自动导航：当前页面没有新的安全候选入口。")
                return

            identity, control, text, frame_url = controls[0]
            self.auto_visited.add(identity)
            href = control.get_attribute("href")
            current_url = frame_url or page.url
            try:
                if href:
                    target_url = urljoin(current_url, href)
                    if not _is_safe_navigation(current_url, target_url):
                        print(f"自动导航：跳过非同源或非 GET 入口「{text}」")
                        continue
                    print(f"自动导航：访问「{text}」 -> {target_url[:120]}")
                    page.goto(target_url, wait_until="domcontentloaded", timeout=30_000)
                else:
                    print(f"自动导航：点击「{text}」")
                    control.click(timeout=10_000)
                page.wait_for_timeout(1200)
                pages_visited += 1
            except Error as exc:
                print(f"自动导航：入口「{text}」失败，继续寻找：{exc}")

        print(f"自动导航：达到页面上限 {max_pages}，停止继续点击。")

    def _wait_for_login(self, page, timeout_seconds: int):
        deadline = time.monotonic() + timeout_seconds
        print(
            f"\n等待教务系统登录完成，最长 {timeout_seconds} 秒；"
            "登录成功后将自动开始探索。"
        )

        while time.monotonic() < deadline:
            if self.context is None:
                raise RuntimeError("浏览器上下文未初始化")
            active_pages = self.context.pages
            current = active_pages[-1] if active_pages else page
            if not _is_login_url(current.url):
                print(f"\n检测到登录成功：{current.url}")
                return current
            current.wait_for_timeout(500)

        raise TimeoutError(
            "等待教务系统登录超时。请确认浏览器中的登录已完成后重试。"
        )

    def run(
        self,
        url: str,
        *,
        max_pages: int = 12,
        login_timeout_seconds: int = 300,
    ) -> Path:
        self.context = self._launch()
        self.context.on("response", self._handle_response)

        page = self.context.pages[0] if self.context.pages else self.context.new_page()
        print(f"浏览器：{self.browser_name}")
        print(f"登录状态目录：{self.profile_dir}")
        print(f"捕获目录：{self.capture_dir}")
        print("安全边界：只监听 Fetch/XHR JSON，不重放请求、不提交表单。")
        page.goto(url, wait_until="domcontentloaded", timeout=60_000)

        try:
            current = self._wait_for_login(page, login_timeout_seconds)
            print("自动导航：开始寻找课程与培养方案相关入口。")
            self._auto_navigate(current, max_pages)
        except KeyboardInterrupt:
            print("\n收到中断，正在保存已捕获内容...")
        finally:
            candidates_path = self.store.write_candidates()
            self.context.close()

        print(f"\n已捕获 JSON 响应：{self.captured}")
        if self.skipped_large:
            print(f"因超过大小限制跳过：{self.skipped_large}")
        print(f"候选接口：{candidates_path}")
        return candidates_path


def make_capture_session(root: Path) -> Path:
    stamp = datetime.now(timezone.utc).strftime("%Y%m%d-%H%M%S")
    return root / stamp
