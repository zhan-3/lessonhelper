"""One authenticated academic-browser session behind a small interface."""

from __future__ import annotations

import time
from pathlib import Path

from playwright.sync_api import BrowserContext, Error, Playwright

from .explorer import _is_login_url, launch_browser_context, resolve_profile_dir


class AcademicBrowserSession:
    """Own persistent-browser launch, stable authentication detection, and cleanup."""

    def __init__(
        self,
        playwright: Playwright,
        *,
        browser_name: str,
        profile_root: Path,
    ):
        self.playwright = playwright
        self.browser_name = browser_name
        self.profile_dir = resolve_profile_dir(profile_root.resolve())
        self.context: BrowserContext | None = None
        self.actual_browser_name = browser_name

    def __enter__(self) -> "AcademicBrowserSession":
        try:
            self.context = launch_browser_context(
                self.playwright, self.browser_name, self.profile_dir
            )
        except Error:
            if self.browser_name != "chromium":
                raise
            print("Chromium 无法读取现有登录 profile，改用系统 Chrome 重试。")
            self.actual_browser_name = "chrome"
            self.context = launch_browser_context(
                self.playwright, "chrome", self.profile_dir
            )
        return self

    def __exit__(self, exc_type, exc_value, traceback) -> None:
        if self.context is not None:
            self.context.close()
            self.context = None

    def open_authenticated(self, url: str, *, timeout_seconds: int = 600):
        if self.context is None:
            raise RuntimeError("浏览器会话尚未启动")
        if timeout_seconds <= 0:
            raise ValueError("认证等待时间必须大于 0")
        page = self.context.pages[0] if self.context.pages else self.context.new_page()
        page.goto(url, wait_until="domcontentloaded", timeout=60_000)
        page.wait_for_timeout(3_000)

        deadline = time.monotonic() + timeout_seconds
        stable_checks = 0
        while time.monotonic() < deadline:
            pages = self.context.pages
            current = pages[-1] if pages else page
            if _is_login_url(current.url):
                stable_checks = 0
            else:
                stable_checks += 1
                if stable_checks >= 4:
                    return current
            current.wait_for_timeout(500)
        raise TimeoutError("等待统一身份认证超时")
