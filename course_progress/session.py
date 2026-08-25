"""One authenticated academic-browser session behind a small interface."""

from __future__ import annotations

import time
from pathlib import Path
from urllib.parse import urlsplit

from playwright.sync_api import Browser, BrowserContext, Error, Playwright

from .explorer import _is_login_url, launch_browser_context, resolve_profile_dir
from .credentials import AUTH_STATE_FILE_NAME, credential_store


WEBVPN_CAS_ENTRY_URL = (
    "https://webvpn.hitwh.edu.cn/login?cas_login=true#!/service"
)


def _is_webvpn_credential_page(url: str) -> bool:
    """Restrict credential entry to HITWH's proxied CAS login page."""
    parsed = urlsplit(url)
    return (
        parsed.scheme == "https"
        and parsed.hostname == "webvpn.hitwh.edu.cn"
        and "authserver/login" in parsed.path.lower()
    )


def _is_legacy_webvpn_login(url: str) -> bool:
    parsed = urlsplit(url)
    return (
        parsed.scheme == "https"
        and parsed.hostname == "webvpn.hitwh.edu.cn"
        and parsed.fragment.lower() == "!/login"
    )


class AcademicBrowserSession:
    """Own persistent-browser launch, stable authentication detection, and cleanup."""

    def __init__(
        self,
        playwright: Playwright,
        *,
        browser_name: str,
        profile_root: Path,
        persistent: bool = True,
    ):
        self.playwright = playwright
        self.browser_name = browser_name
        self.private_root = profile_root.resolve()
        self.profile_dir = resolve_profile_dir(self.private_root)
        self.auth_state_path = self.private_root / AUTH_STATE_FILE_NAME
        self.credentials = credential_store(self.private_root)
        self.persistent = persistent
        self.browser: Browser | None = None
        self.context: BrowserContext | None = None
        self.actual_browser_name = browser_name

    def __enter__(self) -> "AcademicBrowserSession":
        if not self.persistent:
            options = {"headless": False}
            if self.browser_name == "chrome":
                options["channel"] = "chrome"
            self.browser = self.playwright.chromium.launch(**options)
            context_options = {"no_viewport": True}
            if self.auth_state_path.is_file():
                context_options["storage_state"] = str(self.auth_state_path)
                print("登录状态：已加载上次 WebVPN 会话。")
            try:
                self.context = self.browser.new_context(**context_options)
            except Error:
                if "storage_state" not in context_options:
                    raise
                print("登录状态：已有状态无法读取，本次从登录页重新认证。")
                self.context = self.browser.new_context(no_viewport=True)
            return self
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
        if self.browser is not None:
            self.browser.close()
            self.browser = None

    def open_authenticated(self, url: str, *, timeout_seconds: int = 600):
        if self.context is None:
            raise RuntimeError("浏览器会话尚未启动")
        if timeout_seconds <= 0:
            raise ValueError("认证等待时间必须大于 0")
        page = self.context.pages[0] if self.context.pages else self.context.new_page()
        page.goto(url, wait_until="domcontentloaded", timeout=60_000)
        page.wait_for_timeout(3_000)
        if _is_legacy_webvpn_login(page.url):
            print("登录状态：WebVPN 会话已失效，切换到统一身份认证。")
            self.context.clear_cookies()
            page.goto(
                WEBVPN_CAS_ENTRY_URL,
                wait_until="domcontentloaded",
                timeout=60_000,
            )
            page.wait_for_timeout(3_000)

        deadline = time.monotonic() + timeout_seconds
        stable_checks = 0
        auto_login_attempted = False
        while time.monotonic() < deadline:
            pages = self.context.pages
            current = pages[-1] if pages else page
            if _is_login_url(current.url):
                stable_checks = 0
                if not auto_login_attempted and self._fill_and_submit_login(current):
                    auto_login_attempted = True
            else:
                stable_checks += 1
                if stable_checks >= 4:
                    self.private_root.mkdir(parents=True, exist_ok=True)
                    self.context.storage_state(path=str(self.auth_state_path))
                    print("登录状态：已更新，下次将优先直接复用。")
                    return current
            current.wait_for_timeout(500)
        raise TimeoutError("等待统一身份认证超时")

    def _fill_and_submit_login(self, page) -> bool:
        if not _is_webvpn_credential_page(page.url):
            return False
        try:
            credentials = self.credentials.load()
        except RuntimeError as error:
            print(f"自动登录：{error}")
            return True
        if credentials is None:
            print("自动登录：尚未配置，运行 configure-login 后可免手输。")
            return True
        username = page.locator(
            "input[placeholder*='学号'], input[name='username'], input[type='text']"
        ).first
        password = page.locator(
            "input[placeholder*='密码'], input[name='password'], input[type='password']"
        ).first
        try:
            if not username.is_visible() or not password.is_visible():
                return False
            username.fill(credentials.username)
            password.fill(credentials.password)
            remember = page.locator("input[type='checkbox']").first
            if remember.is_visible() and not remember.is_checked():
                remember.check()
            submit = page.get_by_role("link", name="登录", exact=True)
            if submit.count() == 0:
                submit = page.get_by_role("button", name="登录", exact=True)
            if submit.count() == 0:
                password.press("Enter")
            else:
                submit.first.click()
            print("自动登录：已在 HITWH 统一认证页填写并提交。")
            return True
        except Error as error:
            print(f"自动登录：页面控件不匹配（{str(error)[:100]}）")
            return False
