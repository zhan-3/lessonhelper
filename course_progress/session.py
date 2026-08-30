"""One authenticated academic-browser session behind a small interface."""

from __future__ import annotations

import time
from pathlib import Path
from urllib.parse import urljoin, urlsplit

from playwright.sync_api import Browser, BrowserContext, Error, Playwright
from typing_extensions import Self

from .credentials import AUTH_STATE_FILE_NAME, credential_store
from .explorer import _is_login_url, launch_browser_context, resolve_profile_dir

WEBVPN_CAS_ENTRY_URL = (
    "https://webvpn.hitwh.edu.cn/login?cas_login=true#!/service"
)
WEBVPN_USER_INFO_URL = "https://webvpn.hitwh.edu.cn/user/info"

# CAS 中转页自动重导航的上限：到顶后转入被动等待，不再反复刷新教务系统。
CAS_RENAVIGATE_LIMIT = 5

# WebVPN 门户点击资源 tile 时用 window.open(url, "_blank", "noopener,noreferrer")
# 打开代理资源。该 popup 在门户渲染器不健康(未认证时 CPU 空转)时会把新标签
# 僵死在空 URL——连重定向目标都不落地。这里 hook window.open 捕获目标 URL 并
# 阻止真正开窗,由调用方拿到 URL 后自己导航,彻底绕开不可靠的 popup。
_PORTAL_OPEN_CAPTURE = """
() => {
    if (!window.__pi_portal_open_hooked) {
        window.__pi_portal_open_native = window.open;
        window.__pi_portal_open_hooked = true;
    }
    window.__pi_portal_open_url = null;
    window.open = function (url) {
        window.__pi_portal_open_url = String(url);
        return null;
    };
}
"""
_PORTAL_OPEN_READ = "() => window.__pi_portal_open_url"
_PORTAL_OPEN_RESTORE = """
() => {
    if (window.__pi_portal_open_native) {
        window.open = window.__pi_portal_open_native;
        window.__pi_portal_open_hooked = false;
    }
}
"""


class WebVpnSessionExpiredError(RuntimeError):
    """WebVPN rejected the browser session after a seemingly successful login."""


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
        cdp_url: str | None = None,
    ):
        self.playwright = playwright
        if browser_name != "chromium":
            raise ValueError("academic browser must use Playwright Chromium")
        self.browser_name = "chromium"
        self.private_root = profile_root.resolve()
        self.profile_dir = resolve_profile_dir(self.private_root)
        self.auth_state_path = self.private_root / AUTH_STATE_FILE_NAME
        self.credentials = credential_store(self.private_root)
        self.persistent = persistent
        self.cdp_url = cdp_url
        self.browser: Browser | None = None
        self._attached_over_cdp = False
        self.context: BrowserContext | None = None
        self.actual_browser_name = "chromium"

    def __enter__(self) -> Self:
        if self.cdp_url:
            browser = self.playwright.chromium.connect_over_cdp(self.cdp_url)
            if not browser.contexts:
                raise RuntimeError("CDP browser has no persistent context")
            self.browser = browser
            self.context = browser.contexts[0]
            self._attached_over_cdp = True
            return self
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
        self.context = launch_browser_context(
            self.playwright, "chromium", self.profile_dir
        )
        return self

    def __exit__(self, exc_type, exc_value, traceback) -> None:
        if self._attached_over_cdp:
            # The development Browser Host owns the context and browser.  An
            # attached workbench must only detach so hot reload preserves tabs.
            self.context = None
            self.browser = None
            self._attached_over_cdp = False
            return
        if self.context is not None:
            self.context.close()
            self.context = None
        if self.browser is not None:
            self.browser.close()
            self.browser = None

    @staticmethod
    def _is_loopback_url(url: str) -> bool:
        return (urlsplit(url).hostname or "").lower() in {"127.0.0.1", "localhost", "::1"}

    def _default_academic_page(self):
        """复用已存在的非 loopback 选项卡；没有才新建（保持选项卡单例）。"""
        if self.context is None:
            raise RuntimeError("浏览器会话尚未启动")
        for candidate in self.context.pages:
            checker = getattr(candidate, "is_closed", None)
            if checker is not None and checker():
                continue
            if self._is_loopback_url(getattr(candidate, "url", "") or ""):
                continue
            return candidate
        return self.context.new_page()

    def open_authenticated(self, url: str, *, timeout_seconds: int = 600, page=None):
        if self.context is None:
            raise RuntimeError("浏览器会话尚未启动")
        if timeout_seconds <= 0:
            raise ValueError("认证等待时间必须大于 0")
        page = page or self._default_academic_page()
        page.goto(url, wait_until="domcontentloaded", timeout=60_000)
        page.wait_for_timeout(3_000)
        reauthenticated = False
        if _is_login_url(page.url) and "logincas" not in page.url.lower():
            # Any WebVPN login page (legacy "#!/login", plain "/login", or the
            # logoutByIpChange kick page) means this context's cookies no longer
            # map to a session this egress can use.  Drop them and start the
            # unified-authentication flow with a clean ticket.
            print("登录状态：WebVPN 会话已失效，切换到统一身份认证。")
            self.context.clear_cookies()
            reauthenticated = True
            page.goto(
                WEBVPN_CAS_ENTRY_URL,
                wait_until="domcontentloaded",
                timeout=60_000,
            )
            page.wait_for_timeout(3_000)

        deadline = time.monotonic() + timeout_seconds
        stable_checks = 0
        next_auto_login_at = 0.0
        auto_login_attempts = 0
        next_cas_renavigate_at = 0.0
        cas_renavigate_attempts = 0
        while time.monotonic() < deadline:
            # Authentication belongs to the requested academic tab.  The
            # persistent context also contains the loopback workbench shell;
            # selecting context.pages[-1] can mistake that shell for a
            # successful academic login and leave the real tab on CAS.
            current = page
            now = time.monotonic()
            if _is_login_url(current.url):
                healed = False
                if "logoutbyipchange" in current.url.lower():
                    # The kick is a one-shot IP-binding check that self-heals
                    # within a second or two (the SPA re-establishes the token
                    # session).  If the request layer already accepts /user/info
                    # again, don't treat the transient navigation as a real
                    # logout. Resume the originally requested URL in the same
                    # tab instead of returning a healthy session on a login page.
                    try:
                        webvpn_api_get(self.context, WEBVPN_USER_INFO_URL, attempts=1)
                    except WebVpnSessionExpiredError:
                        pass
                    else:
                        current.goto(
                            url,
                            wait_until="domcontentloaded",
                            timeout=60_000,
                        )
                        stable_checks = 0
                        healed = True
                        reauthenticated = True
                if (
                    not healed
                    and cas_renavigate_attempts < CAS_RENAVIGATE_LIMIT
                    and "logincas" in current.url.lower()
                    and now >= next_cas_renavigate_at
                ):
                    # loginCAS 是教务系统自己的 CAS 中转页(带或不带 ?ticket=)。
                    # 带 ticket 时说明统一身份认证已签发成功、后端会话已建立；
                    # 不带 ticket 也可能是 SSO 跳转被 WebVPN 改写而停滞。该中转页
                    # 的后续跳转常被改写中断,重新导航到目标 URL 即可验证会话是否
                    # 已可用,而不是把已登录/卡住的中转页当作仍在登录而无限等待。
                    # 重导航次数有上限：会话真坏了时反复刷新只会骚扰教务系统,
                    # 到顶后转入被动等待（用户可手动接管,或等超时）。
                    cas_renavigate_attempts += 1
                    next_cas_renavigate_at = now + 5.0
                    try:
                        current.goto(
                            url, wait_until="domcontentloaded", timeout=60_000
                        )
                    except (Error, TimeoutError):
                        pass
                    stable_checks = 0
                    healed = True
                    if cas_renavigate_attempts == CAS_RENAVIGATE_LIMIT:
                        print(
                            "登录提示：自动重试教务系统入口已达上限，"
                            "停止自动刷新；请在浏览器里手动完成登录或稍后重试任务。"
                        )
                if not healed:
                    stable_checks = 0
                    now = time.monotonic()
                    if (
                        auto_login_attempts < 2
                        and now >= next_auto_login_at
                        and self._fill_and_submit_login(current)
                    ):
                        # Some CAS pages accept the DOM interaction before their
                        # client-side handlers are ready and silently ignore it.
                        # Permit one delayed retry, never an unbounded submit loop.
                        auto_login_attempts += 1
                        next_auto_login_at = now + 3.0
                        reauthenticated = True
            else:
                stable_checks += 1
                if stable_checks >= 4:
                    # Persist only after an actual re-authentication: persistent
                    # contexts keep cookies in the profile dir, and a healthy
                    # session's storage state is unchanged.  Printing on every
                    # navigation made multi-page reads look like repeated logins.
                    if reauthenticated or not getattr(self, "persistent", True):
                        self.private_root.mkdir(parents=True, exist_ok=True)
                        self.context.storage_state(path=str(self.auth_state_path))
                    if reauthenticated:
                        print("登录状态：已更新，下次将优先直接复用。")
                    return current
            current.wait_for_timeout(500)
        raise TimeoutError("等待统一身份认证超时")

    def assert_webvpn_session(self) -> None:
        """Verify that WebVPN accepts this context before entering an app tile."""
        if self.context is None:
            raise RuntimeError("浏览器会话尚未启动")
        webvpn_api_get(self.context, WEBVPN_USER_INFO_URL)

    def open_portal_application(
        self,
        portal_url: str,
        application_name: str,
        *,
        timeout_seconds: int = 120,
        page=None,
    ):
        """Open one rendered WebVPN resource in the current browser context."""
        if self.context is None:
            raise RuntimeError("浏览器会话尚未启动")
        if timeout_seconds <= 0:
            raise ValueError("应用入口等待时间必须大于 0")
        page = self.open_authenticated(
            portal_url, timeout_seconds=timeout_seconds, page=page
        )
        self.assert_webvpn_session()
        deadline = time.monotonic() + timeout_seconds
        while time.monotonic() < deadline:
            locator = page.get_by_text(application_name, exact=True)
            visible = None
            for index in range(locator.count()):
                candidate = locator.nth(index)
                if candidate.is_visible():
                    visible = candidate
                    break
            if visible is None:
                page.wait_for_timeout(250)
                continue

            # 门户 tile 用 window.open(url, "_blank", "noopener,noreferrer")
            # 打开代理资源。该 popup 在门户渲染器不健康时会僵死在空 URL。
            # hook 捕获 window.open 的目标 URL 并阻止真正开窗,再由我们
            # 主动导航,绕开不可靠的 popup。
            page.evaluate(_PORTAL_OPEN_CAPTURE)
            try:
                visible.click(timeout=10_000)
                captured = page.evaluate(_PORTAL_OPEN_READ)
            finally:
                page.evaluate(_PORTAL_OPEN_RESTORE)
            if not captured:
                page.wait_for_timeout(250)
                continue

            target_url = urljoin(page.url, captured)
            # 复用当前选项卡直接导航到资源,不再另开标签;会话失效时也在
            # 同一选项卡刷新重登,保持「工作台 + 学术页」两个选项卡。
            remaining = max(1, int(deadline - time.monotonic()))
            for _attempt in range(2):
                page.goto(target_url, wait_until="domcontentloaded", timeout=60_000)
                if not _is_login_url(page.url):
                    break
                page = self.open_authenticated(
                    target_url,
                    timeout_seconds=remaining,
                    page=page,
                )
                remaining = max(1, int(deadline - time.monotonic()))
            return page
        raise TimeoutError(f"等待 WebVPN 应用入口超时：{application_name}")

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
                # HITWH CAS defaults to the QR-code tab, which hides the
                # account form entirely.  Switch to the account-login tab
                # (an anchor that reloads with ?type=userNameLogin) first.
                account_tab = page.get_by_role("link", name="账号登录")
                if account_tab.count() == 0:
                    account_tab = page.get_by_text("账号登录", exact=True).first
                if account_tab.count():
                    account_tab.first.click()
                    page.wait_for_timeout(2_000)
                if not username.is_visible():
                    return False
            captcha = page.locator("input[name='captcha']").first
            if captcha.count() and captcha.is_visible() and not captcha.input_value():
                print("自动登录：需要人工输入验证码，请在浏览器中完成登录。")
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


def webvpn_api_get(context, url, *, attempts: int = 3, settle_delay: float = 2.5):
    """GET a WebVPN API endpoint through the shared context request layer.

    Wengine requires the SPA-XHR header (plain requests are treated as page
    navigations and 302 to /login even for a healthy session) and fires a
    one-shot ``logoutByIpChange`` kick on the first request right after a
    fresh login, self-healing within a second or two.  Bounded retries absorb
    that settle window; a persistently redirected session still surfaces as a
    typed ``WebVpnSessionExpiredError``.
    """
    last_error: Error | None = None
    for attempt in range(attempts):
        try:
            response = context.request.get(
                url,
                timeout=30_000,
                max_redirects=0,
                headers={"X-Requested-With": "XMLHttpRequest"},
            )
        except Error as error:
            last_error = error
        else:
            if response.status == 200 and not _is_login_url(response.url):
                return response
            last_error = None
        if attempt < attempts - 1:
            time.sleep(settle_delay)
    raise WebVpnSessionExpiredError(
        "WebVPN 会话已失效（可能因网络/IP 变化）；请在浏览器中重新认证"
    ) from last_error
