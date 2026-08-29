"""Command line entry points for local academic selection exploration."""

from __future__ import annotations

import getpass
import logging
import os
import random
import time
from datetime import datetime
from pathlib import Path

import click
from playwright.sync_api import sync_playwright

from course_progress.credentials import LoginCredentials, credential_store
from course_progress.explorer import (
    DEFAULT_PORTAL_URL,
    _is_login_url,
    launch_browser_context,
    resolve_profile_dir,
)

from . import config
from .discovery import TARGET_SELECTION, TARGET_TIMETABLE, AcademicInterfaceDiscovery
from .notice import (
    CATEGORY_LABELS,
    load_notice,
    notice_selection_categories,
    notice_selection_window_map,
    notice_semester_label,
)
from .selection_entry import (
    STATUS_ENTRY_UNREACHABLE,
    STATUS_LOGIN_REQUIRED,
    SelectionEntryExplorer,
    SelectionObservation,
    save_selection_result,
)
from .student_profile import (
    DEFAULT_STUDENT_PROFILE_PATH,
    create_student_profile,
    load_student_profile,
    save_student_profile,
)

logger = logging.getLogger("course-selection")


def _setup_logging(verbose: bool = False) -> None:
    level = logging.DEBUG if verbose else logging.INFO
    logging.basicConfig(
        format="%(asctime)s [%(levelname)s] %(name)s: %(message)s",
        datefmt="%H:%M:%S",
        level=level,
    )


# ── CLI entry point ──────────────────────────────────────────────────────────


@click.group()
@click.option("--verbose", "-v", is_flag=True, help="启用调试日志")
def main(verbose: bool = False) -> None:
    """选课工作台 CLI — 实验室抢课、选课规划、接口发现等工具集。"""
    _setup_logging(verbose)


# ── workbench ────────────────────────────────────────────────────────────────


@main.command("workbench")
@click.option("--private-root", type=click.Path(path_type=Path), default=config.WORKBENCH_PRIVATE_ROOT)
@click.option("--port", type=int, default=config.WORKBENCH_PORT)
def workbench_cmd(private_root: Path, port: int) -> None:
    """启动本地只读选课规划工作台。"""
    from .application import run_workbench_application

    raise SystemExit(run_workbench_application(private_root, port))


# ── dev-workbench ────────────────────────────────────────────────────────────


@main.command("dev-workbench")
@click.option("--private-root", type=click.Path(path_type=Path), default=config.WORKBENCH_PRIVATE_ROOT)
@click.option("--port", type=int, default=config.WORKBENCH_PORT)
@click.option("--debug-port", type=int, default=config.DEV_DEBUG_PORT)
def dev_workbench_cmd(private_root: Path, port: int, debug_port: int) -> None:
    """开发模式：持久 Chromium + CDP + Python 热重启。"""
    from .dev_workbench import run_dev_workbench

    raise SystemExit(run_dev_workbench(Path.cwd(), private_root, port, debug_port))


# ── configure-login ──────────────────────────────────────────────────────────


@main.command("configure-login")
@click.option("--profile-root", type=click.Path(path_type=Path), default=config.PROGRESS_PROFILE_ROOT)
@click.option("--username", default=None, help="学号/工号；省略时交互输入")
def configure_login_cmd(profile_root: Path, username: str | None) -> None:
    """使用 Windows DPAPI 加密保存 WebVPN 登录信息。"""
    username = (username or input("学号/工号：")).strip()
    password = getpass.getpass("统一身份认证密码（输入不会显示）：")
    credential_store(profile_root).save(LoginCredentials(username=username, password=password))
    click.echo("登录信息已由 Windows DPAPI 加密保存，仅当前 Windows 用户可解密。")


# ── configure-profile ────────────────────────────────────────────────────────


@main.command("configure-profile")
@click.option("--grade", required=True, help="入学年级，例如 2025")
@click.option("--major", default="", help="专业；未知时可省略")
@click.option("--academic-level", default="", help="培养层次；未知时可省略")
@click.option("--campus", default="", help="校区；未知时可省略")
@click.option("--profile", "profile_path", type=click.Path(path_type=Path), default=DEFAULT_STUDENT_PROFILE_PATH)
def configure_profile_cmd(grade: str, major: str, academic_level: str, campus: str, profile_path: Path) -> None:
    """保存用于通知与培养方案匹配的本地学生画像。"""
    profile = create_student_profile(
        grade=grade, major=major, academic_level=academic_level, campus=campus,
    )
    save_student_profile(profile_path, profile)
    click.echo(f"学生画像已保存：{profile.grade}级；{profile_path}")


# ── analyze-interface ────────────────────────────────────────────────────────


@main.command("analyze-interface")
@click.option("--target", type=click.Choice(["student-profile"]), required=True)
@click.option("--url", default=DEFAULT_PORTAL_URL, help="WebVPN 门户入口")
@click.option("--profile-root", type=click.Path(path_type=Path), default=config.PROGRESS_PROFILE_ROOT)
@click.option("--output-root", type=click.Path(path_type=Path), default=Path(".private/interface-analysis"))
@click.option("--login-timeout-seconds", type=int, default=600)
@click.option("--wait-seconds", type=int, default=60)
def analyze_interface_cmd(
    target: str, url: str, profile_root: Path, output_root: Path,
    login_timeout_seconds: int, wait_seconds: int,
) -> None:
    """只读分析学校页面提供的接口契约。"""
    from .student_profile_observation import StudentProfileInterfaceAnalyzer

    analyzer = StudentProfileInterfaceAnalyzer(
        profile_root=profile_root.resolve(),
        output_root=output_root.resolve(),
        portal_url=url,
        login_timeout_seconds=login_timeout_seconds,
        wait_seconds=wait_seconds,
    )
    output = analyzer.run()
    click.echo(f"候选学生画像接口：{output}")


# ── explore-entry ────────────────────────────────────────────────────────────


@main.command("explore-entry")
@click.option("--notice", type=click.Path(path_type=Path), default=config.WORKBENCH_PRIVATE_ROOT / "selection-notice.json")
@click.option("--url", default=DEFAULT_PORTAL_URL, help="教务门户入口")
@click.option("--browser", type=click.Choice(["chromium"]), default="chromium")
@click.option("--private-root", type=click.Path(path_type=Path), default=config.WORKBENCH_PRIVATE_ROOT)
@click.option("--profile-root", type=click.Path(path_type=Path), default=config.PROGRESS_PROFILE_ROOT)
@click.option("--login-timeout-seconds", type=int, default=600)
@click.option("--wait-seconds", type=int, default=600)
def explore_entry_cmd(
    notice: Path, url: str, browser: str, private_root: Path,
    profile_root: Path, login_timeout_seconds: int, wait_seconds: int,
) -> None:
    """只读观察已确认通知对应的教务选课入口。"""
    if login_timeout_seconds <= 0 or wait_seconds <= 0:
        raise SystemExit("等待时间必须大于 0")
    if not notice.is_file():
        raise SystemExit(f"选课通知不存在：{notice}")
    notice_obj = load_notice(notice)
    if notice_obj.status != "confirmed":
        raise SystemExit("选课通知尚未确认，不能探索对应入口")
    profile_dir = resolve_profile_dir(profile_root.resolve())
    with sync_playwright() as playwright:
        context = launch_browser_context(playwright, browser, profile_dir)
        try:
            page = context.pages[0] if context.pages else context.new_page()
            try:
                page.goto(url, wait_until="domcontentloaded", timeout=60_000)
            except Exception as error:
                save_selection_result(
                    private_root.resolve(),
                    SelectionObservation(
                        status=STATUS_ENTRY_UNREACHABLE,
                        request_url=url,
                        method="GET",
                        message=str(error),
                        sections=(),
                    ),
                )
                raise SystemExit(1)
            deadline = time.monotonic() + login_timeout_seconds
            while _is_login_url(page.url) and time.monotonic() < deadline:
                page.wait_for_timeout(500)
            if _is_login_url(page.url):
                save_selection_result(
                    private_root.resolve(),
                    SelectionObservation(
                        status=STATUS_LOGIN_REQUIRED,
                        request_url=page.url,
                        method="GET",
                        message="等待统一身份认证超时",
                        sections=(),
                    ),
                )
                raise SystemExit(1)
            explorer = SelectionEntryExplorer(
                notice=notice_obj,
                output_root=private_root.resolve(),
            )
            explorer.run(context, wait_seconds=wait_seconds)
        finally:
            context.close()
    click.echo(f"只读选课入口结果：{private_root.resolve() / 'selection-entry.json'}")


# ── discover-timetable / discover-selection ──────────────────────────────────


@main.command("discover-timetable")
@click.option("--url", default=DEFAULT_PORTAL_URL, help="教务门户入口")
@click.option("--browser", type=click.Choice(["chromium"]), default="chromium")
@click.option("--private-root", type=click.Path(path_type=Path), default=config.WORKBENCH_PRIVATE_ROOT)
@click.option("--profile-root", type=click.Path(path_type=Path), default=config.PROGRESS_PROFILE_ROOT)
@click.option("--login-timeout-seconds", type=int, default=600)
@click.option("--wait-seconds", type=int, default=30)
@click.option("--max-clicks", type=int, default=8)
@click.option("--persistent-session", is_flag=True, help="复用持久登录 profile")
def discover_timetable_cmd(
    url: str, browser: str, private_root: Path, profile_root: Path,
    login_timeout_seconds: int, wait_seconds: int, max_clicks: int,
    persistent_session: bool,
) -> None:
    """自动点击并发现课表只读接口。"""
    _run_discovery(
        target=TARGET_TIMETABLE, url=url, browser=browser,
        private_root=private_root, profile_root=profile_root,
        login_timeout_seconds=login_timeout_seconds, wait_seconds=wait_seconds,
        max_clicks=max_clicks, persistent_session=persistent_session,
    )


@main.command("discover-selection")
@click.option("--url", default=DEFAULT_PORTAL_URL, help="教务门户入口")
@click.option("--browser", type=click.Choice(["chromium"]), default="chromium")
@click.option("--private-root", type=click.Path(path_type=Path), default=config.WORKBENCH_PRIVATE_ROOT)
@click.option("--profile-root", type=click.Path(path_type=Path), default=config.PROGRESS_PROFILE_ROOT)
@click.option("--login-timeout-seconds", type=int, default=600)
@click.option("--wait-seconds", type=int, default=30)
@click.option("--max-clicks", type=int, default=8)
@click.option("--persistent-session", is_flag=True, help="复用持久登录 profile")
@click.option("--notice", type=click.Path(path_type=Path), default=config.WORKBENCH_PRIVATE_ROOT / "selection-notice.json")
@click.option("--grade", default=None, help="临时覆盖学生画像中的入学年级")
@click.option("--profile", "profile_path", type=click.Path(path_type=Path), default=DEFAULT_STUDENT_PROFILE_PATH)
def discover_selection_cmd(
    url: str, browser: str, private_root: Path, profile_root: Path,
    login_timeout_seconds: int, wait_seconds: int, max_clicks: int,
    persistent_session: bool, notice: Path, grade: str | None, profile_path: Path,
) -> None:
    """自动点击并发现选课只读接口。"""
    _run_discovery(
        target=TARGET_SELECTION, url=url, browser=browser,
        private_root=private_root, profile_root=profile_root,
        login_timeout_seconds=login_timeout_seconds, wait_seconds=wait_seconds,
        max_clicks=max_clicks, persistent_session=persistent_session,
        notice=notice, grade=grade, profile_path=profile_path,
    )


def _run_discovery(
    target: str,
    url: str,
    browser: str,
    private_root: Path,
    profile_root: Path,
    login_timeout_seconds: int,
    wait_seconds: int,
    max_clicks: int,
    persistent_session: bool,
    notice: Path | None = None,
    grade: str | None = None,
    profile_path: Path | None = None,
) -> None:
    if login_timeout_seconds <= 0 or wait_seconds <= 0 or max_clicks <= 0:
        raise SystemExit("等待时间和自动点击次数必须大于 0")
    allowed_categories: tuple[str, ...] = ()
    allowed_windows: dict = {}
    semester_label = ""
    if target == TARGET_SELECTION:
        if notice is None or not notice.is_file():
            raise SystemExit(f"选课通知不存在：{notice}")
        notice_obj = load_notice(notice)
        if notice_obj.status != "confirmed":
            raise SystemExit("选课通知尚未确认，禁止查询任何课程类别")
        if grade:
            resolved_grade = create_student_profile(grade=grade).grade
        else:
            if profile_path is None or not profile_path.is_file():
                raise SystemExit("学生画像不存在，请先运行 configure-profile --grade 2025")
            resolved_grade = load_student_profile(profile_path).grade
        allowed_categories = notice_selection_categories(notice_obj, grade=resolved_grade)
        allowed_windows = notice_selection_window_map(notice_obj, grade=resolved_grade)
        semester_label = notice_semester_label(notice_obj)
        if not allowed_categories:
            raise SystemExit("选课通知未明确开放课程类别，禁止猜测查询入口")
        if not semester_label:
            raise SystemExit("选课通知未明确学期，禁止查询")
        labels = "、".join(CATEGORY_LABELS[code] for code in allowed_categories)
        click.echo(f"通知白名单：{semester_label}；{resolved_grade}级；{labels}")
    with sync_playwright() as playwright:
        click.echo("安全边界：只自动点击导航/查询控件；疑似选课、退课、保存请求会被拦截。")
        report = AcademicInterfaceDiscovery(
            playwright,
            browser_name=browser,
            profile_root=profile_root,
            output_root=private_root.resolve(),
            persistent_session=persistent_session,
        ).discover(
            target,
            portal_url=url,
            login_timeout_seconds=login_timeout_seconds,
            wait_seconds=wait_seconds,
            max_clicks=max_clicks,
            allowed_selection_categories=allowed_categories,
            allowed_selection_windows=allowed_windows,
            notice_semester=semester_label,
        )
    click.echo(
        f"发现结果：目标页面={'是' if report.target_found else '否'}，"
        f"点击={report.clicks}，接口={report.captures}，拦截={report.blocked_requests}"
    )


# ── cas-book (lab-book) ──────────────────────────────────────────────────────

# 随机 UA 池，防检测
_USER_AGENTS = [
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/125.0.0.0 Safari/537.36",
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/126.0.0.0 Safari/537.36",
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/127.0.0.0 Safari/537.36",
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/128.0.0.0 Safari/537.36",
]

_BASE_URL = "http://openlab.hitwh.edu.cn"
_BOOKING_URL = f"{_BASE_URL}/dxwl/booking/#/booking"
_STORAGE_FILE = Path("storage_state.json")
_LOGIN_TIMEOUT_SEC = 180
_POLL_INTERVAL = 5
_MAX_POLL_MINUTES = 120
_JITTER_MIN = 0.3
_JITTER_MAX = 1.5

TARGET_COURSES = [
    "DIY电磁混合磁悬浮",
    "磁阻效应",
    "表面张力",
    "偏振光",
]


def _jitter() -> None:
    time.sleep(random.uniform(_JITTER_MIN, _JITTER_MAX))


def _get_chrome_path() -> str | None:
    paths = [
        r"C:\Program Files\Google\Chrome\Application\chrome.exe",
        r"C:\Program Files (x86)\Google\Chrome\Application\chrome.exe",
    ]
    for p in paths:
        if os.path.exists(p):
            return p
    return None


def _safe_click(page, locator, timeout: int = 5000) -> bool:
    try:
        locator.wait_for(timeout=timeout)
        _jitter()
        locator.click()
        return True
    except Exception:
        return False


def _try_close_dialog(page) -> bool:
    try:
        dialog = page.locator(".ant-modal-content").first
        if dialog.is_visible():
            ok_btn = dialog.locator("button", has_text="确定")
            if ok_btn.count() > 0:
                ok_btn.click()
                page.wait_for_timeout(500)
                return True
    except Exception:
        pass
    return False


def _is_logged_in(page) -> bool:
    try:
        page.goto(_BOOKING_URL, timeout=15000)
        page.wait_for_timeout(3000)
        return "login" not in page.url.lower()
    except Exception:
        return False


def _do_cas_login(page) -> bool:
    logger.info("请在打开的浏览器中手动登录")
    logger.info("  步骤: 点击「统一身份认证登录」→ 输学号/密码/验证码 → 登录")
    logger.info("  等待最长 %d 秒...", _LOGIN_TIMEOUT_SEC)

    page.goto(_BASE_URL, timeout=15000)
    page.wait_for_timeout(2000)

    try:
        cas_btn = page.get_by_text("统一身份认证登录").first
        if cas_btn.is_visible():
            cas_btn.click()
            logger.info("已点击「统一身份认证登录」，页面将跳转到 HIT CAS")
        else:
            logger.info("CAS 登录按钮不可见，可能已自动跳转")
    except Exception:
        logger.info("按钮未找到，可能已跳转到 CAS 页面")

    start = time.time()
    while time.time() - start < _LOGIN_TIMEOUT_SEC:
        time.sleep(2)
        current_url = page.url
        if "ids.hit.edu.cn" not in current_url:
            logger.info("检测到登录完成！当前 URL: %s", current_url[:80])
            page.wait_for_timeout(3000)
            return True

    logger.info("登录超时")
    return False


def _ensure_logged_in(context) -> bool:
    if _STORAGE_FILE.exists():
        logger.info("发现已保存的会话文件: %s", _STORAGE_FILE)
        page = context.new_page()

        if _is_logged_in(page):
            logger.info("会话有效 ✓")
            page.close()
            return True

        logger.info("会话已过期，需要重新登录")
        page.close()

    page = context.new_page()

    if not _do_cas_login(page):
        page.close()
        return False

    if not _is_logged_in(page):
        logger.info("登录后仍无法访问预约页面")
        page.close()
        return False

    context.storage_state(path=str(_STORAGE_FILE))
    logger.info("会话已保存到 %s", _STORAGE_FILE)
    page.close()
    return True


def _go_back(page) -> None:
    try:
        back_btn = page.locator("button").filter(has_text="返回").first
        if back_btn.count() > 0:
            back_btn.click()
            page.wait_for_timeout(1000)
    except Exception:
        pass


def _do_booking(page, course_name: str) -> bool:
    logger.info("▶ 尝试预约: %s", course_name)

    try:
        page.goto(_BOOKING_URL, timeout=20000)
        page.wait_for_timeout(3000)
    except Exception:
        logger.info("  预约页加载失败")
        page.close()
        return False

    if not _safe_click(page, page.locator("button").filter(has_text="选课").first):
        book_link = page.get_by_text("选课").first
        if book_link.count() == 0:
            logger.info("  找不到「选课」按钮")
            page.close()
            return False
        book_link.click()
    page.wait_for_timeout(2000)

    try:
        course_cell = page.locator("td").filter(has_text=course_name).first
        if course_cell.count() == 0:
            logger.info("  未找到课程「%s」", course_name)
            _go_back(page)
            page.close()
            return False
        course_cell.click()
        page.wait_for_timeout(2000)
    except Exception as e:
        logger.info("  选择课程异常: %s", e)
        _go_back(page)
        page.close()
        return False

    try:
        rows = page.locator("table tr").all()
        dates = []
        for row in rows:
            cell = row.locator("td").first
            if cell.count() > 0:
                text = cell.text_content().strip()
                if text and any(c.isdigit() for c in text) and "周" not in text:
                    dates.append(text)
    except Exception:
        dates = []

    if not dates:
        logger.info("  未找到可选日期")
        _go_back(page)
        page.close()
        return False

    for date in dates:
        logger.info("  尝试日期: %s", date)
        try:
            date_cell = page.locator("td").filter(has_text=date).first
            if date_cell.count() == 0:
                continue
            date_cell.click()
            _jitter()

            query_btn = page.locator("button").filter(has_text="查询").first
            if query_btn.count() > 0:
                query_btn.click()
                page.wait_for_timeout(2000)

            try:
                dialog = page.locator(".ant-modal-content").first
                if dialog.is_visible(timeout=3000):
                    content = dialog.text_content() or ""
                    if any(w in content for w in ["没有可供选择", "已满", "人数已满"]):
                        logger.info("      ✗ %s %s 已满", course_name, date)
                        _try_close_dialog(page)
                        continue
                    else:
                        logger.info("      弹窗: %s", content[:60])
            except Exception:
                pass

            seats = page.locator("input[type='radio']")
            if seats.count() == 0:
                logger.info("      %s 无可用座位", date)
                continue

            seats.first.click()
            _jitter()

            confirm_btn = page.locator("button").filter(has_text="确认").first
            if confirm_btn.count() == 0:
                confirm_btn = page.get_by_text("确认").first
            if confirm_btn.count() > 0:
                confirm_btn.click()
                page.wait_for_timeout(3000)
                logger.info("      ✓✓✓ 预约成功！%s @ %s", course_name, date)
                page.close()
                return True

            logger.info("  找不到确认按钮")

        except Exception as e:
            logger.info("  日期处理异常: %s", e)
            continue

    _go_back(page)
    page.close()
    return False


def _build_context(browser):
    ua = random.choice(_USER_AGENTS)
    context = browser.new_context(
        viewport={"width": 1400, "height": 900},
        user_agent=ua,
    )
    context.route("**/*", lambda route: (
        route.abort()
        if route.request.resource_type in ("image", "font", "media")
        else route.continue_()
    ))
    return context


def _run_cas_book(
    courses: list[str],
    login_only: bool = False,
    monitor: bool = False,
    poll_interval: int = _POLL_INTERVAL,
) -> None:
    click.echo()
    click.echo("=" * 60)
    click.echo("  开放式实验系统 - 自动预约脚本 (CAS SSO)")
    click.echo(f"  {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    click.echo("=" * 60)

    chrome_path = _get_chrome_path()
    if chrome_path:
        logger.info("Chrome: %s", chrome_path)

    with sync_playwright() as p:
        browser = p.chromium.launch(
            headless=False,
            executable_path=chrome_path or None,
        )
        context = _build_context(browser)

        if not _ensure_logged_in(context):
            logger.info("登录失败，退出")
            input("\n按 Enter 关闭浏览器...")
            browser.close()
            raise SystemExit(1)

        if login_only:
            logger.info("登录完成（--login-only）。会话已保存。")
            input("\n按 Enter 关闭浏览器...")
            browser.close()
            return

        if monitor:
            max_rounds = (_MAX_POLL_MINUTES * 60) // poll_interval
            logger.info("监控模式启动 (间隔 %ds, 最长 %d 分钟)", poll_interval, _MAX_POLL_MINUTES)
            logger.info("课程优先级: %s", ", ".join(courses))

            for rnd in range(1, max_rounds + 1):
                logger.info("── 第 %d 轮 ──", rnd)

                for course in courses:
                    page = context.new_page()
                    if _do_booking(page, course):
                        click.echo("\n🎉🎉🎉 预约成功！已退出")
                        input("\n按 Enter 关闭浏览器...")
                        browser.close()
                        return

                if rnd < max_rounds:
                    logger.info("等待 %d 秒...", poll_interval)
                    time.sleep(poll_interval)

            logger.info("⏰ 监控超时 (%d 分钟)", _MAX_POLL_MINUTES)
        else:
            logger.info("单次模式：尝试 %d 门课程", len(courses))
            for course in courses:
                page = context.new_page()
                if _do_booking(page, course):
                    click.echo("\n🎉 预约成功！")
                    break
            else:
                click.echo("\n😔 所有课程均无可预约时段")

        input("\n按 Enter 关闭浏览器...")
        browser.close()


@main.command("cas-book")
@click.option("--course", "courses", multiple=True, help="指定课程名（可重复），默认使用全部目标课程")
@click.option("--monitor", is_flag=True, help="监控模式，每 N 秒检查一次")
@click.option("--interval", type=int, default=config.CAS_BOOK_POLL_INTERVAL, help=f"监控间隔（秒），默认 {config.CAS_BOOK_POLL_INTERVAL}")
@click.option("--login-only", is_flag=True, help="仅登录保存会话，不预约")
def cas_book_cmd(courses: tuple[str, ...], monitor: bool, interval: int, login_only: bool) -> None:
    """CAS SSO 登录 → 开放式实验系统 → 自动预约。"""
    _run_cas_book(
        courses=list(courses) if courses else TARGET_COURSES,
        login_only=login_only,
        monitor=monitor,
        poll_interval=interval,
    )


# ── lab-book entry point (for pyproject.toml [project.scripts]) ──────────────


@click.group()
def lab_book_main() -> None:
    """实验室抢课工具（兼容旧命令行入口）。"""
    _setup_logging(False)


@lab_book_main.command("cas-book")
@click.option("--course", "courses", multiple=True, help="指定课程名")
@click.option("--monitor", is_flag=True, help="监控模式")
@click.option("--interval", type=int, default=_POLL_INTERVAL, help=f"监控间隔（秒），默认 {_POLL_INTERVAL}")
@click.option("--login-only", is_flag=True, help="仅登录保存会话")
def _lab_book_cas_book(courses: tuple[str, ...], monitor: bool, interval: int, login_only: bool) -> None:
    """CAS SSO 登录自动预约。"""
    _run_cas_book(
        courses=list(courses) if courses else TARGET_COURSES,
        login_only=login_only,
        monitor=monitor,
        poll_interval=interval,
    )


if __name__ == "__main__":
    main()