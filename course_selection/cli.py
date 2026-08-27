"""Command line entry points for local academic selection exploration."""

from __future__ import annotations

import argparse
import getpass
import time
from pathlib import Path

from playwright.sync_api import sync_playwright

from course_progress.explorer import (
    DEFAULT_PORTAL_URL,
    _is_login_url,
    launch_browser_context,
    resolve_profile_dir,
)
from course_progress.credentials import LoginCredentials, credential_store

from .notice import (
    CATEGORY_LABELS,
    load_notice,
    notice_selection_categories,
    notice_selection_window_map,
    notice_semester_label,
)
from .discovery import AcademicInterfaceDiscovery, TARGET_SELECTION, TARGET_TIMETABLE
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


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(prog="course-selection")
    subparsers = parser.add_subparsers(dest="command", required=True)
    workbench = subparsers.add_parser("workbench", help="启动本地只读选课规划工作台")
    workbench.add_argument("--private-root", type=Path, default=Path(".private/academic-selection"))
    workbench.add_argument("--port", type=int, default=5000)
    configure = subparsers.add_parser(
        "configure-login", help="使用 Windows DPAPI 加密保存 WebVPN 登录信息"
    )
    configure.add_argument("--profile-root", type=Path, default=Path(".private/course-progress"))
    configure.add_argument("--username", help="学号/工号；省略时交互输入")
    profile = subparsers.add_parser(
        "configure-profile", help="保存用于通知与培养方案匹配的本地学生画像"
    )
    profile.add_argument("--grade", required=True, help="入学年级，例如 2025")
    profile.add_argument("--major", default="", help="专业；未知时可省略")
    profile.add_argument("--academic-level", default="", help="培养层次；未知时可省略")
    profile.add_argument("--campus", default="", help="校区；未知时可省略")
    profile.add_argument("--profile", type=Path, default=DEFAULT_STUDENT_PROFILE_PATH)
    analyze = subparsers.add_parser(
        "analyze-interface", help="只读分析学校页面提供的接口契约"
    )
    analyze.add_argument("--target", choices=("student-profile",), required=True)
    analyze.add_argument("--url", default=DEFAULT_PORTAL_URL, help="WebVPN 门户入口")
    analyze.add_argument("--profile-root", type=Path, default=Path(".private/course-progress"))
    analyze.add_argument("--output-root", type=Path, default=Path(".private/interface-analysis"))
    analyze.add_argument("--login-timeout-seconds", type=int, default=600)
    analyze.add_argument("--wait-seconds", type=int, default=60)
    explore = subparsers.add_parser(
        "explore-entry", help="只读观察已确认通知对应的教务选课入口"
    )
    explore.add_argument("--notice", type=Path, default=Path(".private/academic-selection/selection-notice.json"))
    explore.add_argument("--url", default=DEFAULT_PORTAL_URL, help="教务门户入口")
    explore.add_argument("--browser", choices=("chromium",), default="chromium")
    explore.add_argument("--private-root", type=Path, default=Path(".private/academic-selection"))
    explore.add_argument("--profile-root", type=Path, default=Path(".private/course-progress"))
    explore.add_argument("--login-timeout-seconds", type=int, default=600)
    explore.add_argument("--wait-seconds", type=int, default=600)
    for command, target, help_text in (
        ("discover-timetable", TARGET_TIMETABLE, "自动点击并发现课表只读接口"),
        ("discover-selection", TARGET_SELECTION, "自动点击并发现选课只读接口"),
    ):
        discover = subparsers.add_parser(command, help=help_text)
        discover.set_defaults(discovery_target=target)
        discover.add_argument("--url", default=DEFAULT_PORTAL_URL, help="教务门户入口")
        discover.add_argument("--browser", choices=("chromium",), default="chromium")
        discover.add_argument("--private-root", type=Path, default=Path(".private/academic-selection"))
        discover.add_argument("--profile-root", type=Path, default=Path(".private/course-progress"))
        discover.add_argument("--login-timeout-seconds", type=int, default=600)
        discover.add_argument("--wait-seconds", type=int, default=30)
        discover.add_argument("--max-clicks", type=int, default=8)
        discover.add_argument(
            "--persistent-session",
            action="store_true",
            help="复用持久登录 profile；默认使用 Playwright 临时无痕会话",
        )
        if target == TARGET_SELECTION:
            discover.add_argument(
                "--notice",
                type=Path,
                default=Path(".private/academic-selection/selection-notice.json"),
                help="已确认的选课通知；仅查询通知明确开放的课程类别",
            )
            discover.add_argument(
                "--grade",
                help="临时覆盖学生画像中的入学年级，例如 2025",
            )
            discover.add_argument(
                "--profile",
                type=Path,
                default=DEFAULT_STUDENT_PROFILE_PATH,
                help="本地学生画像",
            )
    return parser


def run_interface_analysis(args: argparse.Namespace) -> int:
    from .student_profile_observation import StudentProfileInterfaceAnalyzer

    analyzer = StudentProfileInterfaceAnalyzer(
        profile_root=args.profile_root.resolve(),
        output_root=args.output_root.resolve(),
        portal_url=args.url,
        login_timeout_seconds=args.login_timeout_seconds,
        wait_seconds=args.wait_seconds,
    )
    output = analyzer.run()
    print(f"候选学生画像接口：{output}")
    return 0


def run_discovery(args: argparse.Namespace) -> int:
    if args.login_timeout_seconds <= 0 or args.wait_seconds <= 0 or args.max_clicks <= 0:
        raise SystemExit("等待时间和自动点击次数必须大于 0")
    allowed_categories: tuple[str, ...] = ()
    allowed_windows = {}
    semester_label = ""
    if args.discovery_target == TARGET_SELECTION:
        if not args.notice.is_file():
            raise SystemExit(f"选课通知不存在：{args.notice}")
        notice = load_notice(args.notice)
        if notice.status != "confirmed":
            raise SystemExit("选课通知尚未确认，禁止查询任何课程类别")
        if args.grade:
            grade = create_student_profile(grade=args.grade).grade
        else:
            if not args.profile.is_file():
                raise SystemExit(
                    "学生画像不存在，请先运行 configure-profile --grade 2025"
                )
            grade = load_student_profile(args.profile).grade
        allowed_categories = notice_selection_categories(notice, grade=grade)
        allowed_windows = notice_selection_window_map(notice, grade=grade)
        semester_label = notice_semester_label(notice)
        if not allowed_categories:
            raise SystemExit("选课通知未明确开放课程类别，禁止猜测查询入口")
        if not semester_label:
            raise SystemExit("选课通知未明确学期，禁止查询")
        labels = "、".join(CATEGORY_LABELS[code] for code in allowed_categories)
        print(f"通知白名单：{semester_label}；{grade}级；{labels}")
    with sync_playwright() as playwright:
        print("安全边界：只自动点击导航/查询控件；疑似选课、退课、保存请求会被拦截。")
        report = AcademicInterfaceDiscovery(
            playwright,
            browser_name=args.browser,
            profile_root=args.profile_root,
            output_root=args.private_root.resolve(),
            persistent_session=args.persistent_session,
        ).discover(
            args.discovery_target,
            portal_url=args.url,
            login_timeout_seconds=args.login_timeout_seconds,
            wait_seconds=args.wait_seconds,
            max_clicks=args.max_clicks,
            allowed_selection_categories=allowed_categories,
            allowed_selection_windows=allowed_windows,
            notice_semester=semester_label,
        )
    print(
        f"发现结果：目标页面={'是' if report.target_found else '否'}，"
        f"点击={report.clicks}，接口={report.captures}，拦截={report.blocked_requests}"
    )
    if report.selection_query_path:
        print(f"课程查询结果：{report.selection_query_path}")
    return 0


def run_configure_login(args: argparse.Namespace) -> int:
    username = (args.username or input("学号/工号：")).strip()
    password = getpass.getpass("统一身份认证密码（输入不会显示）：")
    credential_store(args.profile_root).save(
        LoginCredentials(username=username, password=password)
    )
    print("登录信息已由 Windows DPAPI 加密保存，仅当前 Windows 用户可解密。")
    return 0


def run_configure_profile(args: argparse.Namespace) -> int:
    profile = create_student_profile(
        grade=args.grade,
        major=args.major,
        academic_level=args.academic_level,
        campus=args.campus,
    )
    save_student_profile(args.profile, profile)
    print(f"学生画像已保存：{profile.grade}级；{args.profile}")
    return 0


def run_explore_entry(args: argparse.Namespace) -> int:
    if args.login_timeout_seconds <= 0 or args.wait_seconds <= 0:
        raise SystemExit("等待时间必须大于 0")
    if not args.notice.is_file():
        raise SystemExit(f"选课通知不存在：{args.notice}")
    notice = load_notice(args.notice)
    if notice.status != "confirmed":
        raise SystemExit("选课通知尚未确认，不能探索对应入口")
    profile_dir = resolve_profile_dir(args.profile_root.resolve())
    with sync_playwright() as playwright:
        context = launch_browser_context(playwright, args.browser, profile_dir)
        try:
            page = context.pages[0] if context.pages else context.new_page()
            try:
                page.goto(args.url, wait_until="domcontentloaded", timeout=60_000)
            except Exception as error:
                save_selection_result(
                    args.private_root.resolve(),
                    SelectionObservation(
                        status=STATUS_ENTRY_UNREACHABLE,
                        request_url=args.url,
                        method="GET",
                        message=str(error),
                        sections=(),
                    ),
                )
                return 1
            deadline = time.monotonic() + args.login_timeout_seconds
            while _is_login_url(page.url) and time.monotonic() < deadline:
                page.wait_for_timeout(500)
            if _is_login_url(page.url):
                save_selection_result(
                    args.private_root.resolve(),
                    SelectionObservation(
                        status=STATUS_LOGIN_REQUIRED,
                        request_url=page.url,
                        method="GET",
                        message="等待统一身份认证超时",
                        sections=(),
                    ),
                )
                return 1
            explorer = SelectionEntryExplorer(
                notice=notice,
                output_root=args.private_root.resolve(),
            )
            explorer.run(context, wait_seconds=args.wait_seconds)
        finally:
            context.close()
    print(f"只读选课入口结果：{args.private_root.resolve() / 'selection-entry.json'}")
    return 0


def main(argv: list[str] | None = None) -> int:
    args = build_parser().parse_args(argv)
    if args.command == "workbench":
        from .application import run_workbench_application
        return run_workbench_application(args.private_root, args.port)
    if args.command == "configure-login":
        return run_configure_login(args)
    if args.command == "configure-profile":
        return run_configure_profile(args)
    if args.command == "analyze-interface":
        return run_interface_analysis(args)
    if args.command == "explore-entry":
        return run_explore_entry(args)
    if args.command in {"discover-timetable", "discover-selection"}:
        return run_discovery(args)
    return 2
