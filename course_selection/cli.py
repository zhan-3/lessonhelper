"""Command line entry points for local academic selection exploration."""

from __future__ import annotations

import argparse
import time
from pathlib import Path

from playwright.sync_api import sync_playwright

from course_progress.explorer import (
    DEFAULT_PORTAL_URL,
    _is_login_url,
    launch_browser_context,
    resolve_profile_dir,
)

from .notice import load_notice
from .discovery import AcademicInterfaceDiscovery, TARGET_SELECTION, TARGET_TIMETABLE
from .selection_entry import (
    STATUS_ENTRY_UNREACHABLE,
    STATUS_LOGIN_REQUIRED,
    SelectionEntryExplorer,
    SelectionObservation,
    save_selection_result,
)


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(prog="course-selection")
    subparsers = parser.add_subparsers(dest="command", required=True)
    explore = subparsers.add_parser(
        "explore-entry", help="只读观察已确认通知对应的教务选课入口"
    )
    explore.add_argument("--notice", type=Path, default=Path(".private/academic-selection/selection-notice.json"))
    explore.add_argument("--url", default=DEFAULT_PORTAL_URL, help="教务门户入口")
    explore.add_argument("--browser", choices=("chromium", "chrome"), default="chromium")
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
        discover.add_argument("--browser", choices=("chromium", "chrome"), default="chromium")
        discover.add_argument("--private-root", type=Path, default=Path(".private/academic-selection"))
        discover.add_argument("--profile-root", type=Path, default=Path(".private/course-progress"))
        discover.add_argument("--login-timeout-seconds", type=int, default=600)
        discover.add_argument("--wait-seconds", type=int, default=30)
        discover.add_argument("--max-clicks", type=int, default=8)
    return parser


def run_discovery(args: argparse.Namespace) -> int:
    if args.login_timeout_seconds <= 0 or args.wait_seconds <= 0 or args.max_clicks <= 0:
        raise SystemExit("等待时间和自动点击次数必须大于 0")
    with sync_playwright() as playwright:
        print("安全边界：只自动点击导航/查询控件；疑似选课、退课、保存请求会被拦截。")
        report = AcademicInterfaceDiscovery(
            playwright,
            browser_name=args.browser,
            profile_root=args.profile_root,
            output_root=args.private_root.resolve(),
        ).discover(
            args.discovery_target,
            portal_url=args.url,
            login_timeout_seconds=args.login_timeout_seconds,
            wait_seconds=args.wait_seconds,
            max_clicks=args.max_clicks,
        )
    print(
        f"发现结果：目标页面={'是' if report.target_found else '否'}，"
        f"点击={report.clicks}，接口={report.captures}，拦截={report.blocked_requests}"
    )
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
    if args.command == "explore-entry":
        return run_explore_entry(args)
    if args.command in {"discover-timetable", "discover-selection"}:
        return run_discovery(args)
    return 2
