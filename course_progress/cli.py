"""Command-line entry point for course progress tooling."""

from __future__ import annotations

import argparse
import json
from dataclasses import asdict
from datetime import datetime, timezone
from pathlib import Path

from playwright.sync_api import sync_playwright

from .academic_client import AuthenticatedAcademicClient
from .collector import (
    CollectionCheckpoint,
    FixedGradeReader,
    load_checkpoint,
    save_checkpoint,
)
from .explorer import (
    DEFAULT_PORTAL_URL,
    PortalExplorer,
    make_capture_session,
    resolve_profile_dir,
)
from .progress import RequirementBaseline, evaluate_progress, parse_requirements
from .session import AcademicBrowserSession

DEFAULT_PRIVATE_ROOT = Path(".private/course-progress")
DEFAULT_REQUIREMENTS = Path("docs/校园培养方案解读（2026年版）.md")
GUIDE_2026_CATEGORY_MAPPING = {
    "本专业选修": "major_elective",
    "外专业选修": "outside_major_elective",
    "文理通识-文化素质教育课": "cultural_quality",
    "创新研修课": "innovation",
    "社会实践": "social_practice",
}


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        prog="course-progress",
        description="探索教务系统课程与培养方案接口（只读捕获阶段）",
    )
    subparsers = parser.add_subparsers(dest="command", required=True)
    explore = subparsers.add_parser(
        "explore", help="打开可视化浏览器并捕获 Fetch/XHR JSON 响应"
    )
    explore.add_argument("--url", default=DEFAULT_PORTAL_URL, help="教务系统入口 URL")
    explore.add_argument(
        "--browser",
        choices=("chromium",),
        default="chromium",
        help="固定使用可见的 Playwright Chromium",
    )
    explore.add_argument(
        "--private-root",
        type=Path,
        default=DEFAULT_PRIVATE_ROOT,
        help="浏览器 profile 与原始捕获的私有目录",
    )
    explore.add_argument(
        "--max-response-mb",
        type=float,
        default=5.0,
        help="单个 JSON 响应保存上限，默认 5 MiB",
    )
    explore.add_argument(
        "--max-pages",
        type=int,
        default=12,
        help="自动导航最多访问的页面数，默认 12",
    )
    explore.add_argument(
        "--login-timeout-seconds",
        type=int,
        default=600,
        help="等待 CAS 登录和重新认证的最长时间，默认 600 秒",
    )
    collect = subparsers.add_parser(
        "collect", help="登录后按学期和分页采集已通过课程并计算毕业进度"
    )
    collect.add_argument("--url", default=DEFAULT_PORTAL_URL, help="教务系统入口 URL")
    collect.add_argument(
        "--browser", choices=("chromium",), default="chromium"
    )
    collect.add_argument(
        "--private-root", type=Path, default=DEFAULT_PRIVATE_ROOT
    )
    collect.add_argument(
        "--requirements", type=Path, default=DEFAULT_REQUIREMENTS
    )
    collect.add_argument("--baseline-version", default="guide-2026")
    collect.add_argument("--page-size", type=int, default=20)
    collect.add_argument(
        "--login-timeout-seconds",
        type=int,
        default=600,
        help="等待 CAS 登录和重新认证的最长时间，默认 600 秒",
    )
    return parser


def run_explore(args: argparse.Namespace) -> int:
    if args.max_response_mb <= 0:
        raise SystemExit("--max-response-mb 必须大于 0")
    if args.max_pages <= 0:
        raise SystemExit("--max-pages 必须大于 0")
    if args.login_timeout_seconds <= 0:
        raise SystemExit("--login-timeout-seconds 必须大于 0")

    private_root = args.private_root.resolve()
    profile_dir = resolve_profile_dir(private_root)
    capture_dir = make_capture_session(private_root / "captures")

    with sync_playwright() as playwright:
        explorer = PortalExplorer(
            playwright=playwright,
            browser_name=args.browser,
            profile_dir=profile_dir,
            capture_dir=capture_dir,
            max_response_bytes=int(args.max_response_mb * 1024 * 1024),
        )
        explorer.run(
            args.url,
            max_pages=args.max_pages,
            login_timeout_seconds=args.login_timeout_seconds,
        )
    return 0


def _progress_report_data(collection, report) -> dict:
    return {
        "generated_at": datetime.now(timezone.utc).isoformat(timespec="seconds"),
        "baseline_version": report.baseline_version,
        "data_complete": collection.complete,
        "semesters": [asdict(item) for item in collection.semesters],
        "collection_failures": [asdict(item) for item in collection.failures],
        "progress": [
            {
                "key": item.requirement.key,
                "label": item.requirement.label,
                "required_credits": item.requirement.minimum_credits,
                "completed_credits": item.completed_credits,
                "remaining_credits": item.remaining_credits,
                "courses": [asdict(course) for course in item.courses],
            }
            for item in report.progress
        ],
        "conflicts": [asdict(item) for item in report.conflicts],
        "unclassified_courses": [
            asdict(course) for course in report.unclassified_courses
        ],
    }


def run_collect(args: argparse.Namespace) -> int:
    if args.page_size <= 0:
        raise SystemExit("--page-size 必须大于 0")
    if args.login_timeout_seconds <= 0:
        raise SystemExit("--login-timeout-seconds 必须大于 0")
    if not args.requirements.is_file():
        raise SystemExit(f"要求基线不存在：{args.requirements}")

    private_root = args.private_root.resolve()
    output_path = private_root / "progress-report.json"

    with sync_playwright() as playwright:
        with AcademicBrowserSession(
            playwright, browser_name=args.browser, profile_root=private_root
        ) as session:
            page = session.open_portal_application(
                args.url,
                "新教务系统",
                timeout_seconds=args.login_timeout_seconds,
            )
            print("正在检查已有会话；会话失效时请在浏览器中完成统一身份认证。")
            checkpoint_path = private_root / "collection-checkpoint.json"
            checkpoint = load_checkpoint(checkpoint_path)
            checkpoint_state = checkpoint

            def save_page_checkpoint(
                semester,
                page_number: int,
                page_count: int,
                page_records,
            ) -> None:
                nonlocal checkpoint_state
                checkpoint_state = CollectionCheckpoint(
                    records=checkpoint_state.records + tuple(page_records),
                    completed_pages=tuple(
                        sorted(
                            set(checkpoint_state.completed_pages)
                            | {(semester.value, page_number)}
                        )
                    ),
                    page_counts=tuple(
                        sorted(
                            (
                                dict(checkpoint_state.page_counts)
                                | {semester.value: page_count}
                            ).items()
                        )
                    ),
                )
                save_checkpoint(checkpoint_path, checkpoint_state)

            client = AuthenticatedAcademicClient(
                page,
                authenticate=lambda url, target: session.open_authenticated(
                    url, timeout_seconds=args.login_timeout_seconds, page=target
                ),
                timeout_seconds=min(15, args.login_timeout_seconds),
            )
            collection = FixedGradeReader(
                client, page_size=args.page_size
            ).collect(
                on_page_data=save_page_checkpoint,
                checkpoint=checkpoint,
            )

    baseline = RequirementBaseline(
        version=args.baseline_version,
        requirements=parse_requirements(args.requirements),
        category_mapping=GUIDE_2026_CATEGORY_MAPPING,
    )
    report = evaluate_progress(collection.records, baseline)
    output_path.write_text(
        json.dumps(
            _progress_report_data(collection, report),
            ensure_ascii=False,
            indent=2,
        ),
        encoding="utf-8",
    )
    if collection.complete:
        (private_root / "collection-checkpoint.json").unlink(missing_ok=True)

    print(f"学期数：{len(collection.semesters)}")
    print(f"课程记录：{len(collection.records)}")
    print(f"数据完整：{'是' if collection.complete else '否'}")
    for item in report.progress:
        print(
            f"{item.requirement.label}: "
            f"{item.completed_credits:g}/{item.requirement.minimum_credits:g}，"
            f"剩余 {item.remaining_credits:g} 学分"
        )
    if report.unclassified_courses:
        print(f"待归类课程：{len(report.unclassified_courses)}")
    if report.conflicts:
        print(f"课程冲突：{len(report.conflicts)}")
    print(f"私有报告：{output_path}")
    return 0


def main(argv: list[str] | None = None) -> int:
    parser = build_parser()
    args = parser.parse_args(argv)
    if args.command == "explore":
        return run_explore(args)
    if args.command == "collect":
        return run_collect(args)
    parser.error(f"未知命令：{args.command}")
    return 2
