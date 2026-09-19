"""Command line entry points for local academic selection exploration."""

from __future__ import annotations

import getpass
import json
import logging
import time
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
from .categories import CATEGORY_LABELS
from .discovery import TARGET_SELECTION, TARGET_TIMETABLE, AcademicInterfaceDiscovery
from .notice import (
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
    """选课工作台 CLI — 选课规划、毕业进度、实验预约与接口契约检查。"""
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


@main.command("lab-booking")
@click.option("--center", default="dxwl", help="实验教学中心代码，例如 dwxl/dgdz/jskx")
@click.option("--cdp", default="http://127.0.0.1:9222", help="已登录标签页的本地调试端口")
@click.option("--private-root", type=click.Path(path_type=Path), default=config.WORKBENCH_PRIVATE_ROOT)
@click.option("--probe-cap", type=int, default=8, help="每个实验最多探测的空位数量")
@click.option("--confirm", default="", help="规划令牌；只有匹配时才提交")
@click.option("--plan-out", type=click.Path(path_type=Path), default=None, help="把规划写入本地 JSON")
@click.option("--avoid", type=click.Choice(["none", "timetable", "labs", "all"]), default="timetable",
              help="避让来源：不避让 / 只避课表 / 只避已约实验 / 两者都避")
@click.option("--only", "only_ids", multiple=True, type=int, help="只规划指定 subjectId，可重复")
@click.option("--from-plan", type=click.Path(path_type=Path), default=None,
              help="提交时从该规划文件读取目标（与 --confirm 配套，避免重新规划）")
def lab_booking_cmd(
    center: str, cdp: str, private_root: Path, probe_cap: int, confirm: str,
    plan_out: Path | None, avoid: str, only_ids: tuple[int, ...], from_plan: Path | None,
) -> None:
    """实验预约规划与单次提交（实验状态，未通过真实环境验收）。"""
    from .lab_booking import (
        LabSlot,
        OnlySubjects,
        book_slots,
        busy_from_bookings,
        busy_intervals,
        load_workspace_timetable,
        plan_lab_slots,
        plan_to_json,
        plan_token,
    )
    from .lab_browser_session import BrowserLabSession

    entries = load_workspace_timetable(private_root) if private_root.is_dir() else ()
    scheduled = busy_intervals(entries)
    click.echo(f"课表区间 {len(scheduled)} 条（来自本地工作台快照）")

    live = BrowserLabSession.attach(cdp, center)
    target = OnlySubjects(live, only_ids) if only_ids else live
    try:
        # Lab sessions live in a different system: treat the ones already booked
        # in this center as occupied too, so planning cannot double-book a slot.
        booked = busy_from_bookings(live.booked())
        if avoid == "none":
            busy = ()
            click.echo("避让已关闭（--avoid none）：只按空位规划")
        elif avoid == "timetable":
            busy = scheduled
        elif avoid == "labs":
            busy = booked
            click.echo(f"只避已约实验 {len(booked)} 条（忽略课表）")
        else:
            busy = (*scheduled, *booked)
            if booked:
                click.echo(f"已约实验占用 {len(booked)} 条（来自 openlab）")
        if only_ids:
            click.echo(f"只规划 {', '.join(str(value) for value in only_ids)}")
        result = plan_lab_slots(target, busy, probe_cap=probe_cap)
        for slot in result.slots:
            click.echo(
                f"  {slot.subject_id} {slot.subject_name[:18]:18} {slot.class_date} "
                f"第{slot.week}周 {slot.timer_name} {slot.start} {slot.lab} 座位{slot.table_no}"
            )
        for left, right in result.reminders:
            click.echo(
                f"  ⚠ 同一时刻冲突（需人工取消一个）: {left.class_date} {left.timer_name} "
                f"{left.subject_name} × {right.subject_name}"
            )
        for name in result.unresolved:
            click.echo(f"  ✗ 暂无可约时段: {name}")
        token = plan_token(result.slots)
        click.echo(f"可约 {len(result.slots)} 门；规划令牌 {token}")
        if plan_out is not None:
            plan_out.write_text(plan_to_json(result), encoding="utf-8")
            click.echo(f"规划已写入 {plan_out}")

        if not confirm:
            click.echo("未提供 --confirm，仅做只读规划。")
            return
        if from_plan is None:
            raise click.ClickException("提交必须用 --from-plan 指定保存的目标，避免提交与确认不一致。")
        saved = json.loads(from_plan.read_text(encoding="utf-8"))
        targets = [LabSlot.from_dict(item) for item in saved.get("slots") or []]
        if confirm != plan_token(targets):
            raise click.ClickException("确认令牌与规划文件不一致，拒绝提交。")
        for outcome in book_slots(target, targets, confirmation=confirm):
            click.echo(f"  {outcome['subject_id']} {outcome['class_date']} {outcome['start']} -> "
                       f"{outcome['outcome']} {outcome.get('detail') or ''}")
            if outcome["outcome"] != "confirmed_success":
                click.echo("存在需人工核验的结果，已停止提交。")
                break
    finally:
        live.close()


# ── lab-contract ─────────────────────────────────────────────────────────────


def _lab_contract_transport(center: str, cdp: str, kind: str, origin: str):
    """Build the transport for one center, by profile, by CDP borrow, or plain HTTP."""
    from .lab_transport import BrowserLabTransport, HttpLabTransport, acquire_token

    if kind == "profile":
        # Launch the project's own persistent Chromium.  Its profile already
        # carries the openlab sign-in, so neither a manually started browser
        # nor a CDP port is needed.  ``as_transport`` hands over the Playwright
        # instance, and the caller's ``transport.close()`` releases it.
        from .lab_browser_session import BrowserLabSession

        session = BrowserLabSession.from_profile(config.PROGRESS_PROFILE_ROOT, center)
        return session.as_transport()

    if kind == "browser":
        from playwright.sync_api import sync_playwright

        playwright = sync_playwright().start()
        browser = playwright.chromium.connect_over_cdp(cdp)
        page = next((item for item in browser.contexts[0].pages if f"/{center}/" in item.url), None)
        if page is None:
            playwright.stop()
            raise click.ClickException(f"没有 {center} 的已打开标签页")
        from .lab_transport import install_token_hook, page_token

        install_token_hook(page)
        token = page_token(page)
        if not token:
            playwright.stop()
            raise click.ClickException("未读到应用令牌；请先在页面里操作一次")
        transport = BrowserLabTransport(page, center, token)
        transport._playwright = playwright  # keeps the borrowed page usable until close()
        return transport
    token, _url = acquire_token(cdp, center)
    return HttpLabTransport(center, token, origin=origin)


@main.command("lab-contract")
@click.argument("action", type=click.Choice(["record", "check", "promote"]))
@click.option("--center", default="dxwl")
@click.option("--channel", default="direct", help="通道名；baseline 按通道分开存")
@click.option("--origin", default="http://openlab.hitwh.edu.cn")
@click.option("--cdp", default="http://127.0.0.1:9222")
@click.option("--transport", "kind", type=click.Choice(["http", "browser", "profile"]), default="http")
@click.option("--observations-dir", type=click.Path(path_type=Path), default=Path(".private/lab-contracts"))
@click.option("--baselines-dir", type=click.Path(path_type=Path), default=Path("docs/contracts"))
@click.option("--env", "env_pairs", multiple=True, help="额外环境项 KEY=VALUE，只写入本机观测")
@click.option("--json", "as_json", is_flag=True, help="输出机器可读 JSON")
def lab_contract_cmd(
    action: str, center: str, channel: str, origin: str, cdp: str, kind: str,
    observations_dir: Path, baselines_dir: Path, env_pairs: tuple[str, ...], as_json: bool,
) -> None:
    """记录、比对或提升一个中心的接口契约（只读，不调用写入端点）。"""
    from .lab_contract import diff, observe, without_environment

    name = f"{channel}-{center}"
    baseline_path = baselines_dir / f"{name}.json"
    observation_path = observations_dir / f"{name}.observed.json"

    if action == "promote":
        if not observation_path.is_file():
            raise click.ClickException(f"没有观测文件 {observation_path}；先跑 record")
        from .lab_contract import ContractSnapshot, EndpointContract

        observed = ContractSnapshot.from_json(observation_path.read_text(encoding="utf-8"))
        locked = {}
        if baseline_path.is_file():
            previous = ContractSnapshot.from_json(baseline_path.read_text(encoding="utf-8"))
            locked = {endpoint: item.locked for endpoint, item in previous.endpoints.items()}
        promoted = without_environment(observed)
        promoted = ContractSnapshot(
            center=promoted.center, channel=promoted.channel, origin=promoted.origin,
            app_version=promoted.app_version, header_required=promoted.header_required,
            static_assets=promoted.static_assets, static_config=promoted.static_config,
            environment={},
            endpoints={
                endpoint: EndpointContract(
                    endpoint=item.endpoint, request_fields=item.request_fields,
                    response_fields=item.response_fields,
                    locked=tuple(locked.get(endpoint, item.locked)),
                    envelope=item.envelope, codes=item.codes,
                )
                for endpoint, item in promoted.endpoints.items()
            },
            unavailable=promoted.unavailable,
        )
        baseline_path.parent.mkdir(parents=True, exist_ok=True)
        baseline_path.write_text(promoted.to_json(), encoding="utf-8")
        click.echo(f"基线已更新 {baseline_path}（保留已有 locked 字段，环境块不入仓）")
        return

    environment = {"channel": channel, "transport": kind}
    for pair in env_pairs:
        key, _, value = pair.partition("=")
        environment[key.strip()] = value.strip()

    transport = _lab_contract_transport(center, cdp, kind, origin)
    try:
        observation = observe(transport, origin=origin, channel=channel, environment=environment)
    finally:
        transport.close()

    observation_path.parent.mkdir(parents=True, exist_ok=True)
    observation_path.write_text(observation.snapshot.to_json(), encoding="utf-8")

    if observation.unavailable:
        click.echo(f"读不到的端点: {', '.join(observation.unavailable)}")

    baseline = None
    if baseline_path.is_file():
        from .lab_contract import ContractSnapshot

        baseline = ContractSnapshot.from_json(baseline_path.read_text(encoding="utf-8"))
    report = diff(baseline, observation.snapshot) if baseline else None

    if action == "record":
        click.echo(f"观测已写入 {observation_path}（{len(observation.snapshot.endpoints)} 个端点）")
        if report is None:
            click.echo(f"尚无基线 {baseline_path}；确认无误后跑 promote 建立第一份")
            return
        click.echo("与基线的差异：")
        for line in report.describe():
            click.echo(line)
        return

    if as_json:
        payload = {"center": center, "channel": channel,
                   "observation": str(observation_path),
                   "baseline": str(baseline_path) if baseline else None,
                   "unavailable": observation.unavailable,
                   "report": report.to_dict() if report else None}
        click.echo(json.dumps(payload, ensure_ascii=False, indent=2))
    if baseline is None:
        click.echo(f"无法比对：基线 {baseline_path} 不存在。先跑 record 再 promote。")
        raise SystemExit(2)
    if not as_json:
        click.echo(f"基线 {baseline_path}")
        for line in report.describe():
            click.echo(line)
    if report.blocks:
        raise SystemExit(1)


# ── lab-exam ────────────────────────────────────────────────────────────────


@main.command("lab-exam")
@click.option("--center", default="dxwl", help="教学中心代码")
@click.option("--subject-id", type=int, default=None, help="考核科目 ID；用 --list-subjects 查看")
@click.option("--list-subjects", "list_subjects", is_flag=True, help="列出可选考核科目及其状态后退出")
@click.option("--cdp", default="http://127.0.0.1:9222", help="已登录浏览器的 CDP 端点（仅 --transport browser）")
@click.option("--transport", "kind", type=click.Choice(["profile", "browser", "http"]), default="profile",
              help="profile=项目持久化 Chromium（推荐，无需手动开浏览器）；browser=借用 CDP 标签页")
@click.option("--origin", default="http://openlab.hitwh.edu.cn")
@click.option("--answers", type=click.Path(exists=True, path_type=Path), default=None,
              help="答案文件（JSON：{题目ID: [选项...]}）；不提供则只读")
@click.option("--confirm", default="", help="确认令牌；不提供则只做干跑校验")
@click.option("--plain", is_flag=True, help="只输出题干与选项的纯文本块，便于复制到别处检索")
@click.option("--json", "as_json", is_flag=True, help="输出机器可读 JSON（含 questionId，供答案文件使用）")
def lab_exam_cmd(center: str, subject_id: int | None, list_subjects: bool, cdp: str, kind: str, origin: str,
                 answers: Path | None, confirm: str, plain: bool, as_json: bool) -> None:
    """实验预考核：读取状态与题目（只读）；提交需显式确认，单次且不重试。

    答案由使用者提供。本命令不读取、不推测服务端随题目下发的答案字段。
    """
    from .lab_exam import (
        GROUP_SINGLE,
        exam_token,
        parse_answer_mapping,
        submit_exam,
        validate_answers,
    )
    from .lab_exam_transport import TransportLabExam

    try:
        transport = _lab_contract_transport(center, cdp, kind, origin)
    except click.ClickException:
        raise
    except Exception as error:                      # Playwright 异常类型随版本变化
        raise click.ClickException(
            f"无法连接 {cdp}：{error}\n"
            "请先用 --remote-debugging-port=9222 启动浏览器、登录 openlab，"
            "并打开目标中心的页面后再执行。"
        ) from error

    try:
        exam = TransportLabExam(transport)

        if list_subjects:
            payload = transport.call("view/subjects")
            rows = payload.get("result") if payload.get("code") == 0 else None
            if not isinstance(rows, list):
                raise click.ClickException(
                    f"读取科目失败：{payload.get('code')} {payload.get('message')}"
                )
            for row in rows:
                name = str(row.get("subjectName") or "")
                passed = "已通过" if row.get("izPass") else "未通过"
                click.echo(f"  {row.get('subjectId'):>6}  {name:22s} {row.get('claim') or '':6s} {passed}")
            return

        if subject_id is None:
            raise click.ClickException("需要 --subject-id（或用 --list-subjects 查看可选科目）")

        status = exam.exam_status(subject_id)
        sheet = exam.exam_sheet(subject_id)

        if plain and answers is None:
            # 每题一段，题干 + 选项，不带序号与题型标注，便于整块复制检索。
            # 要 questionId 请用 --json。
            for question in sheet.questions:
                click.echo(question.text)
                for option_label, option_text in question.options:
                    click.echo(f"{option_label}. {option_text}")
                click.echo()
            return

        if as_json and answers is None:
            click.echo(json.dumps(
                {"status": status.to_dict(), "sheet": sheet.to_dict()},
                ensure_ascii=False, indent=2,
            ))
            return

        if not as_json and not plain:
            if status.allowed:
                detail = f"（{status.message}）" if status.message else ""
                click.echo(f"状态：可参加{detail}")
            else:
                click.echo(f"状态：不可参加（{status.message_code} {status.message}）")
            click.echo(f"科目 {sheet.subject_id} {sheet.subject_name}，共 {len(sheet.questions)} 题")
            for index, question in enumerate(sheet.questions, start=1):
                label = "单选" if question.group == GROUP_SINGLE else "多选"
                click.echo(f"\n[{index}/{len(sheet.questions)}] ({label}) {question.text}")
                for option_label, option_text in question.options:
                    click.echo(f"    {option_label}. {option_text}")

        if answers is None:
            click.echo("\n（未提供 --answers：仅读取，未提交）")
            return

        try:
            raw = json.loads(answers.read_text(encoding="utf-8"))
        except (OSError, json.JSONDecodeError) as error:
            raise click.ClickException(f"答案文件无法读取：{error}") from error
        if not isinstance(raw, dict):
            raise click.ClickException('答案文件必须是 JSON 对象，形如 {"<题目ID>": ["A"]}')
        try:
            targets = parse_answer_mapping(raw, sheet)
        except ValueError as error:
            raise click.ClickException(str(error)) from error

        token = exam_token(subject_id, targets)
        if not confirm:
            problems = validate_answers(sheet, targets)
            click.echo(f"\n干跑完成，未提交。确认令牌：{token}")
            if problems:
                click.echo("答案校验未通过：")
                for problem in problems:
                    click.echo(f"  - {problem}")
            else:
                click.echo("答案校验通过。")
                click.echo(f"确认后重新执行并加 --confirm {token} 提交（仅一次，不会重试）。")
            return

        outcome = submit_exam(exam, subject_id, targets, confirmation=confirm)
    finally:
        transport.close()

    if as_json:
        click.echo(json.dumps(outcome, ensure_ascii=False, indent=2))
    else:
        suffix = f" — {outcome['detail']}" if outcome.get("detail") else ""
        click.echo(f"\n结果：{outcome['outcome']}{suffix}")
        if outcome.get("verdict"):
            click.echo(f"判定：{outcome['verdict']}")
    if outcome["outcome"] != "confirmed_success":
        raise SystemExit(1)


if __name__ == "__main__":
    main()
