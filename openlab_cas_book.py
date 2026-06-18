"""
openlab_cas_book.py

Playwright 自动化：CAS SSO 登录 → 开放式实验系统 → 自动预约

流程：
  首次运行：自动打开浏览器 → 跳转到 CAS 登录页 → 用户手动输账号密码/验证码 → 登录成功 → 保存会话
  后续运行：加载已保存会话 → 直接操作预约页面
  支持 --monitor 监控模式：每隔 N 秒检查，有名额立即抢

用法：
  python openlab_cas_book.py                          # 单次尝试所有课程
  python openlab_cas_book.py --monitor                 # 监控模式
  python openlab_cas_book.py --monitor --interval=3    # 每 3 秒查一次
  python openlab_cas_book.py --course "课程名"         # 指定单个课程
  python openlab_cas_book.py --login-only              # 仅登录保存会话，不预约
"""

import os
import sys
import time
import random
from datetime import datetime
from pathlib import Path

from playwright.sync_api import sync_playwright, TimeoutError as PwTimeout

BASE_URL = "http://openlab.hitwh.edu.cn"
BOOKING_URL = f"{BASE_URL}/dxwl/booking/#/booking"

STORAGE_FILE = Path("storage_state.json")
LOGIN_TIMEOUT_SEC = 180
POLL_INTERVAL = 5
MAX_POLL_MINUTES = 120

# 随机 UA 池，防检测
_USER_AGENTS = [
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/125.0.0.0 Safari/537.36",
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/126.0.0.0 Safari/537.36",
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/127.0.0.0 Safari/537.36",
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/128.0.0.0 Safari/537.36",
]

# 请求间隔抖动范围（秒）
_JITTER_MIN = 0.3
_JITTER_MAX = 1.5


def jitter():
    """随机延迟，模拟人类操作节奏"""
    time.sleep(random.uniform(_JITTER_MIN, _JITTER_MAX))

TARGET_COURSES = [
    "DIY电磁混合磁悬浮",
    "磁阻效应",
    "表面张力",
    "偏振光",
]

def get_chrome_path():
    paths = [
        r"C:\Program Files\Google\Chrome\Application\chrome.exe",
        r"C:\Program Files (x86)\Google\Chrome\Application\chrome.exe",
    ]
    for p in paths:
        if os.path.exists(p):
            return p
    return None


def log(msg: str):
    print(f"[{datetime.now().strftime('%H:%M:%S')}] {msg}")


def safe_click(page, locator, timeout=5000):
    try:
        locator.wait_for(timeout=timeout)
        jitter()
        locator.click()
        return True
    except Exception:
        return False


def try_close_dialog(page):
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


def is_logged_in(page) -> bool:
    try:
        page.goto(BOOKING_URL, timeout=15000)
        page.wait_for_timeout(3000)
        return "login" not in page.url.lower()
    except Exception:
        return False


def do_cas_login(page) -> bool:
    log("请在打开的浏览器中手动登录")
    log("  步骤: 点击「统一身份认证登录」→ 输学号/密码/验证码 → 登录")
    log(f"  等待最长 {LOGIN_TIMEOUT_SEC} 秒...")

    page.goto(BASE_URL, timeout=15000)
    page.wait_for_timeout(2000)

    try:
        cas_btn = page.get_by_text("统一身份认证登录").first
        if cas_btn.is_visible():
            cas_btn.click()
            log("已点击「统一身份认证登录」，页面将跳转到 HIT CAS")
        else:
            log("CAS 登录按钮不可见，可能已自动跳转")
    except Exception:
        log("按钮未找到，可能已跳转到 CAS 页面")

    start = time.time()
    while time.time() - start < LOGIN_TIMEOUT_SEC:
        time.sleep(2)
        current_url = page.url
        if "ids.hit.edu.cn" not in current_url:
            log(f"检测到登录完成！当前 URL: {current_url[:80]}")
            page.wait_for_timeout(3000)
            return True

    log("登录超时")
    return False


def ensure_logged_in(context) -> bool:
    if STORAGE_FILE.exists():
        log(f"发现已保存的会话文件: {STORAGE_FILE}")
        page = context.new_page()

        if is_logged_in(page):
            log("会话有效 ✓")
            page.close()
            return True

        log("会话已过期，需要重新登录")
        page.close()

    page = context.new_page()

    if not do_cas_login(page):
        page.close()
        return False

    if not is_logged_in(page):
        log("登录后仍无法访问预约页面")
        page.close()
        return False

    context.storage_state(path=str(STORAGE_FILE))
    log(f"会话已保存到 {STORAGE_FILE}")
    page.close()
    return True


def do_booking(page, course_name: str) -> bool:
    log(f"▶ 尝试预约: {course_name}")

    try:
        page.goto(BOOKING_URL, timeout=20000)
        page.wait_for_timeout(3000)
    except Exception:
        log("  预约页加载失败")
        page.close()
        return False

    if not safe_click(page, page.locator("button").filter(has_text="选课").first):
        book_link = page.get_by_text("选课").first
        if book_link.count() == 0:
            log("  找不到「选课」按钮")
            page.close()
            return False
        book_link.click()
    page.wait_for_timeout(2000)

    try:
        course_cell = page.locator("td").filter(has_text=course_name).first
        if course_cell.count() == 0:
            log(f"  未找到课程「{course_name}」")
            _go_back(page)
            page.close()
            return False
        course_cell.click()
        page.wait_for_timeout(2000)
    except Exception as e:
        log(f"  选择课程异常: {e}")
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
        log("  未找到可选日期")
        _go_back(page)
        page.close()
        return False

    for date in dates:
        log(f"  尝试日期: {date}")
        try:
            date_cell = page.locator("td").filter(has_text=date).first
            if date_cell.count() == 0:
                continue
            date_cell.click()
            jitter()

            query_btn = page.locator("button").filter(has_text="查询").first
            if query_btn.count() > 0:
                query_btn.click()
                page.wait_for_timeout(2000)

            try:
                dialog = page.locator(".ant-modal-content").first
                if dialog.is_visible(timeout=3000):
                    content = dialog.text_content() or ""
                    if any(w in content for w in ["没有可供选择", "已满", "人数已满"]):
                        log(f"      ✗ {course_name} {date} 已满")
                        try_close_dialog(page)
                        continue
                    else:
                        log(f"      弹窗: {content[:60]}")
            except Exception:
                pass

            seats = page.locator("input[type='radio']")
            if seats.count() == 0:
                log(f"      {date} 无可用座位")
                continue

            seats.first.click()
            jitter()

            confirm_btn = page.locator("button").filter(has_text="确认").first
            if confirm_btn.count() == 0:
                confirm_btn = page.get_by_text("确认").first
            if confirm_btn.count() > 0:
                confirm_btn.click()
                page.wait_for_timeout(3000)
                log(f"      ✓✓✓ 预约成功！{course_name} @ {date}")
                page.close()
                return True

            log("  找不到确认按钮")

        except Exception as e:
            log(f"  日期处理异常: {e}")
            continue

    _go_back(page)
    page.close()
    return False


def _go_back(page):
    try:
        back_btn = page.locator("button").filter(has_text="返回").first
        if back_btn.count() > 0:
            back_btn.click()
            page.wait_for_timeout(1000)
    except Exception:
        pass


def _build_context(browser):
    """Create browser context with anti-detection measures."""
    ua = random.choice(_USER_AGENTS)
    context = browser.new_context(
        viewport={"width": 1400, "height": 900},
        user_agent=ua,
    )
    # 屏蔽不必要的资源加载，提速并减少特征
    context.route("**/*", lambda route: (
        route.abort()
        if route.request.resource_type in ("image", "font", "media")
        else route.continue_()
    ))
    return context


def main():
    print()
    print("=" * 60)
    print("  开放式实验系统 - 自动预约脚本 (CAS SSO)")
    print(f"  {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    print("=" * 60)

    args = set(sys.argv[1:])
    courses = TARGET_COURSES

    for arg in sys.argv:
        if arg.startswith("--course="):
            courses = [arg.split("=", 1)[1]]
            break

    login_only = "--login-only" in args
    monitor_mode = "--monitor" in args
    poll_interval = POLL_INTERVAL
    for arg in sys.argv:
        if arg.startswith("--interval="):
            poll_interval = int(arg.split("=")[1])

    chrome_path = get_chrome_path()
    if chrome_path:
        log(f"Chrome: {chrome_path}")

    with sync_playwright() as p:
        browser = p.chromium.launch(
            headless=False,
            executable_path=chrome_path or None,
        )

        context = _build_context(browser)

        if not ensure_logged_in(context):
            log("登录失败，退出")
            input("\n按 Enter 关闭浏览器...")
            browser.close()
            sys.exit(1)

        if login_only:
            log("登录完成（--login-only）。会话已保存。")
            input("\n按 Enter 关闭浏览器...")
            browser.close()
            return

        if monitor_mode:
            max_rounds = (MAX_POLL_MINUTES * 60) // poll_interval
            log(f"\n监控模式启动 (间隔 {poll_interval}s, 最长 {MAX_POLL_MINUTES} 分钟)")
            log(f"课程优先级: {', '.join(courses)}")

            for rnd in range(1, max_rounds + 1):
                log(f"\n── 第 {rnd} 轮 ──")

                for course in courses:
                    page = context.new_page()
                    if do_booking(page, course):
                        log(f"\n🎉🎉🎉 预约成功！已退出")
                        input("\n按 Enter 关闭浏览器...")
                        browser.close()
                        return

                if rnd < max_rounds:
                    log(f"等待 {poll_interval} 秒...")
                    time.sleep(poll_interval)

            log(f"\n⏰ 监控超时 ({MAX_POLL_MINUTES} 分钟)")
        else:
            log(f"\n单次模式：尝试 {len(courses)} 门课程")
            for course in courses:
                page = context.new_page()
                if do_booking(page, course):
                    log(f"\n🎉 预约成功！")
                    break
            else:
                log(f"\n😔 所有课程均无可预约时段")

        input("\n按 Enter 关闭浏览器...")
        browser.close()


if __name__ == "__main__":
    main()
