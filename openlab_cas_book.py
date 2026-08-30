"""
openlab_cas_book.py

Playwright 自动化：CAS SSO 登录 → 开放式实验系统 → 自动预约

流程：
  首次运行：自动打开浏览器 → 跳转到 CAS 登录页 → 用户手动输账号密码/验证码 → 登录成功 → 保存会话
  后续运行：加载已保存会话 → 直接操作预约页面
  支持 --monitor 监控模式：每隔 N 秒检查，有名额立即抢

用法：
  uv run lab-book cas-book                          # 单次尝试所有课程
  uv run lab-book cas-book --monitor                # 监控模式
  uv run lab-book cas-book --monitor --interval=3   # 每 3 秒查一次
  uv run lab-book cas-book --course "课程名"         # 指定单个课程
  uv run lab-book cas-book --login-only             # 仅登录保存会话，不预约

  或直接用主 CLI：
  uv run course-selection cas-book [选项]
"""

from course_selection.cli import lab_book_main

if __name__ == "__main__":
    lab_book_main()
