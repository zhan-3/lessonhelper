# lab-scraper

HITWH 开放式实验系统抢课自动化工具。

## 快速开始

```powershell
uv sync                      # 安装依赖（含 dev 工具）
uv run course-selection --help
```

## 常用命令

| 命令 | 说明 |
|------|------|
| `uv run course-selection workbench` | 启动本地选课规划工作台（默认 `http://127.0.0.1:5000`） |
| `uv run course-selection dev-workbench` | 开发模式：持久 Chromium + 源码热重启 |
| `uv run course-selection cas-book` | CAS SSO 登录自动预约（实验室抢课） |
| `uv run course-selection cas-book --monitor --interval=3` | 监控模式，每 3 秒查一次 |
| `uv run course-selection cas-book --course "DIY电磁混合磁悬浮"` | 只预约指定课程 |
| `uv run course-selection cas-book --login-only` | 仅登录并保存会话 |
| `uv run course-selection configure-login` | 保存 WebVPN 登录信息（DPAPI 加密） |
| `uv run course-selection configure-profile --grade 2025` | 保存本地学生画像 |
| `uv run course-selection explore-entry` | 只读探测选课入口 |
| `uv run course-selection discover-timetable` / `discover-selection` | 自动发现只读接口 |

Windows 下双击 `start-workbench.cmd` 也可启动工作台（无控制台窗口）。

## 配置

配置通过环境变量或项目根目录的 `.env` 文件提供（见 `.env.example`），
环境变量优先级最高。常用项：

- `WORKBENCH_PORT`（默认 5000）
- `WORKBENCH_PRIVATE_ROOT`（默认 `.private/academic-selection`）
- `CAS_BOOK_POLL_INTERVAL`（默认 5 秒）

## 测试

```powershell
uv run pytest tests/
```

## 架构

- `course_selection/` — 选课工作台（Flask + Playwright + SQLite 持久化）
- `course_progress/` — 教务成绩进度采集
- `openlab_cas_book.py` — 实验室抢课脚本入口（兼容旧用法，转发到
  `course-selection cas-book`）
