# lab-scraper

HITWH（哈尔滨工业大学威海）校园自动化工具集，包含三个部分：

1. **选课工作台** — 本地只读的教务选课规划工具：导入选课通知 → 确认选课窗口 →
   导入课表 → 只读发现选课接口 → 规划课程目标
2. **实验室抢课** — 开放式实验系统自动预约：CAS SSO 登录、会话持久化、监控抢座
3. **毕业进度** — 学业成绩与毕业要求匹配，估算剩余学业任务

所有认证信息仅在本地加密保存（Windows DPAPI），数据不上传。

## 快速开始

```powershell
# 1. 安装依赖（含 dev 工具）
uv sync

# 2. 首次使用安装浏览器内核（Playwright 需要）
uv run playwright install chromium

# 3. 查看可用命令
uv run course-selection --help
```

要求 Python >= 3.10（项目用 `.python-version` 固定版本，`uv sync` 自动处理）。

## 常用命令

### 选课工作台

| 命令 | 说明 |
|------|------|
| `uv run course-selection workbench` | 启动本地工作台（默认 `http://127.0.0.1:5000`，离线启动） |
| `uv run course-selection dev-workbench` | 开发模式：持久 Chromium + 源码热重启（watchfiles） |
| `uv run course-selection configure-profile --grade 2025` | 保存本地学生画像 |
| `uv run course-selection explore-entry` | 只读探测选课入口 |
| `uv run course-selection discover-timetable` | 自动点击发现课表只读接口 |
| `uv run course-selection discover-selection` | 自动点击发现选课只读接口 |
| `uv run course-selection analyze-interface --target student-profile` | 分析学生画像接口契约 |

Windows 下双击 `start-workbench.cmd` 也可启动工作台（无控制台窗口）。

### 实验室抢课（cas-book）

| 命令 | 说明 |
|------|------|
| `uv run course-selection cas-book` | 单次尝试所有目标课程 |
| `uv run course-selection cas-book --monitor` | 监控模式，每 5 秒查一轮，最长 120 分钟 |
| `uv run course-selection cas-book --monitor --interval=3` | 自定义监控间隔（秒） |
| `uv run course-selection cas-book --course "DIY电磁混合磁悬浮"` | 只预约指定课程 |
| `uv run course-selection cas-book --login-only` | 仅登录并保存会话 |

首次运行会在浏览器中手动登录 CAS，成功后会话保存到 `storage_state.json`，
后续自动复用；会话失效时才会再次要求登录。

### 登录配置

| 命令 | 说明 |
|------|------|
| `uv run course-selection configure-login` | 保存 WebVPN 登录信息（DPAPI 加密） |

## 配置

配置项通过环境变量或项目根目录 `.env` 文件提供（模板见 `.env.example`），
环境变量优先级高于 `.env`。

| 变量 | 默认值 | 说明 |
|------|--------|------|
| `WORKBENCH_PORT` | `5000` | 工作台端口 |
| `WORKBENCH_HOST` | `127.0.0.1` | 工作台监听地址（仅回环） |
| `WORKBENCH_PRIVATE_ROOT` | `.private/academic-selection` | 选课工作区数据目录 |
| `PROGRESS_PROFILE_ROOT` | `.private/course-progress` | 教务登录与进度数据目录 |
| `ACADEMIC_BROWSER_DEBUG_PORT` | `9222` | dev-workbench CDP 调试端口 |
| `ACADEMIC_WORKBENCH_DEV_DIAGNOSTICS` | `0` | 开发诊断开关 |
| `CAS_BOOK_POLL_INTERVAL` | `5` | 抢课监控间隔（秒） |
| `CAS_BOOK_MAX_POLL_MINUTES` | `120` | 抢课监控最长时长（分钟） |
| `CAS_BOOK_BASE_URL` | `http://openlab.hitwh.edu.cn` | 实验系统入口 |
| `ACADEMIC_BROWSER_CDP_URL` | — | 复用外部浏览器 CDP 地址 |

## 测试与开发

```powershell
uv run pytest tests/        # 运行全部测试（196 个）
```

开发工作台 `dev-workbench` 会启动一个长期持有的 Chromium（profile 持久化），
并监控 `course_selection/` 与 `course_progress/` 下的 Python 源码变化：
只重启工作台进程，浏览器标签页、登录状态和 CDP 连接保持不变。
进行中的教务任务完成前不会触发重启；`Ctrl+C` 关闭工作台和浏览器，保留本地数据。

## 架构

```
course_selection/         选课工作台
├── cli.py                click 命令入口（course-selection / lab-book）
├── application.py        waitress 服务器 + 可见 Chromium 外壳
├── workbench.py          回环 Flask 适配层（CSRF/CSP 防护）
├── gateway.py            Playwright 教务网关（只读边界）
├── discovery.py          接口发现（拦截疑似选课/退课/保存请求）
├── tasks.py              观察任务与执行任务队列（SQLite 持久化）
├── config.py             集中配置（环境变量 + .env）
└── dev_workbench.py      开发监督：持久 Chromium + watchfiles 热重启

course_progress/          教务成绩与毕业进度
openlab_cas_book.py       实验室抢课旧入口（shim，转发到 cas-book 子命令）
frontend/                 Vite 前端构建产物来源
tests/                    测试（196 passed）
```

## 安全边界

- 工作台只监听 `127.0.0.1`，带 CSRF token 与 CSP 头防护
- 接口发现只自动点击导航/查询控件，疑似选课、退课、保存请求会被拦截
- 选课执行任务必须用户明确确认具体教学班后才提交
- 登录信息用 Windows DPAPI 加密，仅当前用户可解密；`.private/` 不入库
- 会话状态 `storage_state.json` 包含登录令牌，已在 `.gitignore` 中排除

## 兼容旧入口

`openlab_cas_book.py` 保留为转发脚本，旧用法仍然可用：

```powershell
uv run python openlab_cas_book.py --monitor --interval=3
```
