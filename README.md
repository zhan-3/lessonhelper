# LessonHelper

HITWH（哈尔滨工业大学威海）校园自动化工具集，提供本地化的选课规划、实验预约和毕业进度分析能力。

> 本项目面向个人学习与研究用途。请遵守学校系统使用规范，不要高频请求、绕过验证码或自动提交未经确认的操作。

## 功能

- **选课工作台**：解析选课通知、导入课表、识别时间冲突、读取待选课程并在用户确认后提交单个教学班。
- **实验预约**：通过 HIT CAS 统一身份认证预约开放式实验，可单次尝试或定时监控。
- **毕业进度**：读取个人教务数据，按培养要求估算课程和学分缺口。
- **本地优先**：认证信息、Cookie、课表和教务快照保存在本机，不上传到项目服务器。

## 安全与隐私

- 登录信息使用 Windows DPAPI 加密保存，仅当前 Windows 用户可解密。
- 浏览器会话、个人数据和预约结果不会提交到 Git。
- 工作台默认只监听 `127.0.0.1`，远程操作必须由用户明确触发。
- 每位使用者必须使用自己的 CAS 账号登录；不要分享 `.private/`、`.env` 或 `storage_state.json`。

## 快速开始

### 环境要求

- Windows
- Python >= 3.10
- [uv](https://docs.astral.sh/uv/getting-started/installation/)

### 安装

```powershell
# 克隆仓库
git clone https://github.com/zhan-3/lessonhelper.git
cd lessonhelper

# 安装 Python 依赖和 Playwright Chromium
setup.cmd
```

也可以手动执行：

```powershell
uv sync
uv run playwright install chromium
```

### 首次配置

```powershell
# 配置教务系统登录信息（本地 DPAPI 加密）
uv run course-selection configure-login

# 配置学生画像，年级按实际情况填写
uv run course-selection configure-profile --grade 2025
```

### 启动工作台

```powershell
uv run course-selection workbench
```

也可以双击 `start-workbench.cmd`。工作台默认地址为 <http://127.0.0.1:5000>。

## 常用命令

### 选课工作台

| 命令 | 说明 |
| --- | --- |
| `uv run course-selection workbench` | 启动本地工作台 |
| `uv run course-selection dev-workbench` | 开发模式，保留浏览器和登录状态并支持热重启 |
| `uv run course-selection explore-entry` | 只读探测选课入口 |
| `uv run course-selection discover-timetable` | 发现课表只读接口 |
| `uv run course-selection discover-selection` | 诊断选课只读接口 |
| `uv run course-selection analyze-interface --target student-profile` | 分析学生画像接口契约 |

### 实验预约

```powershell
# 使用默认目标课程进行一次预约尝试
uv run course-selection cas-book

# 监控模式
uv run course-selection cas-book --monitor --interval=5

# 指定课程
uv run course-selection cas-book --course "DIY电磁混合磁悬浮"

# 只登录并保存实验系统会话
uv run course-selection cas-book --login-only
```

实验系统首次使用时会打开浏览器，用户需要在 HIT CAS 页面完成登录。会话保存在本地，失效后需要重新认证。

### 毕业进度

```powershell
uv run python -m course_progress --help
```

详细说明见 [`docs/course-progress-explorer.md`](docs/course-progress-explorer.md)。

## 为什么默认显示浏览器

本项目目前**不提供默认后台/无头操作模式**，这是有意设计：

- CAS 登录可能需要验证码或人工确认；
- 用户需要看到实际页面，避免误提交选课或预约；
- 部分校园系统对无头浏览器和高频访问更敏感；
- 登录失效时，后台进程无法可靠地完成重新认证。

技术上，在已有有效会话的情况下可以增加 headless 模式，但它不能消除首次登录、验证码和会话失效时的人工步骤，也不建议把它作为默认行为。后续如增加，也应设计为“用户先可见登录，之后可选后台运行”，并保留明确的停止和失败保护。

## 配置

配置通过环境变量或项目根目录 `.env` 提供，模板见 [`.env.example`](.env.example)。

常用配置：

| 变量 | 默认值 | 说明 |
| --- | --- | --- |
| `WORKBENCH_PORT` | `5000` | 工作台端口 |
| `WORKBENCH_HOST` | `127.0.0.1` | 工作台监听地址 |
| `WORKBENCH_PRIVATE_ROOT` | `.private/academic-selection` | 工作台本地数据目录 |
| `PROGRESS_PROFILE_ROOT` | `.private/course-progress` | 教务登录和进度数据目录 |
| `CAS_BOOK_BASE_URL` | `http://openlab.hitwh.edu.cn` | 实验系统入口 |
| `CAS_BOOK_POLL_INTERVAL` | `5` | 实验监控间隔（秒） |
| `CAS_BOOK_MAX_POLL_MINUTES` | `120` | 实验监控最长时间（分钟） |

## 开发与测试

```powershell
uv sync
uv run pytest tests/
```

前端源码位于 `frontend/`，构建后的工作台静态资源位于 `course_selection/workbench_static/`。

项目结构和教务领域约定见 [`CONTEXT.md`](CONTEXT.md)；选课工作台详细说明见 [`docs/academic-selection.md`](docs/academic-selection.md)。

## 项目状态

当前版本适合个人本地使用。学校页面、认证流程和选课规则可能变化，任何真实环境操作都应先人工核验，并从低风险操作开始。

## 许可证

本仓库当前尚未声明开源许可证。未经作者另行授权，不应将代码用于再分发或商业用途。
