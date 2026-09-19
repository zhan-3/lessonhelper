# LessonHelper

[![CI](https://github.com/zhan-3/lessonhelper/actions/workflows/ci.yml/badge.svg)](https://github.com/zhan-3/lessonhelper/actions/workflows/ci.yml)

HITWH（哈尔滨工业大学（威海））校园教务辅助工具集：本地优先的选课规划、毕业进度推算，以及实验预约的探索性实现。

> **这是开发快照，不是稳定发布版，也不是可依赖的抢课/选课服务。**
> 部分功能只经过自动化测试，尚未在真实教务环境完成验收。请勿在真实选课或预约场景中直接依赖当前代码。

---

## 功能成熟度

本项目按证据强度区分三种状态，README 中的描述与代码实际证据保持一致：

- `implemented` — 代码存在且可运行；
- `automated-test verified` — 有自动化测试覆盖，只能证明代码行为，**不能证明与学校线上系统兼容**；
- `real-environment verified` — 已在真实学校系统上人工验证。

| 能力 | 状态 | 说明 |
| --- | --- | --- |
| 选课工作台（本地画像、通知确认、课表导入、冲突规划、只读查询） | `implemented` + `automated-test verified` | 真实教务读取链路尚未完成完整验收 |
| 选课提交（单个教学班单次提交） | `implemented` + `automated-test verified` | 真实提交流程未验收；默认不启用自动重试 |
| 毕业进度推算 | `implemented` + `automated-test verified` | 结果仅为规划参考，必须人工核对培养方案 |
| 实验预约 `lab-booking` | **实验性，未完成开发与真实环境验证** | 视为不可用功能，仅在小范围、可人工核对时尝试 |
| Academic Browser Observer（只读浏览器诊断） | 部分 `automated-test verified`，价值主张 `not_demonstrated` | 未在真实教务系统验证，不能替代现有 DevTools/Playwright 工具 |

当前测试基线：`uv run pytest tests/` → **282 passed**。

---

## 安全与隐私边界

- **数据全部留在本机。** 凭据、Cookie、浏览器 profile、课表、成绩与预约结果一律不提交到 Git，也不上传到任何项目服务器。
- **登录信息使用 Windows DPAPI 加密**，保存在 `.private/` 下，只有当前 Windows 用户可解密。
- **严格区分观察任务与执行任务**（见 [`CONTEXT.md`](CONTEXT.md)）。只读观察不会改变任何选课状态。
- **执行任务必须显式确认**：提交前需要用户针对具体教学班确认，最多提交一次；结果不明时立即停止，不自动重试。
- **优先使用可见浏览器**处理认证和会改变状态的操作；本项目当前不提供默认后台/无头模式。
- 分享或提交前请确认以下内容未被跟踪：`.private/`、`.env`、`storage_state.json`、`*.xls`、`*.xlsx`。

每位使用者应使用自己的 CAS 账号登录，不要共享凭据、会话状态或个人数据目录。

---

## 认证框架

两个目标系统共享同一个统一身份认证（CAS）凭据，但**不共享会话**，认证链彼此独立：

```
                     统一身份认证 / CAS  (ids.hit.edu.cn)
                                 │
             ┌───────────────────┴───────────────────┐
             │                                       │
     【教务系统链路】                            【openlab 链路】
  webvpn.hitwh.edu.cn 门户                    openlab.hitwh.edu.cn
  ├ /login?cas_login=true#!/service           ├ 「统一身份认证登录」按钮
  ├ 凭据自动填写 …/authserver/login            ├ 在可见浏览器中手动完成
  └ 会话校验 GET /user/info                    └ 会话存 storage_state.json
             │                                       │
  WebVPN 反向代理 /http/<hex>/…                应用自定义请求头
  （编码指向 jwts.hitwh.edu.cn）               vctchauthorization
             │                                       │
    新教务系统 jwts.hitwh.edu.cn             /<center>/StuApi/<path>
```

“校园网 / 校外 VPN”这个变量只对 openlab 有意义：

| 网络位置 ↓ ／ 目标系统 → | 教务系统（`jwts.hitwh.edu.cn`） | openlab 开放式实验系统 |
| --- | --- | --- |
| **校园网内网直连** | 代码中**没有**直连路径，仍走 WebVPN 反向代理 | 设计上的主路径：明文 HTTP + `vctchauthorization`，无需 Cookie |
| **校外 ／ 已连 VPN** | 主路径，也是唯一路径：WebVPN 反向代理 + CAS 会话 | 代码中**没有** WebVPN 代理 openlab 的路径，只能借用户自己已可达的已登录标签页 |

要点：

- 教务系统在代码中始终经由 `webvpn.hitwh.edu.cn/http/<hex>/…` 代理访问，即使身处校园网也不例外。
- openlab 的 `vctchauthorization` 由页面自身请求产生，仍需一次浏览器（CDP 借用已登录标签页）才能读到；该 token 只存在于内存，从不落盘。
- `HttpLabTransport`（校园网明文 HTTP + 请求头认证）**已实现且有测试覆盖，但 CLI 尚未接线**，当前 `lab-booking` 一律走页面内 `fetch`。
- WebVPN 会话与客户端 IP 绑定；换网络会触发一次 `/login?logoutByIpChange=true`，代码将其作为一次性自愈检查处理。
- 代码中的 openlab 主机是明文 `http://`，其是否强制 HTTPS 尚未在真实环境核实。

完整矩阵、逐格证据与代码位置索引见 [`docs/authentication.md`](docs/authentication.md)。

---

## 环境要求

- Windows（DPAPI 凭据加密与 `.cmd` 脚本依赖 Windows）
- Python >= 3.10
- [uv](https://docs.astral.sh/uv/getting-started/installation/)

## 安装

```powershell
git clone https://github.com/zhan-3/lessonhelper.git
cd lessonhelper
setup.cmd
```

`setup.cmd` 等价于：

```powershell
uv sync
uv run playwright install chromium
```

## 首次配置

```powershell
# 配置统一身份认证凭据（Windows DPAPI 加密，仅当前用户可解密）
uv run course-selection configure-login

# 配置本地学生画像，年级按实际情况填写
uv run course-selection configure-profile --grade 2025
```

## 启动工作台

```powershell
uv run course-selection workbench
```

或双击 `start-workbench.cmd`。默认地址为 <http://127.0.0.1:5000>，仅监听 `127.0.0.1`。

工作台**离线启动**：只恢复本地画像、通知与教务快照；只有点击“连接教务”或明确的刷新按钮才会访问学校系统。

---

## 命令参考

### 选课工作台

| 命令 | 说明 |
| --- | --- |
| `uv run course-selection workbench` | 启动本地只读选课规划工作台 |
| `uv run course-selection dev-workbench` | 开发模式：持久 Chromium + CDP + Python 热重启 |
| `uv run course-selection explore-entry` | 只读探测已确认通知对应的选课入口 |
| `uv run course-selection discover-timetable` | 发现课表只读接口 |
| `uv run course-selection discover-selection` | 诊断选课只读契约为何失效 |
| `uv run course-selection analyze-interface --target student-profile` | 结构化分析学生画像接口契约（维护者用） |

工作台内的待选课程刷新使用版本化只读契约 `hitwh-jwts-selection-query-v1`。只有“已确认通知 + 学生画像年级”共同匹配出的选课类别才进入查询白名单；退课、申请、纸质/邮件办理事项会被排除。详细说明见 [`docs/academic-selection.md`](docs/academic-selection.md)。

### 实验预约（实验性，请勿视为可用）

```powershell
uv run course-selection lab-booking --center dxwl --plan-out .private/lab-plan.json
uv run course-selection lab-booking --center dxwl --confirm <规划令牌>
```

该命令**未通过真实环境验收**，仅适合小范围、可人工核对的场景。关键约束：

- 只读规划：读取实验项目、可选日期与大节，以及每个时段的空位；占用区间来自本地工作台课表快照。
- 规划令牌把一次提交绑定到具体目标，令牌不匹配时拒绝提交。
- 每门实验最多提交一次；结果不明记录为 `possibly_applied` 并立即停止，不自动重试。
- 服务端会拒绝查询**已预约**实验的排期与空位，换时段需先自行取消。

早期的 `cas-book` 命令、`lab-book` 入口与根目录 `openlab_cas_book.py` 已删除：它们面向旧版实验系统，与现系统的接口、证书与提交语义都不兼容。

### 毕业进度

```powershell
# 工作台内操作（推荐）
uv run course-selection workbench

# 或使用独立的只读探索/采集命令
uv run python -m course_progress explore
uv run python -m course_progress collect
```

毕业进度基于显式选择且不可原地修改的要求基线：

- 工作台使用 `basic-graduation-reference-v1`；命令行采集默认 `guide-2026`。
- 面板区分学校快照中的已确认贡献、本学期已选、队列预览、用户申报与手动标签；“预计可满足”不等于学校确认，“未知”表示必要证据仍不足。
- 申报、课程标签、体系选择与队列投影只写入本机 SQLite，不触发学校访问。切换基线后必须显式重新同步，旧报告仅作历史保留。
- **结果不包含专业必修总量或完整培养方案审核，不是学校正式毕业结论，必须对照个人培养方案人工核验。**

详细说明见 [`docs/course-progress-explorer.md`](docs/course-progress-explorer.md)。

### Academic Browser Observer（实验性，只读）

连接用户明确指定的本机 CDP 浏览器，在外部操作者执行一次操作期间捕获脱敏的拓扑与网络增量。

```text
academic_browser_begin(endpoint) → 外部操作 → academic_browser_finish()
```

也可使用完整生命周期：`connect` → `inspect` → `start` → `checkpoint`/`stop` → `disconnect`。

安全边界：只接受明确提供的 loopback CDP endpoint，不扫描端口、不启动浏览器、不提供 click/fill/navigate/任意 JavaScript 执行，borrowed 连接只能 detach；不保存原始 HAR、HTML、请求体、Cookie 或截图。

其 Agent A/B 评估中 target recall 与 candidate accuracy 均为 0%，整体价值主张仍为 `not_demonstrated`，候选结果**不是**已验证的生产读取契约。详见 [`docs/academic-browser-observer.md`](docs/academic-browser-observer.md)。

---

## 为什么默认使用可见浏览器

本项目目前**不提供默认后台/无头操作模式**，这是有意设计：

- CAS 登录可能需要验证码或人工确认；
- 用户需要看到实际页面，避免误提交选课或预约；
- 部分校园系统对无头浏览器和高频访问更敏感；
- 登录失效时，后台进程无法可靠地完成重新认证。

技术上可以在已有有效会话时增加 headless 模式，但它不能消除首次登录、验证码和会话失效时的人工步骤。后续如增加，也应设计为“用户先可见登录，之后可选后台运行”，并保留明确的停止与失败保护。

---

## 配置

配置通过环境变量或项目根目录 `.env` 提供（环境变量优先），模板见 [`.env.example`](.env.example)。

| 变量 | 默认值 | 说明 |
| --- | --- | --- |
| `WORKBENCH_PORT` | `5000` | 工作台端口 |
| `WORKBENCH_HOST` | `127.0.0.1` | 工作台监听地址 |
| `WORKBENCH_PRIVATE_ROOT` | `.private/academic-selection` | 工作台本地数据目录 |
| `ACADEMIC_BROWSER_DEBUG_PORT` | `9222` | `dev-workbench` 的 CDP 调试端口 |
| `ACADEMIC_WORKBENCH_DEV_DIAGNOSTICS` | `0` | 开发诊断开关 |
| `PROGRESS_PROFILE_ROOT` | `.private/course-progress` | 教务登录与进度数据目录 |
| `ACADEMIC_BROWSER_CDP_URL` | 空 | 观察器默认连接的 CDP 地址 |

## 本地数据目录

```text
.private/                        # 已 Git 忽略，不可分享
├── academic-selection/          # 工作台数据（画像、通知、快照、workspace.sqlite3）
├── course-progress/             # DPAPI 凭据、浏览器 profile、采集结果
├── interface-analysis/          # 学生画像接口分析结果（超 7 天自动清理）
└── observer-eval-runs.json      # Observer 本地评估结果
```

---

## 开发与测试

```powershell
uv sync
uv run pre-commit install     # 一次性：装提交钩子（之后的提交会自动检查）

uv run pytest tests/          # Python 测试（295 个）
uv run ruff check .           # Lint
uv run mypy                   # 类型检查（渐进式，排除清单见 pyproject.toml）
uv run lint-imports           # 分层契约与循环依赖
uv run python tools/check_project.py   # 结构自检：模块登记 / 敏感文件

cd frontend
npm install
npm test                      # 前端测试（23 个）

npm run build                 # 构建到 course_selection/workbench_static/
```

前端源码位于 `frontend/`（React + Vite + TypeScript），构建产物位于 `course_selection/workbench_static/`。

提交钩子与 CI（`.github/workflows/ci.yml`）使用同一套配置，避免「本地过、CI 挂」。
`pytest` 只在 CI 运行（本地钩子里约 23 秒太慢）。质量门的完整说明见
[`docs/architecture.md`](docs/architecture.md) 第 8 节。

## 文档索引

| 文档 | 内容 |
| --- | --- |
| [`CONTEXT.md`](CONTEXT.md) | 教务领域术语与项目约定 |
| [`docs/authentication.md`](docs/authentication.md) | 认证框架：WebVPN / CAS / openlab 四条组合 |
| [`docs/academic-selection.md`](docs/academic-selection.md) | 选课工作台、查询契约与执行边界 |
| [`docs/course-progress-explorer.md`](docs/course-progress-explorer.md) | 毕业进度探索器与基线规则 |
| [`docs/academic-browser-observer.md`](docs/academic-browser-observer.md) | 只读浏览器观察器与证据等级 |
| [`docs/DEVLOG.md`](docs/DEVLOG.md) | 开发日志 |
| [`docs/reading-map.md`](docs/reading-map.md) | 代码阅读地图：目录树、阅读优先级、交叉验证清单 |
| [`docs/adr/`](docs/adr/) | 架构决策记录 |
| [`AGENTS.md`](AGENTS.md) | 贡献者与 Agent 协作约定 |

---

## 项目状态与已知风险

当前版本是开发快照：

- 实验预约尚未完成开发与真实环境验证，不构成可用功能；
- 选课功能有较完整的实现与自动化测试，但真实教务读取与提交流程尚未完成完整验收；
- 毕业进度结果仅供辅助参考，必须人工核对；
- Observer 的价值主张未通过评估门槛，不要用它替代已验证的读取契约；
- 学校页面、认证流程与选课规则变化都可能使功能失效；
- 自动化测试只证明代码行为，**不证明**与学校线上系统兼容。

未完成的真实环境验证以发布风险形式记录，不会用 mock、fixture 或页面选择器推断成功。

## 许可证

本仓库当前尚未声明开源许可证。未经作者另行授权，不应将代码用于再分发或商业用途。

## 免责声明

本项目面向个人学习与研究用途。请遵守学校系统使用规范，不要高频请求、绕过验证码，或自动提交未经用户确认的操作。使用本工具产生的一切后果由使用者自行承担。
