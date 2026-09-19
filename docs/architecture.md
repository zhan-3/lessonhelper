# 项目框架与技术栈

**一句话**：本地优先的校园教务辅助工具——读取教务与开放式实验系统的数据，在本机做规划、推算与受控提交；数据、凭据与结果都留在本机。

> 这是开发快照。各功能的证据等级见 README 的成熟度表与 `AGENTS.md` 的完成度约定。

---

## 1. 技术栈

| 层 | 选型 | 实际版本 | 说明 |
| --- | --- | --- | --- |
| 语言 / 运行时 | Python | 3.10.20（`requires-python >=3.10`） | 后端全部逻辑 |
| 包与虚拟环境 | uv（`uv run` / `uv sync`） | — | `[tool.uv] package = true` |
| 打包 | setuptools | ≥64 | 只打包 `course_selection*`、`course_progress*` |
| HTTP 服务 | Flask + waitress | flask ≥3.1、waitress ≥3.0 | 只监听 `127.0.0.1`；waitress 作生产级 WSGI |
| 浏览器自动化 | Playwright（Chromium） | ≥1.58 | 可见浏览器；支持 `connect_over_cdp` 借用标签页 |
| CLI | click | — | 单一入口 `course-selection` |
| 本地存储 | SQLite（标准库 `sqlite3`） | — | 工作台唯一事实来源（ADR-0004） |
| 凭据保护 | Windows DPAPI（`ctypes`） | — | 账号密码加密后落盘，仅当前 Windows 用户可解 |
| 表格导入 | openpyxl / xlrd | — | 用户导入课表（`.xlsx` / `.xls`） |
| 配置 | python-dotenv | ≥1.0 | `.env` + 环境变量 |
| 前端 | React + TypeScript + Vite | React 19.2.8 / TS 5.9.3 / Vite 8.2.2 | 工作台界面 |
| 前端类型来源 | openapi-typescript | latest | 由 `openapi.yaml` 生成 `api.generated.ts` |
| 前端测试 | Vitest + Testing Library + jsdom | Vitest 4.1.11 | 23 个界面测试 |
| 后端测试 | pytest + `unittest`（含 subTest） | pytest ≥8 | 295 个测试 |
| 静态检查 | ruff | ≥0.8 | `target-version = "py310"`，忽略 `BLE001` |
| 开发辅助 | watchfiles | ≥1.0 | 开发时热重载 |
| AI 侧扩展 | Pi Extension（TypeScript） | — | `.pi/extensions/academic-browser-observer.ts`，只读观察浏览器 |

---

## 2. 运行时与入口

```text
course-selection      = course_selection.cli:main           ← 当前唯一正式入口
course-selection-gui  = course_selection.cli:main
```

CLI 子命令：

| 命令 | 作用 |
| --- | --- |
| `workbench` / `dev-workbench` | 启动本地工作台（Flask + 可见 Chromium 外壳） |
| `configure-login` / `configure-profile` | 保存 DPAPI 凭据 / 学生画像 |
| `discover-timetable` / `discover-selection` / `explore-entry` / `analyze-interface` | 只读发现与契约候选分析 |
| `lab-booking` | 实验预约：只读规划 + 单次提交（实验状态） |
| `lab-contract` | 接口契约快照：`record` / `check` / `promote` |

工作台的 HTTP 面：`course_selection/workbench.py`（Flask 蓝图/工厂）+ `openapi.yaml`（契约）+ `workbench_static/`（Vite 构建产物）。
入口与进程管理：`cli.py`（click 子命令）、`application.py`（waitress 托管）、
`dev_workbench.py`（开发期热重载外壳）、`single_instance.py`（可恢复的 workspace 级单实例锁）。

---

## 3. 模块地图

### 3.1 应用核心（无 Flask、无传输实现）

核心模块不 import 具体传输实现。需要外部世界的用例通过构造参数接收依赖，
因此可以直接用普通函数或假对象测试。

| 模块 | 职责 |
| --- | --- |
| `workbench_service.py` | 应用服务：状态、基线选择、认定学分、课程标签、进度投影、规划；通知读取经 `notice_fetcher` / `notice_discoverer` 注入 |
| `planning.py` | 只读规划与冲突计算（纯函数） |
| `persistence.py` | SQLite：快照、画像、通知、计划、执行历史、契约相关表；`reset_personal_workspace` |
| `timetable.py` | 课表导入：解析 xls/xlsx 与页面网格，规范化为冲突可用模型 |
| `categories.py` | 课程类别映射 |
| `notice.py` | 选课通知的领域模型、正文解析与本地读写（无 IO 传输） |
| `notice_discovery.py` | 官方通知的纯解析、主机白名单校验、候选与差异 |
| `lab_booking.py` | 实验预约域：占用区间、互斥规划、单次提交护栏；`LabSession` Protocol |
| `lab_contract.py` | 接口契约：快照、指纹、分级 diff、只读探针 |
| `lab_ports.py` | 实验侧抽象接缝：`LabTransport` Protocol 与端点常量；实现见 3.2 |
| `config.py` | 环境变量与路径 |

### 3.2 外部适配器（具体 IO）

| 模块 | 职责 |
| --- | --- |
| `gateway.py` | Playwright 网关：WebVPN 代理路径、课表/待选读取、CAS 会话自愈 |
| `notice_transport.py` | 通知读取 IO：公开页面 HTTP 抓取、已登录浏览器读取、官方索引发现 |
| `lab_transport.py` | 实验侧传输：纯 HTTP 主后端 / 页面内 fetch 回退 / 令牌获取；re-export `lab_ports` 的名字 |
| `lab_browser_session.py` | `BrowserLabSession`：在 `LabTransport` 之上实现核心的 `LabSession`，借用已登录标签页 |
| `selection_query.py` / `selection_entry.py` / `selection_execution.py` | 待选课程查询、只读入口、单次提交执行（暂停使用） |
| `current_enrollment.py` / `personal_timetable.py` | 本学期已选、个人课表快照 |
| `discovery.py` / `deep_observation.py` / `manual_observation.py` | 只读发现与观察路径 |
| `tasks.py` | 观察与执行任务的串行调度、状态机与崩溃恢复 |
| `shadow_acceptance.py` | 只读影子验收的 fail-closed 报告 |
| `student_profile.py` / `student_profile_observation.py` | 学生画像的本地事实与只读读取探测 |
| `web.py` | 旧版选课工作台（基于本地 JSON 文件，与 `workbench.py` 并存） |

### 3.3 浏览器观察（实验性）

`browser_observer.py`、`browser_observer_worker.py`、`browser_redaction.py`、`request_contracts.py`、`observer_evals.py` 组成"借用标签页 + 脱敏证据 + 请求候选"的只读观察链；对应的 Pi Extension 在 `.pi/extensions/`。

### 3.4 `course_progress`（成绩与毕业进度）

`academic_client.py`（教务成绩读取）、`collector.py`（采集与快照组装）、`progress.py`（进度计算）、`baselines.py`（不可变要求基线）、`credentials.py`（DPAPI）、`session.py`（CAS/WebVPN 会话状态机）、`explorer.py`（门户导航）、`sanitizer.py`（脱敏）、`capture.py`（网络交换存档与候选排序）、`cli.py`（`python -m course_progress` 入口）。

### 3.5 前端

```text
frontend/src/
  main.tsx                 应用外壳、状态、数据操作抽屉
  ScheduleBoard.tsx        课表 + 进度面板 + 待选课程浏览器
  schedule.ts              纯函数：课表展开、冲突、筛选、身份规范化
  api.ts / api.generated.ts  类型（由 openapi.yaml 生成）
  style.css
```

---

## 4. 分层与依赖方向

```text
                        ┌────────────────────────────────────────┐
                        │  前端 (React/TS, Vite 构建)             │
                        └───────────────┬────────────────────────┘
                                        │ HTTP + CSRF（同源）
                        ┌───────────────▼────────────────────────┐
                        │  HTTP 适配器  workbench.py / openapi   │
                        └───────────────┬────────────────────────┘
                                        │ 只调用应用服务
                        ┌───────────────▼────────────────────────┐
                        │  应用核心  workbench_service / planning│
                        │            lab_booking / lab_contract  │
                        └──────┬───────────────────┬─────────────┘
                               │                   │
                ┌──────────────▼──────┐   ┌────────▼─────────────────┐
                │ 存储 SQLite          │   │ 外部适配器                │
                │ persistence.py       │   │ gateway.py（教务，浏览器）│
                │ .private/（Git 忽略）│   │ notice_transport.py（通知）│
                └─────────────────────┘   │ lab_transport.py（实验）  │
                                          │ lab_browser_session.py    │
                                          └──────────────────────────┘
```

依赖规则：

1. **核心不 import Flask**，也不 import 具体传输实现（`gateway`、`lab_transport`、
   `lab_browser_session`、`notice_transport`）。
2. 外部世界（浏览器、HTTP、磁盘）通过构造参数或端口注入，因此核心可以直接用假对象做单元测试。
3. 抽象与实现分离：`LabTransport`（`lab_ports`）与 `LabSession`（`lab_booking`）
   定义在核心，`lab_transport` / `lab_browser_session` 指向它们。

层间依赖方向由 **import-linter 契约**强制（`pyproject.toml` 的 `[tool.importlinter]`）：

| 方向 | 数量 |
| --- | --- |
| core → adapter | 0 |
| core → http | 0 |
| adapter → http | 0 |
| adapter → core | 允许（实现依赖抽象） |

例外：`lab_contract._default_get` 用标准库 `urllib` 提供只读探针的默认实现，
便于 `observe` 在无注入时也能工作；需要避免 IO 的调用方传入 `static_get`。

---

## 5. 数据与落盘

| 内容 | 位置 | 入仓？ |
| --- | --- | --- |
| 工作台数据库（画像、快照、计划、执行历史、申报、标签） | `.private/academic-selection/workbench.sqlite3` | 否 |
| 教务凭据（DPAPI）与会话状态 | `.private/course-progress/` | 否 |
| 契约观测 | `.private/lab-contracts/*.observed.json` | 否 |
| 学校原始参考资料（培养方案 / 生活指南 `.docx`，含文档作者元数据） | `.private/reference/` | 否（`.gitignore`：`docs/*.docx`） |
| **接口契约基线** | `docs/contracts/<channel>-<center>.json` | **是**（无个人数据） |
| 前端构建产物 | `course_selection/workbench_static/` | 是 |
| 浏览器快照（`storage_state.json`、`*.har`）、课表表格（`*.xls*`） | 仓库根或桌面 | 否（`.gitignore`：`.private/`、`.env`、`storage_state.json`、`*.har`、`*.xls*`） |

`reset_personal_workspace` 会清空 9 张表 + 两个 metadata 键，用于"换人使用"。

---

## 6. 认证（详见 `docs/authentication.md`）

```text
教务（jwts） ：WebVPN 反向代理 /http/<hex>/… + CAS 会话
openlab      ：gateway goto/{中心} → StuApi auth/cas/login → 请求头 vctchauthorization
共同点        ：同一套统一身份认证凭据，但不共享会话
```

工具侧约定：**业务调用走纯 HTTP + 请求头**（实测不需要 Cookie）；浏览器只在"取令牌 / 人工登录"时参与；令牌只存内存。

---

## 7. 契约与漂移

```text
lab-contract record   只读观测 → .private/lab-contracts/
lab-contract check    与基线分级比对；breaking 或读不到 → 退出码 1
lab-contract promote  观测提升为 docs/contracts/ 基线（保留人工 locked 字段）
```

比对三样：应用版本串、端点集合、响应顶层字段名（含 `data[].classDate` 这类点路径）。
分级：`breaking` / `unavailable` / `additive` / `notice` / `cosmetic` / `rebaseline`。

---

## 8. 构建与质量门

```powershell
uv sync                                   # 后端依赖
uv run pre-commit install                 # 一次性：装提交钩子

# 提交时自动运行（配置见 .pre-commit-config.yaml）：
#   文件卫生（含大文件拦截）/ ruff check --fix / mypy / import-linter / 结构自检
uv run pre-commit run --all-files         # 手动全量跑一遍钩子

uv run pytest tests/                      # 295 个测试（约 23 秒，未进钩子）
uv run mypy                               # 类型检查，排除清单见 pyproject.toml
uv run lint-imports                       # 分层契约与循环依赖，契约见 pyproject.toml

cd frontend; npm run generate:api         # 由 openapi.yaml 生成类型
cd frontend; npm test                     # 23 个界面测试
cd frontend; npm run build                # tsc -b && vite build → workbench_static/
```

`pytest` 不放进提交钩子（约 23 秒太慢），只在 CI 运行。

质量门在两个地方执行**同一套配置**：本地提交钩子（`.pre-commit-config.yaml`）与
GitHub Actions（`.github/workflows/ci.yml`）。CI 的后端 job 用 `windows-latest`
——`tests/test_workbench.py` 的登录配置测试会调用真实的 Windows DPAPI，非 Windows
平台按设计抛错；前端 job 在 ubuntu 上跑 Vitest。依赖安装用 `uv sync --locked`，
lock 与 pyproject 不一致会直接失败。

`mypy` 采用渐进式类型化：`pyproject.toml` 的 `ignore_errors` 列出了 15 个仍有
已知类型错误的模块（共 114 个错误，多数源于 `dict[str, object]` 承载结构化数据）。
其余 37 个模块与新增代码受类型保护；修复一个模块就从清单里删掉一行。

前端构建产物是**提交进仓库的**，因此改前端后必须重新 build，否则界面与源码不一致。

---

## 9. 关键约定（ADR）

| ADR | 约定 |
| --- | --- |
| 0001 | 学术数据只留本地 |
| 0002 | 要求与类别映射共同版本化 |
| 0003 | 应用核心独立于 Flask |
| 0004 | SQLite 作为本地唯一事实来源 |
| 0005 | 选课规划界面用 React |
| 0006 | 个人课表快照作为事实 |
| 0007 | 教务读取固定为携带身份的请求 |

其他不写成 ADR 但同样生效的约定：观察任务与执行任务分离；写操作需显式确认、每目标一次、失败不重试；`HTTP 200` 不等于成功（业务码在信封里）。

---

## 10. 已知不一致（待清理）

- `lab-booking` 仍使用页面内 `fetch` 后端；`lab-contract` 已走纯 HTTP 主后端。两者共用同一传输层，接线统一尚未完成。
- 毕业基线、实验预约的真实环境验收均未完成；Observer 的语义价值仍未证明。
- `docs/` 下两个 `.docx` 原件（共 15.5 MB，含文档作者元数据）已移出仓库至 `.private/reference/`；**但 Git 历史中仍保留副本，仓库体积尚未真正瘦身**，参见本文第 5 节与 `reading-map.md`。

---

## 11. 文档索引

| 文档 | 内容 |
| --- | --- |
| `README.md` | 功能成熟度与使用方式 |
| `docs/reading-map.md` | 目录树、阅读优先级、交叉验证清单 |
| `docs/authentication.md` | 认证框架、冷启动时序、会话寿命与失败分类 |
| `docs/academic-selection.md` | 工作台、实验预约、契约检查的操作说明 |
| `docs/course-progress-explorer.md` | 成绩与毕业进度读取 |
| `docs/academic-browser-observer.md` | 浏览器观察扩展 |
| `CONTEXT.md` | 领域词汇与边界 |
| `docs/adr/` | 架构决策记录 |
| `.scratch/` | 议题与规格（本地 issue tracker） |
