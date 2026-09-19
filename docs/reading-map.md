# 代码阅读地图

回答两个问题：**这些目录是什么**，以及**该读哪些**。

面向接手、复审或修改本仓库的人与 agent。分层与依赖的完整说明见 `architecture.md`；
领域词汇见 `CONTEXT.md`；功能成熟度与证据等级见 `README.md` 与 `AGENTS.md`。

> 行数为撰写时快照，仅用于判断阅读量级，不是契约。文件增删不会让这份文档失效，
> 但具体数字会漂移。

---

## 1. 目录树

`course_selection/` 在磁盘上**是平铺的**，下面的分组是逻辑分层，与
`architecture.md` 第 4 节的依赖规则一致。

```text
10.course-crawler-py/
├── README.md / AGENTS.md / CONTEXT.md     成熟度声明 / 完成度约定 / 领域词汇
├── pyproject.toml                          依赖 + ruff 配置
│
├── course_selection/            选课工作台（后端主体，10.7k 行）
│   ├── cli.py                     567   唯一入口 `course-selection`
│   │
│   │  ② 应用核心 ─ 不 import 具体传输，可用假对象单测
│   ├── workbench_service.py       676   用例编排：状态/基线/标签/投影/规划
│   ├── persistence.py             569   SQLite 唯一事实来源 + 落盘脱敏
│   ├── lab_contract.py            474   契约快照 / 指纹 / 分级 diff
│   ├── lab_booking.py             369   实验预约：占用区间 / 单次提交护栏
│   ├── notice.py                  331   通知领域模型 + 正文解析
│   ├── timetable.py               318   课表导入（xls/xlsx 与页面网格）
│   ├── planning.py                206   只读规划 + 冲突判定（纯函数）
│   ├── notice_discovery.py        159   官方通知解析 + 主机白名单
│   ├── categories.py              101   课程类别映射
│   ├── config.py                   44   环境变量与路径
│   ├── lab_ports.py                28   LabTransport Protocol（抽象接缝）
│   │
│   │  ③ 外部适配器 ─ 真实 IO
│   ├── discovery.py               770   浏览器探索（一次性）
│   ├── gateway.py                 760   Playwright 教务网关 + CAS 自愈
│   ├── selection_entry.py         641   待选课程读取
│   ├── tasks.py                   501   观察/执行任务的串行调度与状态机
│   ├── student_profile_observation.py 287
│   ├── personal_timetable.py      286   个人课表快照
│   ├── deep_observation.py        264   深度观察
│   ├── lab_transport.py           231   实验侧传输（HTTP / 页面内 fetch）
│   ├── selection_query.py         211   待选查询
│   ├── shadow_acceptance.py       153   只读影子验收（fail-closed）
│   ├── manual_observation.py      147   人工观察
│   ├── notice_transport.py        143   通知抓取（HTTP / 已登录浏览器）
│   ├── lab_browser_session.py     118   借用已登录标签页
│   ├── current_enrollment.py       96   本学期已选
│   ├── selection_execution.py      92   单次提交执行（暂停使用）
│   ├── student_profile.py          76   学生画像本地事实
│   │
│   │  ④ 浏览器观察（实验性）
│   ├── browser_observer.py        634
│   ├── request_contracts.py       186
│   ├── observer_evals.py          144
│   ├── browser_redaction.py       106
│   ├── browser_observer_worker.py 103
│   │
│   │  ⑤ HTTP 适配器（Flask）
│   ├── workbench.py               475   loopback + CSRF + 38 路由
│   ├── dev_workbench.py           149   开发期热重载外壳
│   ├── web.py                     142   旧版选课工作台（本地 JSON）
│   ├── application.py             120   waitress 托管
│   ├── single_instance.py          71   workspace 级单实例锁
│   │
│   ├── openapi.yaml                     HTTP 契约（前端类型来源）
│   └── workbench_static/                281K Vite 构建产物（提交进仓库）
│
├── course_progress/             成绩与毕业进度（3.1k 行，输出为顾问性估算）
│   ├── progress.py                939   进度计算（核心）
│   ├── session.py                 430   CAS/WebVPN 会话状态机
│   ├── explorer.py                375   门户导航
│   ├── collector.py               341   采集与快照组装
│   ├── cli.py                     263
│   ├── baselines.py               197   不可变要求基线
│   ├── academic_client.py         194   成绩读取
│   ├── capture.py                 149
│   ├── sanitizer.py               140   脱敏
│   └── credentials.py             121   DPAPI 凭据加密
│
├── frontend/                    React 工作台（3.6k 行，含生成代码）
│   ├── src/main.tsx               919   外壳 + 25 个 useState   ← 复杂度热点
│   ├── src/schedule.ts            531   纯函数：展开/冲突/筛选
│   ├── src/ScheduleBoard.tsx      262   课表 + 进度面板 + 待选浏览器
│   ├── src/api.generated.ts      1270   ← 由 openapi.yaml 生成，不要读
│   ├── src/api.ts                  25
│   ├── src/style.css              400
│   └── src/*.test.ts(x)                 23 个前端测试
│
├── tests/                       295 个后端测试项（7.4k 行）
├── docs/                        架构 / 认证 / 契约基线 / 7 个 ADR
├── .agents/skills/              academic-system-fieldwork, security-review
├── .scratch/                    本地 issue tracker（7 个主题）
└── .private/                    🔒 git-ignored：sqlite / 凭据 / 观测结果
```

---

## 2. 阅读优先级

按"错了是否有真实后果"排序，而不是按代码量。

### 🔴 必读：安全与正确性边界（约 1,800 行）

这些地方出错会导致**真实后果**（错误提交、凭据泄露、本地服务暴露），
必须由人亲自持有，不能只看测试绿灯。

| 位置 | 为什么必读 |
| --- | --- |
| `planning.py` | 冲突判定决定"能不能选"。`conflict_unknown` 必须区别于"空闲" |
| `lab_booking.py` 的 `book_slots` | 写操作的护栏：每目标至多一次、失败不重试、结果未知即停 |
| `workbench.py` 的 `before_request` / `after_request` | loopback 限制 + CSRF + 安全响应头；错了等于把本地服务暴露出去 |
| `persistence.py` 的 `sanitize_for_storage` | 阻止 cookie / 学号 / 原始页面 HTML 落盘 |
| `credentials.py` | DPAPI 加密；错了泄露统一身份认证凭据 |
| `lab_contract.py` | 写操作前的契约校验；`HTTP 200` 不等于业务成功 |

### 🟡 懂公开接口即可（约 3,400 行）

读公开函数的签名、docstring 和返回值形状就够了，实现细节交给测试。

`workbench_service.py`、`gateway.py`、`persistence.py`、`notice.py`、
`notice_discovery.py`、`timetable.py`、`tasks.py`、`categories.py`

### ⚪ 可以不读（约 9,000 行）

| 内容 | 原因 |
| --- | --- |
| `frontend/src/api.generated.ts` | 由 `openapi.yaml` 自动生成 |
| `course_selection/workbench_static/` | Vite 构建产物 |
| `discovery.py`、`deep_observation.py`、`manual_observation.py` | 一次性浏览器探索路径 |
| `observer_evals.py`、`shadow_acceptance.py` | 评估/验收脚本 |
| `docs/校园培养方案解读（2026年版）.md` | 原始资料，非代码 |

---

## 3. 交叉验证清单

本仓库的缺陷**主要不是"代码写错"，而是"声明与实现不一致"**——
注释描述了 A、代码做的是 B，或文档把某模块归到了错误的层。
这类问题逐行阅读永远发现不了，因为每一行单看都成立。

下面两个检查可在仓库根直接执行。

### 3.1 已固化为命令的三条检查

```bash
uv run python tools/check_project.py    # 退出码 0 = 通过，1 = 有失败
```

覆盖三条**不需要人工判断**的客观规则：

| 检查 | 失败时会看到 |
| --- | --- |
| 内部依赖不成环 | `course_selection: categories -> notice -> categories` |
| 模块已在架构文档中登记 | `course_progress: 未登记 capture` |
| 敏感文件未入仓 | `被跟踪的敏感文件: probe.xlsx` |

该脚本已接入 `architecture.md` §8 的质量门。

### 3.2 应用核心不得依赖具体传输（未固化）

对应 `architecture.md` 第 4 节的第 1 条规则。它需要「哪条边界算架构规则」的
判断——例如 `urllib.parse` 是否算 IO、`lab_contract._default_get` 这类已知例外
如何记账——因此暂未并入 `tools/check_project.py`。期望输出 `OK: core imports no transport`。

```bash
python - <<'EOF'
import ast, os
CORE = {"workbench_service","planning","lab_booking","lab_contract","categories",
        "persistence","config","timetable","lab_ports","notice","notice_discovery"}
TRANSPORTS = {"gateway","lab_transport","lab_browser_session","notice_transport",
              "discovery","tasks","selection_entry","selection_query"}
bad = []
for f in sorted(os.listdir("course_selection")):
    if not f.endswith(".py") or f[:-3] not in CORE:
        continue
    tree = ast.parse(open(f"course_selection/{f}", encoding="utf-8").read())
    for node in ast.walk(tree):
        mod = node.module if isinstance(node, ast.ImportFrom) else None
        if mod and mod.split(".")[-1] in TRANSPORTS:
            bad.append(f"  {f}:{node.lineno} -> {mod}")
print("\n".join(bad) if bad else "OK: core imports no transport")
EOF
```

### 3.3 人工检查点

自动化查不到的部分，改动相关代码时顺手确认：

- 注释描述的比较/分支逻辑，与相邻那几行代码是否一致
  （`planning.py` 的 `_overlap` 曾经注释说按开始节比较、代码按真实区间比较）
- 模块的导入是否与其声称的职责相符
  （`notice.py` 曾自称解析模块，实际含 HTTP 与 Playwright）
- 新增的公开常量/类型是否放对了层
  （`lab_transport.py` 曾被登记为"应用核心"，但持有 HTTP/浏览器实现）
- `AGENTS.md` 的成熟度约定是否被遵守：`implemented` / `automated-test verified` /
  `real-environment verified` 三者不能混用

---

## 4. 有效的 review 动作

### 4.1 读测试的断言，不读实现

断言是**意图的可执行声明**，比 docstring 可信。判断"注释和代码谁对"时，
先找对应断言。

例：`tests/test_planning.py` 断言「周一 1-2 节 vs 周一 2-3 节应判为冲突」。
这条断言直接证明 `_overlap` 的代码是对的、注释是错的。

### 4.2 用脚本验证声明，而不是逐行读

45 个模块的依赖关系用 AST 扫描不到 1 分钟就能核完；人眼读做不到。
第 3 节的三个检查就是为此准备的。

### 4.3 专找"文档说 A、代码做 B"

这是 AI 辅助开发最高产的缺陷类型，也是唯一一类靠读单个文件发现不了的问题。
重点看：模块职责描述、注释里的算法说明、架构文档的层归属、docstring 的参数表。

### 4.4 测试数不是正确性

295 个测试证明**代码行为符合测试**，不证明**系统能与真实教务系统交互**。
真实环境验收状态见 `README.md` 的成熟度表。

---

## 5. 自测：你是否已经"持有"这个项目

能回答下面这些问题（不必读完全部代码），就足以安全地修改这个仓库：

1. `conflict_unknown` 为什么不等于"空闲"？
2. `book_slots` 遇到一次失败的提交后，为什么是 `break` 而不是重试？
3. 工作台的 CSRF token 从哪来，写操作还受什么约束？
4. 哪些数据被禁止落盘，由哪个函数保证？
5. 毕业进度为什么不能对外表述为"官方毕业判定"？
6. `HTTP 200` 与业务成功的关系是什么？
7. 改前端代码后，为什么必须重新 build 并提交产物？

答案分布在 `CONTEXT.md`、`AGENTS.md`、`architecture.md`
以及第 2 节 🔴 表格列出的文件里。

---

## 6. 相关文档

| 文档 | 内容 |
| --- | --- |
| `architecture.md` | 技术栈、模块地图、分层规则、质量门 |
| `CONTEXT.md` | 领域词汇与边界（冲突未知、认定学分、要求基线等） |
| `AGENTS.md` | 完成度约定、安全边界、agent 技能索引 |
| `README.md` | 功能成熟度表与使用方式 |
| `authentication.md` | 认证框架、冷启动时序、会话寿命与失败分类 |
| `adr/` | 7 条架构决策记录 |
