# 教务选课规划（第一阶段）

当前第一阶段实现本地学生画像、通知白名单、课表导入展示，以及通知匹配课程的只读查询。

启动本地页面：

```powershell
uv run python -m course_selection
```

## 开发时保持浏览器与实时监控

开发时使用以下命令。它启动一个长期持有 profile 的 Chromium，并在 Python 源码变化后只重启工作台进程；浏览器标签页、登录状态和 DevTools 连接不会重启：

```powershell
uv run python -m course_selection dev-workbench
```

DevTools/CDP 仅监听本机 `127.0.0.1:9222`，可在另一终端实时附加：

```powershell
playwright-cli -s=academic-live attach --cdp=http://127.0.0.1:9222
playwright-cli -s=academic-live tab-list
playwright-cli -s=academic-live snapshot
playwright-cli -s=academic-live console
playwright-cli -s=academic-live requests
```

开发监督进程只监控 `course_selection/` 与 `course_progress/` 下的 Python 文件。正在进行的教务任务完成前不会触发重启；按 `Ctrl+C` 会关闭工作台和 Chromium，但保留本地 profile 与凭据。

首次使用可配置统一身份认证自动登录：

```powershell
uv run python -m course_selection configure-login
```

命令在终端中读取一次学号和密码（密码输入不会显示），并使用 Windows
DPAPI 加密保存到 `.private/course-progress/webvpn-login.dpapi`。该文件只能由
当前 Windows 用户解密，不会写入 Git。认证成功后还会保存 Playwright 的
WebVPN 会话状态；后续优先直接复用，失效时才在
`webvpn.hitwh.edu.cn/.../authserver/login` 自动填写并提交。浏览器会话状态含
认证 Cookie，同样只保存在已忽略的 `.private/` 下。

打开 `http://127.0.0.1:5000/` 后：

1. 粘贴学校选课通知链接和正文，或直接录入通知内容。公开链接可以自动读取；需要 WebVPN/CAS 会话的链接请在浏览器完成登录后粘贴正文，避免把认证信息交给普通 HTTP 请求。
2. 检查学期、选课类型和起止时间，确认选课窗口。
3. 上传学校导出的 `.xls` 或 `.xlsx` 当前课表。
4. 检查课程、星期、节次和教学周是否正确；替换已有课表时需要显式勾选确认。

数据只保存到 `.private/academic-selection/`。待选课程快照会同时保存页面提供的 `saveXsxk` 教学班身份（`action_rwh`）、查询类别、学期和页码来源；只有存在该真实身份的课程才标记为 `execution_ready`，由课程代码和名称拼接的展示回退身份不可用于执行。

工作台允许对一个具体教学班执行一次选课：按钮只在当前课表冲突已知为无冲突、时间可解析且教学班具有真实 `rwh` 时启用。点击后还需确认，后端会验证快照、画像、通知、学期和白名单，随后用可见浏览器重新读取该类别并取得新表单令牌。每次任务最多提交一次，不会自动重试；Cookie、令牌和返回 HTML 不写入数据库。当前不提供退课操作。

配置学生画像（当前用户为 2025 级）：

```powershell
uv run python -m course_selection configure-profile --grade 2025
```

只有“已确认通知 + 学生画像年级”共同匹配出的选课类别会进入查询白名单；退课、申请和纸质/邮件办理事项会被排除。

探索已确认通知对应的教务选课入口：

```powershell
uv run python -m course_selection explore-entry
```

命令会复用 `.private/course-progress/` 的教务浏览器 profile。登录完成后，请在可见浏览器中手动打开通知对应的“学生选课”页面；程序只监听该页面产生的 Fetch/XHR JSON，并将脱敏结果保存到：

```text
.private/academic-selection/
├── selection-entry.json
└── selection-contracts.json
```

当前探索器不会点击选课、退课、保存或提交控件。

自动分析课表或选课入口：

```powershell
uv run python -m course_selection discover-timetable
uv run python -m course_selection discover-selection
```

两个命令默认启动 Playwright 自带 Chromium 的临时可见会话，并加载上次保存的
WebVPN 状态；传入 `--browser chrome` 才会显式使用系统 Chrome。

`discover-selection` 仅用于诊断已验证读取契约为何失效。它会记录脱敏的候选接口、
页面结构和点击证据，但不会把发现响应转换为待选课程班，也不会生成或替换教务快照。
正常刷新只使用工作台内的版本化只读契约。

维护者可以运行结构化的学生画像接口分析器：

```powershell
uv run python -m course_selection analyze-interface --target student-profile
```

该命令复用本机 DPAPI 登录和唯一的可见 Chromium，只允许 GET/HEAD/OPTIONS 与精确的
学校认证 POST。它验证 WebVPN 用户信息和门户目录能力，动态进入新教务系统，并在短暂
观察窗口内只记录候选响应的字段名和类型；不会保存姓名、学号等响应值，不修改 SQLite
或当前学生画像。结果位于 `.private/interface-analysis/student-profile/<时间>/`，超过
七天的旧分析目录会在下次运行时清理。普通工作台同步不会自动运行接口分析器。

工作台的普通“刷新选课班”使用版本化的 `hitwh-jwts-selection-query-v1` 只读契约：
浏览器只负责统一认证、Cookie 和动态查询表单，刷新会直接进入已验证页面并通过页面
同源接口查询通知白名单内的全部类别和分页，不再逐次点击教务菜单。只有已验证入口或
查询表单缺失时才进入受控接口发现流程；某一类别或分页的临时失败只记录为不完整刷新，
不会通过另一种路径重复查询，也不会替换上一次完整快照。

2026-08-25 真实页面验证得到的“学生选课”二级菜单与查询类别如下：

| 页面名称 | `pageXklb` |
| --- | --- |
| 英语 | `yy` |
| 体育 | `ty` |
| 文化素质核心（页面内显示“素质教育”） | `szhx` |
| 创新研修 | `cxyx` |
| 创新实验 | `cxsy` |
| 创新创业 | `cxcy` |
| 新生研讨 | `xsyt` |
| 未来技术学院课程 | `tsk` |
| 外专业课程 | `xsxk` |
| 微专业选课 | `wzy` |

WebVPN storage state 失效时，旧入口可能显示 EasyConnect `#!/login`。程序不会在
该页面填写统一认证密码，而是清理失效 Cookie，转到明确的 CAS 入口后再自动
填写。新教务系统自身还包含一个“统一身份认证登录”中间页，自动发现流程会先
进入该入口，再展开“学生选课”。

## 长期运行工作台

请在独立终端中运行 `uv run python -m course_selection workbench`，并保持该进程运行。工作台默认离线启动，只恢复本地画像、通知和教务快照；只有点击“连接教务”或某个明确的刷新按钮才会访问学校系统。“连接教务”仅验证教务会话，不会隐式刷新课表或待选课程。

待选课程以学生画像版本、学期、通知版本和查询契约版本为上下文保存在本地。上下文未变化时直接使用最近完整快照；“强制刷新待选课程”会在确认后读取白名单内全部类别和分页。每个上下文保留最近三份完整待选课程快照。选课执行前只重新读取目标类别和来源页，按真实 `rwh` 验证并获取新令牌。

工作台任一时刻只允许一个远程任务；相同请求会去重，其他请求会被拒绝而不是形成隐藏队列。界面分别显示浏览器、WebVPN、教务会话及各类快照状态。选课执行历史仅保存脱敏结果；未知结果会阻断相同教学班再次提交，直到用户核实解除。

升级到生命周期版时保留学生画像、已确认通知、登录配置和最近一份课表参考；清理旧待选课程、毕业进度、规划及遗留任务，首次使用需要强制刷新待选课程。
工作台进程内只维护一个可见 Chromium，连接、刷新、手动诊断和空闲轮询都会复用它：

- 关闭整个 Chromium 窗口不会结束本地工作台；会话状态变为未知，下一次明确的远程操作会重新启动浏览器，并从 `.private/` 的持久化 profile 尝试恢复仍有效的教务会话。
- 退出终端中的工作台进程会幂等关闭 Chromium，但不会删除持久化登录状态。
- 在页面中“重新配置”或清除登录会执行重置：先关闭 Chromium，再删除认证状态；删除失败时会阻止后续教务任务，避免继续使用不确定会话。
- 普通刷新失败、取消或认证等待超时不会关闭 Chromium，可以在同一窗口中处理认证后重试。
