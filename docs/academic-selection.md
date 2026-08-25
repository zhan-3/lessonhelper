# 教务选课规划（第一阶段）

当前第一阶段实现本地学生画像、通知白名单、课表导入展示，以及通知匹配课程的只读查询。

启动本地页面：

```powershell
uv run python -m course_selection
```

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

数据只保存到 `.private/academic-selection/`。当前页面不会执行选课、退课或提交操作。

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

`discover-selection` 会在确认通知和学生画像后，直接向已验证的
`POST /xsxk/queryXsxkList` 发起只读查询，自动读取服务端分页并合并课程班、教师、
上课时间、容量等字段。结果保存在本次运行目录的 `selection-query.json`；不会调用
`saveXsxk`，也不会执行选课或退课。

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
