# 03: 记录接口契约版本并检测漂移

**Status:** resolved

## 目标

把每个中心的接口契约固化成可 diff 的快照，页面或后端升级时能主动发现变化，
而不是等到取数失败。

## 可用的版本信号（已实测）

- 预约页外壳里的应用版本串：`/_app.config.js?v=2.8.0-1774602871364`
- 前端构建产物名：`assets/index.<hash>.js`
- 每中心的静态配置键：`VITE_GLOB_API_URL` 等
- 网关返回体固定信封：`{code, message, result, timestamp}`

## 需求

- [x] 定义契约快照：中心代码、通道、origin、应用版本串、端点清单、必需请求头名、
      请求体字段名、响应字段名（含 `data[].x` 点路径）、业务状态码取值、静态产物与配置键。
- [x] 提供 `lab-contract record` / `check` / `promote`：前两个只读，`promote` 只写本地基线与观测文件。
- [x] `check` 分级报告：`breaking` / `unavailable` / `additive` / `notice` / `cosmetic` / `rebaseline`。
- [x] 漂移只报告，不自动改写解析逻辑，也不调用写入端点。
- [x] 基线入仓 `docs/contracts/<channel>-<center>.json`（无个人数据）；本机观测留 `.private/`（Git 忽略）。
- [x] 人工锁定必需字段：`locked` 字段消失 = breaking，未锁定字段增删 = additive。
- [x] 读不到端点报 `unavailable` 且退出码非零，绝不报「无漂移」。
- [x] 通道或 origin 变化归入 `rebaseline`（换基线），避免换网络误报。
- [x] 环境块只在本机观测之间比较，基线不带环境。

## 已交付

```text
course_selection/lab_transport.py   纯 HTTP 主后端 + 页面内 fetch 回退 + 令牌获取
course_selection/lab_contract.py    快照、指纹、分级 diff、只读探针、锁定字段
course_selection/cli.py             lab-contract record|check|promote
docs/contracts/direct-dxwl.json     第一份基线（5 个端点，fingerprint b2b4e5a02610251d）
tests/test_lab_transport.py         传输层：无 Cookie、失败分级、回退只对传输失败生效
tests/test_lab_contract.py          契约核心与探针
tests/test_lab_contract_cli.py      CLI：退出码、基线缺失、promote 保留 locked
```

实测：`check` 对当前系统返回「无变化」、退出码 0；观测到 5 个端点、
应用版本 `2.8.0-1774602871364`、2 个构建产物。

## 边界

不抓取或保存响应正文中的个人数据；不把候选请求自动提升为正式契约；
不调用写入端点（`view/booking/doyyxkzw`）。
