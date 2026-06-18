# DEVLOG - 开放式实验系统抢课自动化

## Phase 1: 初始探索 (2026-03-27)

### 发现
- 目标系统: `openlab.hitwh.edu.cn` — 哈尔滨工业大学(威海) 开放式实验室预约系统
- 旧脚本 `auto_book.py` / `quick_book.py` 基于 requests 直接调 API
- 猜测的 API 路径: `/dxwl/booking/api/`
- 首次提交 `29c2eaa` 建立项目骨架

### 文件
- `auto_book.py` — requests 版自动预约（假设 RESTful JSON API）
- `quick_book.py` — requests 版快速预约变体
- `pyproject.toml` — 初始依赖（playwright, requests, openpyxl, pandas 等）

---

## Phase 2: 真实系统探测 (2026-04-12)

### Playwright 实测发现
- 用 Playwright 真实探测网站，发现 SPA 配置 `__PRODUCTION__LSM__CONF__`
- **真实 API 基础路径**: `/dxwl/StuApi/`（不是旧脚本猜的 `/dxwl/booking/api/`）
- **认证方式**: HIT CAS SSO（`ids.hit.edu.cn`），回调 `lmsAuthApi/auth/cas/loginSuccess`
- SPA 技术栈: Vue + Ant Design
- 旧 API 路径全部错误，旧脚本不可用

### 决策
- 废弃所有 requests 直调方案
- 转向 Playwright 浏览器自动化方案
- Java 后端方案（`13.course-helper-server`）暂停，等 Python 方案验证通过再说

### commit: `1f05de7` — feat: add lab booking automation scripts

---

## Phase 3: 代码精简 (2026-04-12)

### 操作
- 删除废弃代码:
  - `auto_book.py` / `quick_book.py`
  - `grab_system/` 整个目录
  - `新一代选课助手/` 整个目录（FastAPI 后端 + 微信小程序 + 配置文件，~38 文件/3700 行）
- 编写 `openlab_cas_book.py`（348 行）：CAS SSO 登录 + 会话持久化 + 自动预约 + 监控模式

### 核心设计
- 首次运行: Playwright 打开浏览器 → CAS 登录页 → 用户手动输账号密码 → 登录成功 → 保存 storage_state.json
- 后续运行: 加载已保存会话 → 自动操作预约页面
- `--monitor` 模式: 每隔 N 秒检查，有名额立即抢
- 目标课程优先级: DIY电磁混合磁悬浮 > 磁阻效应 > 表面张力 > 偏振光

---

## Phase 4: 反检测增强 (2026-06-18)

### 改进
- 新增随机 UA 池（4 个 Chrome 版本）
- 新增操作抖动延迟 `jitter()`（0.3~1.5s），模拟人类节奏
- 新增 `_build_context()` 拦截图片/字体/媒体资源，提速并减少浏览器特征
- 修复页面泄漏: `do_booking()` 在全部退出路径关闭页面
- 修复监控模式每轮每课程泄漏一个页面的 bug
- 移除 `requests`、`pandas`、`openpyxl` 等无用依赖
- 更新项目描述为: "HITWH 开放式实验系统抢课自动化工具"

### commit: `78c3209` — refactor: consolidate to openlab_cas_book.py with anti-detection & page leak fixes

---

## Phase 5: 真实 DOM 探测 (2026-06-18)

### Playwright 实测发现

**确认真实信息:**
| 项目 | 真实值 |
|---|---|
| 用户 | 张浩翔 / 2025211052 |
| SPA 框架 | Ant Design Vue（`ant-table`, `ant-btn`, `ant-layout` 等） |
| API Base | `/dxwl/StuApi/` ✅ |
| 认证 | CAS SSO，**无验证码**，只需学号+密码 |
| 预约页 URL | `/dxwl/booking/#/booking`（**旧脚本猜对了**） |
| 课程列表 API | `POST /dxwl/StuApi/view/subjects` → 返回 22 门课含 subjectId |
| 用户信息 API | `POST /dxwl/StuApi/auth/currentUserInfo` |
| 预约查询 API | `POST /dxwl/StuApi/view/booking/yyxh`（传 `id=subjectId`） |

**课程实际情况（从 API 拉取）:**
- DIY电磁混合磁浮 (1007): 未通过，可预约
- 磁阻效应 (1011): 未通过，可预约
- 表面张力系数测定 (1008): 未通过，可预约
- 偏振光 (1015): **已通过**，无需预约
- 弦振动 / 温度传感器 / 新型电机: **系统中不存在**

**旧脚本的误判:**
- CAS 登录假设有验证码 → 实际没有
- API 路径 `/dxwl/booking/api/` → 真实为 `/dxwl/StuApi/`
- 部分目标课程不存在或已通过

### 保存会话
`storage_state.json` 已生成，后续运行脚本无需手动登录。

### 当前阻塞
预约查询 API 返回 `暂无相关信息，可能实验周期已过`（code: 5000），非选课期，预约页弹窗结构无法验证。`do_booking()` 中的交互逻辑（选课程→选日期→确认）需等选课开放后实机测试调整。

---

## 架构回顾

```
openlab_cas_book.py     ← 核心脚本（CAS登录 + 自动预约 + 监控）
pyproject.toml          ← 项目配置（仅依赖 playwright）
README.md               ← 使用说明
docs/DEVLOG.md          ← 开发日志（本文件）

storage_state.json      ← Playwright 会话（运行时生成，gitignored）
```

### 关键决策记录
1. **2026-04-12**: 废弃所有 requests 方案，转向 Playwright 浏览器自动化
2. **2026-04-12**: 废弃 FastAPI 后端 + 微信小程序方案，只保留单一 Python 脚本
3. **2026-04-12**: Java 项目暂停，等 Python 方案验证通过后决定是否用 Java 重写
4. **2026-06-18**: 确认 CAS 无验证码，简化登录流程设计
