# openlab_cas_book.py

HITWH 开放式实验系统（大学物理实验教学中心）抢课自动化工具。

Playwright 浏览器自动化：CAS SSO 登录 → 会话持久化 → 自动预约 → 监控模式。

## 快速开始

```bash
# 安装依赖
pip install playwright
playwright install chromium

# 首次运行（需要在浏览器中手动登录 CAS）
python openlab_cas_book.py

# 后续运行自动复用已保存会话，无需再次登录
```

## 用法

| 命令 | 说明 |
|---|---|
| `python openlab_cas_book.py` | 单次尝试所有目标课程 |
| `python openlab_cas_book.py --monitor` | 监控模式，每 5 秒检查一轮 |
| `python openlab_cas_book.py --monitor --interval=3` | 自定义间隔（秒） |
| `python openlab_cas_book.py --course="DIY电磁混合磁悬浮"` | 只预约指定课程 |
| `python openlab_cas_book.py --login-only` | 仅登录保存会话，不预约 |

## 工作流程

```
首次运行                   后续运行
───────                   ───────
打开浏览器                加载 storage_state.json
  │                         │
跳转 CAS 登录页            检查会话是否有效
  │                         ├─ 有效 → 直接操作预约页
手动输入学号/密码            └─ 过期 → 走首次流程
  │
登录成功 → 保存会话
  │
进入预约页面 → 按优先级尝试课程
  ├─ 有名额 → 立即预约
  └─ 满员 → 下一门 / 等待下一轮
```

- 首次登录只需执行一次，后续自动复用会话
- 登录不需要验证码
- 支持 `--monitor` 监控模式，适合蹲守名额释放

## 目标课程（按优先级）

| 课程 | subjectId | 状态 |
|---|---|---|
| DIY电磁混合磁悬浮 | 1007 | 待预约 |
| 磁阻效应 | 1011 | 待预约 |
| 表面张力系数测定 | 1008 | 待预约 |

## 项目结构

```
├── openlab_cas_book.py    # 核心脚本
├── pyproject.toml         # 项目配置
├── README.md              # 本文件
├── docs/DEVLOG.md         # 开发日志
├── storage_state.json     # 浏览器会话（自动生成，不提交）
└── .gitignore
```

## 依赖

- Python >= 3.10
- [Playwright](https://playwright.dev/python/)（浏览器自动化）

## 开发

```bash
# 使用 uv（推荐）
uv sync
uv run openlab_cas_book.py
```

## 已知局限

- 预约页 SPA 弹窗结构（选日期→选座位→确认）尚未经实机验证，需等待选课期开放后调试
- 部分旧版本 Chrome 可能不兼容，建议使用 Chrome 125+
