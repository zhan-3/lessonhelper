# openlab_cas_book.py

HITWH 开放式实验系统抢课自动化工具。

Playwright 浏览器自动化：CAS SSO 登录 → 自动预约 → 监控模式。

## 用法

```bash
# 首次运行（需手动输入验证码登录）
python openlab_cas_book.py

# 监控模式，每 3 秒检查一次
python openlab_cas_book.py --monitor --interval=3

# 仅登录保存会话
python openlab_cas_book.py --login-only

# 指定单个课程
python openlab_cas_book.py --course="DIY电磁混合磁悬浮"
```

## 流程

1. **首次运行**：打开浏览器 → HIT CAS SSO 登录页 → 手动输账号/密码/验证码 → 登录成功 → 保存会话
2. **后续运行**：复用已保存会话 → 自动操作预约页面
3. **监控模式**：每隔 N 秒检查各目标课程，有名额立即预约

## 依赖

- Python >= 3.10
- Playwright (浏览器自动化)

```bash
pip install playwright
playwright install chromium
``` 
