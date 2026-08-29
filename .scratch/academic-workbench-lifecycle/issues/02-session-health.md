# 02: 以受保护业务页验证教务会话

**What to build:** 分别展示浏览器、WebVPN 和教务会话状态；只有固定只读业务页验证成功才显示教务可用。

**Blocked by:** 01: 统一活动任务状态与互斥控制

**Status:** resolved

- [x] loginCAS 只表示认证进行中
- [x] 固定受保护业务页成功后才标记教务可用并记录验证时间
- [x] 不做后台定时网络探测
- [x] 浏览器关闭后工作台继续运行并在下次操作重新启动

## Answer

Implemented protected-page session verification plus browser/WebVPN/academic status and last verification time.
