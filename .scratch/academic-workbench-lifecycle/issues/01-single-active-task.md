# 01: 统一活动任务状态与互斥控制

**What to build:** 工作台始终只存在一个可见的远程活动任务；重复点击不会产生隐藏队列，任务状态、阶段、运行时间和超时原因在页面与 API 中一致。

**Blocked by:** None (can start immediately)

**Status:** resolved

- [x] 相同查询去重，其他并发远程任务被拒绝
- [x] 读取、认证、完整刷新和执行具有明确超时
- [x] 取消、超时和重启均进入终态并释放互斥
- [x] 执行进入提交阶段后不可取消

## Answer

Implemented one-active-task rejection/coalescing, active task API/UI, restart terminalization and execution cancellation boundary.
