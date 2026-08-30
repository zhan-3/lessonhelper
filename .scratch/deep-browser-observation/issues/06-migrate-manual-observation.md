# 06: 将手动观察迁移为纯诊断任务

**What to build:** 学生开始手动观察后，可以在同一个持久化教务会话中导航并按需结束任务；系统记录脱敏诊断证据和教务请求轨迹，但不产生或替换任何教务快照。

**Blocked by:** 01/通过新观察接缝发布带轨迹的个人课表事实

**Status:** resolved

- [x] 手动观察通过类型化诊断请求运行，并复用现有 Chromium 与教务会话。
- [x] 用户完成、取消、超时和浏览器关闭都有明确且可检查的终止结果。
- [x] 手动观察只记录允许的页面导航与请求结构，继续拦截不允许的写请求。
- [x] 诊断结果与可发布读取结果使用互斥类型，无法进入教务快照发布路径。
- [x] 教务请求轨迹按任务记录全局观察顺序，并在任务结束时移除监听器、完成文件和执行保留策略。
- [x] 测试证明监听器不会在连续任务之间重复注册或泄漏事件。

## Answer

Migrated manual navigation to `ManualObservationRequest` and diagnostic-only `ManualObservationResult`, with explicit completion, cancellation, timeout, and browser-close statuses. Requests are sanitized into the shared per-task trace store, no snapshot publication path exists, and route/response handlers are removed after every task without accumulating across consecutive observations.
