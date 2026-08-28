# 05: 统一工作台连接、会话重置与浏览器生命周期

**What to build:** 学生启动并持续使用工作台时，工作台连接、认证等待、连续读取、轮询和重置登录由一个深浏览器观察模块管理，使一个 Chromium 在工作台进程生命周期内稳定复用。

**Blocked by:** 02/通过已验证读取刷新待选课程班；03/将自动发现限制为诊断证据

**Status:** resolved

- [x] 工作台启动、连接并读取、轮询和重置会话均通过同一个观察接缝运行。
- [x] 连续执行连接、课表刷新和待选课程班刷新只启动一个 Chromium。
- [x] 普通任务成功、失败、取消和认证超时均不会关闭 Chromium。
- [x] 等待认证状态在阻塞等待前可被任务检查方观察，认证完成后同一任务能够继续。
- [x] 明确重置登录会先关闭 Chromium，再删除认证状态；删除失败会阻止后续任务继续使用不确定状态。
- [x] 工作台进程退出或浏览器发生不可恢复故障时，浏览器资源被幂等关闭。
- [x] 关闭可见 Chromium 会结束工作台；重新运行工作台可继续使用持久化配置目录中仍有效的本地会话状态。

## Answer

Unified connect-time timetable and selection reads behind typed observation results, merged their request traces, and retained one worker-owned gateway/browser across shell launch, connection, and consecutive refreshes. Browser relaunch now resets lifecycle notification flags; existing reset, timeout, shutdown, and visible-browser-close behavior remains covered by the lifecycle boundary.
