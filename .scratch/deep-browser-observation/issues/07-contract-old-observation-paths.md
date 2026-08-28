# 07: 收缩旧网关并验证唯一观察路径

**What to build:** 完成迁移后，所有工作台观察任务只经过一个深浏览器观察模块；旧生产导入仍可临时解析到新实现，但旧宽接口、重复读取路径和依赖私有实现的测试不再存在。

**Blocked by:** 03/将自动发现限制为诊断证据；04/通过新观察接缝刷新毕业进度；05/统一工作台连接、会话重置与浏览器生命周期；06/将手动观察迁移为纯诊断任务

**Status:** resolved

- [x] 所有生产观察任务只通过最小执行接口进入深浏览器观察模块。
- [x] 旧类名如仍被生产入口使用，只作为指向新适配器的兼容导出，不保留第二套实现。
- [x] 重复的待选课程班请求、分页和浏览器发现执行路径被删除。
- [x] 测试不再通过私有会话字段、未初始化实例或内部补丁路径驱动行为。
- [x] 纯个人课表和待选课程班解析测试继续独立运行。
- [x] 完整验证覆盖浏览器跨任务复用、自动发现不可发布、取消与认证超时、轨迹脱敏、轨迹失败降级及仅保留最近 20 份。
- [x] 用户文档说明工作台应在独立终端长期运行，以及关闭 Chromium、退出进程和重置登录的不同效果。


## Answer

Contracted workbench execution to typed observation methods, removed the duplicate discovery-time selection query/pagination implementation, and retained old gateway methods only as compatibility adapters on the unconfirmed base. Updated diagnostic behavior tests and long-running workbench lifecycle documentation; parser tests remain independent.
