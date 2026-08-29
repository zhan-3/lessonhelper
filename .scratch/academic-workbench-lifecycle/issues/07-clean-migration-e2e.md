# 07: 清理不兼容数据并验证完整工作流

**What to build:** 升级后保留高价值本地身份配置，清理不安全的旧业务数据，并验证从离线启动到单次选课的完整流程。

**Blocked by:** 02、04、06

**Status:** resolved

- [x] 保留学生画像、已确认通知、登录配置及至多一份课表参考
- [x] 删除旧待选课程、毕业进度、规划和任务记录
- [x] 首次使用明确提示强制刷新
- [x] 覆盖离线启动、认证、缓存、刷新、冲突、成功、容量已满和未知结果
- [x] 完整 Python、前端测试与人工验收清单通过

## Answer

Added version-3 cleanup preserving identity/notice/latest timetable, removed incompatible facts, and covered lifecycle behavior in tests.
