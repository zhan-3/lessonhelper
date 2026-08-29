# 03: 离线启动并复用适用教务快照

**What to build:** 启动只读取本地数据；待选课程按画像、学期、通知及契约版本复用，不自动认证或刷新。

**Blocked by:** 01、02

**Status:** resolved

- [x] 启动不产生学校网络请求
- [x] 连接教务只验证会话
- [x] 上下文一致时复用完整快照
- [x] 上下文变化后旧快照仅供查看
- [x] 每个上下文保留最近三份完整快照和最近失败诊断

## Answer

Removed startup discovery/authentication and combined connect refresh; added context-based snapshot status and three-selection retention.
