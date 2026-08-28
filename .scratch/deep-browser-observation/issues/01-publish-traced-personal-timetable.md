# 01: 通过新观察接缝发布带轨迹的个人课表事实

**What to build:** 学生刷新个人课表时，读取应通过新的深浏览器观察模块完成。完整的个人课表事实可以发布为教务快照；任务同时留下可定位的教务请求轨迹，并能用 Replay 适配器确定性验证整个路径。

**Blocked by:** None (can start immediately)

**Status:** resolved

- [x] 工作线程通过最小观察接口提交类型化个人课表刷新请求，不接触 Playwright 实现细节。
- [x] 只有类型化的完整读取结果能够发布个人课表事实；取消、不完整读取和错误均保留已有教务快照。
- [x] 每次刷新生成按实际观察顺序排列的脱敏教务请求轨迹，覆盖导航、document、xhr 和 fetch。
- [x] 轨迹只记录顺序、时间、方法、脱敏 URL、资源类型和脱敏字段名，不保存认证信息、字段值或完整请求内容。
- [x] 轨迹写入或清理失败不会阻断课表刷新，但任务详情会标明轨迹不完整。
- [x] 每个任务一份 JSON Lines 轨迹，且只保留最近 20 份已结束轨迹。
- [x] Replay 适配器覆盖完整发布、取消、不完整读取、轨迹失败和不发布旧快照替代结果。

## Answer

Implemented the typed timetable-observation seam, sanitized per-task JSONL trace retention, and deterministic Replay coverage. Complete observations publish snapshots; all other outcomes retain the prior snapshot.
