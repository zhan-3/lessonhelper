# 04: 提供明确的强制刷新与快照状态界面

**What to build:** 用户能分别查看课表、待选课程和毕业进度状态，并明确触发所需刷新，不再使用含义混杂的同步操作。

**Blocked by:** 03

**Status:** resolved

- [x] 提供连接教务、刷新课表、刷新毕业进度、强制刷新待选课程、诊断监听和刷新本地页面
- [x] 活动任务期间禁用其他远程按钮
- [x] 强制刷新前显示范围和上次刷新时间并确认
- [x] 分别显示当前、历史、不完整、缺失及原因

## Answer

Replaced combined sync controls with explicit remote actions, busy disabling, force-refresh confirmation and per-snapshot status.
