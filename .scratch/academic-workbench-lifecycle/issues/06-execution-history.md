# 06: 保存脱敏执行历史并阻断未知结果重提

**What to build:** 用户可以查看脱敏选课执行历史；未知结果会阻断同一教学班再次提交，直到用户核实。

**Blocked by:** 05

**Status:** resolved

- [x] 保存时间、课程、rwh、类别、结果、提示和关联版本
- [x] 支持 selected、capacity_full、rejected、unknown
- [x] unknown 阻断同一 rwh，用户核实后可解除
- [x] 可清除历史，成功后不自动全量刷新

## Answer

Added redacted execution history, unknown-result blocking, resolution and clear-history workflows.
