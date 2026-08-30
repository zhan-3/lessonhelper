# 05: 执行前仅验证目标教学班

**What to build:** 用户确认具体教学班后，只重新读取目标类别、来源页和 rwh，取得新 token 并单次提交。

**Blocked by:** 01、03

**Status:** resolved

- [x] 仅适用快照中的真实 action_rwh 可执行
- [x] 重新验证教学班存在及冲突状态
- [x] 每次确认绑定一个 rwh，单次提交且不自动重试
- [x] 不持久化 Cookie、token 或原始 HTML

## Answer

Retained targeted category/page/rwh revalidation with fresh form state and single-submit execution.
