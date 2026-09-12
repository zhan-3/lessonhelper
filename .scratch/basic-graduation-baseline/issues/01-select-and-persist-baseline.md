# 01: 选择并持久化基础毕业要求基线

**What to build:** 学生可以在工作台查看基础毕业要求基线的规则、来源、适用边界和参考性质，并主动选择 `basic-graduation-reference-v1`。选择只改变本地计算上下文，不访问学校系统；原有报告继续作为历史记录，新版本在重新同步前不会借用旧版本结论。

**Blocked by:** None (can start immediately)

**Status:** resolved

- [x] 工作台列出既有基线和 `basic-graduation-reference-v1`，并展示版本、来源、覆盖范围、人工补充规则及非正式声明。
- [x] 新基线包含本专业选修、创新创业、社会实践、创新创业与社会实践合计、文化素质、D 类、四史及外专业课程体系要求。
- [x] 用户必须主动确认基线选择；安装、升级、年级推导和页面加载均不会自动选择或切换版本。
- [x] 基线选择持久保存在本地唯一事实来源中，重启工作台后保持不变。
- [x] 选择或切换基线不创建远程任务、不启动认证，也不访问学校系统。
- [x] 旧基线报告保持内容不变并标记为历史；新基线显示需要重新同步，而不是重新解释旧报告。
- [x] 基线定义与类别映射共同版本化且不可由用户编辑；后续规则变化需要新版本。
- [x] 升级现有工作区时保留画像、通知、快照、规划和任务，迁移失败不会留下半选择状态。
- [x] HTTP 行为测试覆盖选择、重复选择、重启恢复、历史状态、CSRF/同源保护和无远程副作用。

## Answer

Implemented explicit local selection of immutable, source-labelled requirement baselines. The selected version survives restart, resets with the personal workspace, never starts a remote task, and makes mismatched reports inspectable history pending an explicit refresh. Baseline definitions own their category mappings, the generated API models the unselected state as nullable, and HTTP plus frontend interaction tests cover confirmation, persistence, upgrade preservation, atomic failure, and historical disclosure.
