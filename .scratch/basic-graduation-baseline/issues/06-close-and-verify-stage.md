# 06: 完成升级、回归与阶段验收

**What to build:** 学生从现有工作区升级后，可以稳定使用基础毕业参考进度，旧数据和已验收的工作台能力不受破坏。维护者获得完整的自动化与人工核对证据，并明确记录基础毕业参考、完整学校审核和暂缓抢课之间的边界。

**Blocked by:** 05/统一本学期、队列和申报的预计贡献

**Status:** resolved

- [x] 从现有数据库升级时保留画像、凭据状态、确认通知、学校快照、规划和任务，基线选择、申报及标签迁移具有原子性和可恢复性。
- [x] 旧基线报告仍可查看且内容未被重写；新版报告、历史状态和需要同步状态在重启后保持一致。
- [x] 账号重新配置或个人工作区重置不会把基线选择、申报、标签或个人进度带给另一身份。
- [x] 离线启动、连接、课表、待选课程、毕业进度刷新、通知、快照保留和只读规划的既有回归继续通过。
- [x] 新增计算器、HTTP、持久化、前端交互、迁移、安全和失败恢复测试全部通过。
- [x] 前端类型检查、生产构建和项目静态检查通过，生成资源与源码一致。
- [x] 使用合成数据完成人工核对：低于、等于和高于边界，组合要求，D 类，四史，外专业体系，用户申报及未知状态均符合规格。
- [x] 用户文档说明基线版本、参考来源、申报方式、三态含义、重新同步行为及本地隐私边界。
- [x] 页面与文档明确声明结果不含专业必修及完整培养方案审核，不是学校正式毕业结论。
- [x] 本阶段不执行任何真实选课或预约；抢课因窗口未开放继续记录为待验收风险。
- [x] 全量测试中进入本阶段前已有的无关失败单独报告，不以忽略测试或无关修改制造通过结论。
- [x] 完成报告分别标明已实现、自动测试验证、用户人工核对和仍待真实环境验证的能力。

## Completion report

- **Implemented:** fixed baseline selection, confirmed progress sync, local declarations/labels/track selection, and unified local projection.
- **Automated-test verified:** calculator, HTTP, persistence/migration, reset/restart, frontend interaction, generated types, build, lint, and failure recovery.
- **Human-reviewed with synthetic data:** boundary values, combined requirements, D category, Four Histories, selected track, declarations, deduplication, and unknown states.
- **Real-environment verified:** none for the new graduation baseline; live academic reads still require acceptance.
- **Release risks:** graduation output remains advisory; full major-required-course auditing is excluded. Course-selection submission remains paused and has not passed live acceptance.
