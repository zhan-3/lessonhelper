# 03: 将自动发现限制为诊断证据

**What to build:** 当已验证的待选课程班读取规则失效时，系统可以自动发现读取路径并留下诊断材料，但发现结果只能用于诊断，绝不能成为待选课程班事实或教务快照。

**Blocked by:** 02/通过已验证读取刷新待选课程班

**Status:** resolved

- [x] 已验证读取规则失效时，刷新结果明确标记为接口未确认，并保留已有教务快照。
- [x] 自动发现返回仅含诊断证据的结果类型，不能携带可发布读取结果。
- [x] 工作线程只接受完整的已验证读取结果进行发布，诊断结果在类型结构上无法进入发布路径。
- [x] 自动发现复用当前持久化教务会话，不创建第二个 Chromium。
- [x] 自动发现遵守现有只读请求规则和查询白名单，不执行改变选课状态的请求。
- [x] 回归测试证明“自动发现成功但已验证读取失败”时不会发布教务快照。

## Answer

Implemented a distinct `SelectionDiscoveryDiagnostic` result. Contract discovery now ends in `interface_unconfirmed`, stores diagnostic evidence, and cannot publish or replace a selection snapshot. Added Replay regression coverage; existing session reuse and read-only route guards remain in the discovery adapter.
