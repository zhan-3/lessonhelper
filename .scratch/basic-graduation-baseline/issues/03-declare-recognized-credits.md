# 03: 申报成绩接口未覆盖的认定学分

**What to build:** 学生可以在工作台本地申报成绩接口无法确认的认定学分，并随时修改或删除。申报清楚标记为用户申报，只参与预计完成和预计缺口，不会伪装成学校确认，也不会触发任何远程访问。

**Blocked by:** 02/同步并展示新版已确认毕业进度

**Status:** resolved

- [x] 用户可以创建包含稳定身份、受支持类别、正数有限学分、说明和日期的认定学分申报。
- [x] 用户可以修改和删除申报，变更后预计贡献立即重新计算。
- [x] 拒绝零、负数、非有限学分、非法日期、不支持类别及超出合理边界的输入。
- [x] 重复创建请求具有幂等行为，不产生两条相同身份的申报。
- [x] 申报持久保存在本地唯一事实来源中，工作台重启后可以恢复。
- [x] 申报不写入学校快照、不覆盖课程事实、不提高已确认完成量，也不消除已确认数据覆盖缺口。
- [x] 页面分别展示已确认缺口与计入申报后的预计缺口；预计达到要求不显示为已确认满足。
- [x] 与已有课程或贡献明确关联的申报不会重复增加总额；无法确定重叠关系时显示人工核验提示，不做模糊名称合并。
- [x] 不上传或保存证书图片，不要求浏览器认证，不访问学校系统。
- [x] 清除个人工作区时删除个人申报，但不删除公共基线定义和按既有约定保留的通知。
- [x] 本地写操作继续通过同源和 CSRF 保护，日志及响应不新增凭据、Cookie 或具体成绩。
- [x] HTTP、持久化和界面交互测试覆盖创建、编辑、删除、校验、幂等、重启、重置和无远程副作用。

## Answer

Implemented baseline-scoped local recognized-credit declarations for innovation, social practice, and cultural quality. The workbench supports create, edit, and delete with bounded validation and idempotent identities. Declarations affect only estimated amounts and gaps; confirmed school facts, condition states, coverage gaps, and stored school snapshots remain unchanged. Explicit links prevent double counting against confirmed, enrolled, queued, or duplicate declaration contributions, while unverified overlap remains visibly marked for manual review. Data persists in SQLite, is removed by personal-workspace reset, and all mutation APIs retain same-origin and CSRF protection.
