# 04: 补充课程标签并核验外专业课程体系

**What to build:** 当学校课程记录不能确认 D 类、四史或外专业课程体系时，学生可以在本地为具体课程补充标签并选择一个外专业课程体系。补充内容保留为用户申报，只影响预计子约束和预计体系进度，不改写教务事实或重复增加课程学分。

**Blocked by:** 02/同步并展示新版已确认毕业进度

**Status:** resolved

- [x] 用户可以按稳定课程身份查看并补充 D 类、四史和外专业体系标签。
- [x] 用户可以选择当前基础毕业要求采用的唯一外专业课程体系，并可显式修改或清除选择。
- [x] 手动标签和体系归属与学校原始课程字段分开保存，界面明确标记为用户申报。
- [x] 已通过课程增加 D 类或四史标签后，课程总学分不变，只产生对应子约束的预计贡献。
- [x] 四史门数按课程身份去重；无对应已完成课程的自由文本不能直接作为已确认门数。
- [x] 外专业总额只累计所选体系内的明确课程；其他体系课程不计入，体系未知课程保持可见并形成待核验项。
- [x] 未选择体系、课程体系未知或标签依据缺失时，必要条件保持未知而不是已满足或零完成。
- [x] 不使用课程名称关键词自动推断 D 类、四史或体系。
- [x] 标签和体系选择在重启后恢复，清除个人工作区时一并删除，不会跨账号继承。
- [x] 本地修改不触发成绩、课表或待选课程刷新，也不访问学校系统。
- [x] HTTP、计算和界面测试覆盖标签增删、课程去重、体系切换、未知状态、重启、重置和无远程副作用。

## Answer

Implemented baseline-scoped local course labels for D-category, Four Histories, and outside-major track membership, plus one explicit selected outside-major track. Labels are accepted only for stable identities found in the current complete progress report, persist separately from school facts, and are removed by personal-workspace reset. D/Four-Histories labels change only estimated subconstraints while parent cultural credits remain unchanged; Four Histories deduplicates by course identity. Outside-major estimates count only courses explicitly assigned to the selected track and keep unknown/other-track courses visible. All mutations are local, CSRF-protected, and do not create academic tasks or remote access.
