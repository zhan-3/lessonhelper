# 02: 同步并展示新版已确认毕业进度

**What to build:** 学生选择基础毕业要求基线并显式同步成绩后，可以在工作台看到由学校成绩记录确认的基础毕业进度。全部要求固定呈现，合计、子约束、缺口和无法确认的原因使用统一计算口径，不会因待选课程查询范围而隐藏。

**Blocked by:** 01/选择并持久化基础毕业要求基线

**Status:** resolved

- [x] 显式同步继续使用既有固定成绩读取契约，完整读取所有学期和分页后才发布新版进度快照。
- [x] 新版报告绑定 `basic-graduation-reference-v1`，并包含每项要求的最低值、已确认贡献、已确认缺口、条件状态、原因和贡献明细。
- [x] 本专业选修至少 3、创新创业至少 4、社会实践至少 1、两者合计至少 6、文化素质至少 8、D 类至少 2、四史至少 1 门、外专业至少 10 及单一体系约束均被表达。
- [x] 创新创业 4 学分明确标为人工补充参考规则，不宣称为所有年级统一正式要求。
- [x] 合计和单项分别判断；创新创业 4 加社会实践 1 不会被判定满足合计 6。
- [x] 一门文化素质课程可贡献 D 类或四史子约束，但在文化素质总量中只累计一次。
- [x] 课程按课程身份去重，冲突事实不计入已确认完成，未知类别保留为待归类课程。
- [x] 不通过课程名关键词猜测 D 类、四史或外专业体系；缺少必要证据时对应条件为未知。
- [x] 已满足、未满足、未知由证据决定；总额足够但必要子约束未知时不显示整个要求已满足。
- [x] 成绩分页完整不等同于认定来源完整，报告保留活动认定和体系信息的数据覆盖缺口。
- [x] 全部基线要求在界面固定展示，与通知白名单和待选课程查询类别无关。
- [x] 同步失败或不完整时不发布新快照，保留上一份完整报告及可检查的失败状态。
- [x] 计算边界、组合要求、去重、未知状态和完整发布均有外部行为测试。

## Answer

Implemented explicit confirmed-progress synchronization under the selected immutable baseline. The fixed grade reader remains unchanged; only complete collections publish. Reports contain all eight requirements, confirmed amounts and gaps, evidence-driven three-state results, backend-authored explanations, coverage gaps, conflicts, and unclassified courses. Explicit multi-target classification can satisfy cultural subconstraints while the parent total deduplicates by course identity. HTTP and frontend tests cover exact thresholds and units, `4 + 1`, unknown reasons, query-independent rendering, and preservation of the previous complete snapshot after incomplete synchronization.
