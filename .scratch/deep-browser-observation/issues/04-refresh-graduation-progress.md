# 04: 通过新观察接缝刷新毕业进度

**What to build:** 学生同步毕业进度时，系统复用同一个教务会话读取完整学期和分页，通过类型化结果发布不含具体成绩的毕业进度快照，并留下对应的教务请求轨迹。

**Blocked by:** 01/通过新观察接缝发布带轨迹的个人课表事实

**Status:** resolved

- [x] 毕业进度刷新通过新的观察接口运行，工作线程不接触 Playwright 或成绩页面实现细节。
- [x] 读取覆盖所有预期学期和分页，并在每个可取消节点检查任务状态。
- [x] 只有数据完整且符合要求基线版本的结果可以发布毕业进度快照。
- [x] 发布结果不包含账号凭据、Cookie、学生身份、具体成绩或原始页面内容。
- [x] 认证超时、分页失败、解析失败和数据不完整均保留已有完整快照，并留下刷新尝试诊断状态。
- [x] Replay 测试覆盖完整读取、会话恢复、分页失败、取消和数据不完整。

## Answer

Implemented typed graduation-progress observations with score-free payloads, baseline-version validation, per-attempt request traces, and publish-only-on-complete orchestration. Replay coverage proves complete publication and retention of the prior snapshot for cancellation, page failure, and unconfirmed contracts.
