import React, { useEffect, useMemo, useRef, useState } from "react";
import type { WorkbenchState } from "./api";
import {
  activeInWeek,
  candidateOption,
  compressWeeks,
  courseAlreadyCompleted,
  deriveCurrentWeek,
  describeOptionMeetings,
  expandScheduleItems,
  formatWeekGroup,
  locationsByWeek,
  planningFilterFor,
  progressFilterByKey,
  progressKeysByQueryCode,
  projectedCourseCredits,
  transitionsIntoGroup,
  weekItems,
  type CandidateOption,
  type RequirementFilter,
  type ScheduleDiff,
  type ScheduleItem,
  type WeekCalibration,
} from "./schedule";

const weekdays = ["一", "二", "三", "四", "五", "六", "日"];
const sessionTimes = ["08:00–09:45", "10:05–11:50", "14:00–15:45", "16:05–17:50", "18:40–20:25", "20:45–22:30"];
const pageSize = 12;
const maximumPreviews = 3;
const formatCredits = (credits: number) => Number.isInteger(credits) ? String(credits) : String(credits).replace(/0+$/, "").replace(/\.$/, "");

const timeOverlaps = (a: ScheduleItem, b: ScheduleItem) => a.day !== null && a.day === b.day && a.start !== null && b.start !== null && a.end !== null && b.end !== null && a.start <= b.end && b.start <= a.end;
const executionTimeOverlaps = (a: ScheduleItem, b: ScheduleItem) => timeOverlaps(a, b) &&
  (!a.weeks.length || !b.weeks.length || a.weeks.some(week => b.weeks.includes(week)));
const formatLocalDate = (date: Date) => {
  const year = date.getFullYear();
  const month = String(date.getMonth() + 1).padStart(2, "0");
  const day = String(date.getDate()).padStart(2, "0");
  return `${year}-${month}-${day}`;
};

const readCalibration = (term: string): WeekCalibration | null => {
  try {
    const parsed = JSON.parse(window.localStorage.getItem(`workbench-current-week:${term}`) ?? "null") as WeekCalibration | null;
    return parsed?.term === term ? parsed : null;
  } catch {
    return null;
  }
};

const diffCount = (diff: ScheduleDiff) => diff.added.length + diff.ended.length + diff.timeAdjusted.length + diff.teacherAdjusted.length;

export function ScheduleBoard({ timetable, selection, graduationProgress, onExecuteSection, executionPending }: {
  timetable: WorkbenchState["snapshots"]["timetable"];
  selection: WorkbenchState["snapshots"]["selection"];
  graduationProgress?: WorkbenchState["graduation_progress"];
  onExecuteSection?: (sectionId: string, courseName: string) => void;
  executionPending?: boolean;
}) {
  const term = timetable?.term || selection?.term || "当前学期";
  const [selectedWeek, setSelectedWeek] = useState(1);
  const [query, setQuery] = useState("");
  const [requirementFilter, setRequirementFilter] = useState("全部");
  const [conflictFilter, setConflictFilter] = useState("全部");
  const [page, setPage] = useState(1);
  const [previewKeys, setPreviewKeys] = useState<string[]>([]);
  const [calibrationRevision, setCalibrationRevision] = useState(0);
  const [calibrationOpen, setCalibrationOpen] = useState(false);
  const [calibrationWeek, setCalibrationWeek] = useState(1);
  const [diffOpen, setDiffOpen] = useState(false);
  const [locationDetailKey, setLocationDetailKey] = useState<string | null>(null);
  const groupButtons = useRef(new Map<string, HTMLButtonElement>());

  const current = useMemo(() => ((timetable?.payload.entries as Record<string, unknown>[] | undefined) ?? []).flatMap((item, index) => expandScheduleItems(item, "current", index)), [timetable]);
  const candidateOptions = useMemo(() => ((selection?.payload.sections as Record<string, unknown>[] | undefined) ?? []).map(candidateOption), [selection]);
  // 筛选选项由实际数据推导：没有对应课程的培养要求不再出现在选项栏。
  const requirementFilterOptions = useMemo(() => {
    const present = new Set<string>();
    for (const option of candidateOptions) {
      const value = planningFilterFor(option);
      if (value) present.add(value);
    }
    return ["全部", ...[...present].sort((a, b) => a.localeCompare(b, "zh-CN"))];
  }, [candidateOptions]);
  const allMeetings = useMemo(() => [...current, ...candidateOptions.flatMap(option => option.meetings)], [current, candidateOptions]);
  const maxWeek = Math.max(1, ...allMeetings.flatMap(item => item.weeks));
  const calibration = useMemo(() => readCalibration(term), [term, calibrationRevision]);
  const currentWeek = deriveCurrentWeek(calibration, maxWeek);
  const progressItems = graduationProgress?.report?.progress ?? [];
  const completedCourses = progressItems.flatMap(item => item.courses).filter((course, index, courses) =>
    courses.findIndex(other => other.code === course.code && other.name === course.name) === index
  );
  const culturalProgress = progressItems.find(item => item.key === "cultural_quality");
  const fourHistoriesComplete = culturalProgress?.courses.some(course => /四史/.test(course.name)) ?? false;

  useEffect(() => {
    setPreviewKeys([]);
    const nextCalibration = readCalibration(term);
    const nextCurrentWeek = deriveCurrentWeek(nextCalibration, maxWeek);
    setSelectedWeek(nextCurrentWeek ?? 1);
    setCalibrationWeek(nextCurrentWeek ?? 1);
    setCalibrationOpen(!nextCalibration || nextCurrentWeek === null);
  }, [term, maxWeek]);

  const previewOptions = candidateOptions.filter(option => previewKeys.includes(option.key));
  // 进度面板与待选课程同步：只展示本次选课查询覆盖的培养要求。
  const syncedProgress = (() => {
    const queries = (selection?.payload.queries as Array<{ category?: unknown }> | undefined) ?? [];
    const keys = new Set(queries.flatMap(item => progressKeysByQueryCode[String(item.category ?? "")] ?? []));
    return progressItems.filter(item => keys.has(item.key));
  })();
  const scheduleItems = [...current, ...previewOptions.flatMap(option => option.meetings)];
  const weekGroups = compressWeeks(scheduleItems, maxWeek);
  const selectedGroup = weekGroups.find(group => group.weeks.includes(selectedWeek)) ?? weekGroups[0];
  const visibleCurrent = weekItems(current, selectedWeek);
  const visiblePreviews = weekItems(previewOptions.flatMap(option => option.meetings), selectedWeek);
  const boardItems = [...visibleCurrent, ...visiblePreviews];
  const sessionCount = Math.max(3, ...boardItems.map(item => Math.ceil((item.end ?? 0) / 2)));

  useEffect(() => {
    setDiffOpen(false);
    setLocationDetailKey(null);
    const button = selectedGroup ? groupButtons.current.get(selectedGroup.signature) : undefined;
    button?.scrollIntoView({ behavior: "smooth", block: "nearest", inline: "nearest" });
  }, [selectedGroup?.signature]);

  const optionHasConflict = (option: CandidateOption) => {
    const meetings = option.meetings.filter(item => activeInWeek(item, selectedWeek) && !item.unknown);
    const otherPreviewMeetings = previewOptions.filter(preview => preview.key !== option.key).flatMap(preview => preview.meetings).filter(item => activeInWeek(item, selectedWeek));
    return meetings.some(meeting => [...visibleCurrent, ...otherPreviewMeetings].some(other => timeOverlaps(meeting, other)));
  };
  const filteredOptions = candidateOptions.filter(option => {
    const planningCategory = planningFilterFor(option);
    const searchable = `${option.name} ${option.courseCode} ${option.teacher} ${option.category} ${planningCategory ?? ""}`.toLowerCase();
    const hasConflict = optionHasConflict(option);
    return !courseAlreadyCompleted(option, completedCourses) && (requirementFilter === "全部" || planningCategory === requirementFilter) && (!query || searchable.includes(query.toLowerCase())) &&
      (conflictFilter === "全部" || (conflictFilter === "有冲突" ? hasConflict : !hasConflict));
  }).sort((a, b) => Number(optionHasConflict(a)) - Number(optionHasConflict(b)) || (planningFilterFor(a) ?? a.category).localeCompare(planningFilterFor(b) ?? b.category, "zh-CN") || b.credits - a.credits || a.name.localeCompare(b.name, "zh-CN"));
  const pages = Math.max(1, Math.ceil(filteredOptions.length / pageSize));
  const safePage = Math.min(page, pages);
  const pagedOptions = filteredOptions.slice((safePage - 1) * pageSize, safePage * pageSize);

  useEffect(() => setPage(1), [selectedWeek, query, requirementFilter, conflictFilter]);

  const currentConflictKeys = new Set(visibleCurrent.filter((item, index) => visibleCurrent.some((other, otherIndex) => index !== otherIndex && timeOverlaps(item, other)) || visiblePreviews.some(other => timeOverlaps(item, other))).map(item => item.key));
  const previewConflictKeys = new Set(visiblePreviews.filter((item, index) => visibleCurrent.some(other => timeOverlaps(item, other)) || visiblePreviews.some((other, otherIndex) => index !== otherIndex && timeOverlaps(item, other))).map(item => item.key));
  const transitions = selectedGroup ? transitionsIntoGroup(scheduleItems, selectedGroup).filter(diff => diffCount(diff) > 0) : [];
  const aggregateDiff = transitions.reduce((aggregate, diff) => ({
    added: aggregate.added + diff.added.length,
    ended: aggregate.ended + diff.ended.length,
    adjusted: aggregate.adjusted + diff.timeAdjusted.length + diff.teacherAdjusted.length,
  }), { added: 0, ended: 0, adjusted: 0 });

  const saveCalibration = () => {
    const next: WeekCalibration = { term, anchorDate: formatLocalDate(new Date()), week: calibrationWeek };
    window.localStorage.setItem(`workbench-current-week:${term}`, JSON.stringify(next));
    setCalibrationRevision(value => value + 1);
    setSelectedWeek(calibrationWeek);
    setCalibrationOpen(false);
  };

  const togglePreview = (key: string) => {
    setPreviewKeys(keys => {
      if (keys.includes(key)) return keys.filter(value => value !== key);
      if (keys.length >= maximumPreviews) return keys;
      return [...keys, key];
    });
  };

  const renderBlock = (item: ScheduleItem) => {
    const isConflict = item.source === "current" ? currentConflictKeys.has(item.key) : previewConflictKeys.has(item.key);
    const related = boardItems.filter(other => other.day === item.day && timeOverlaps(item, other)).sort((a, b) => a.key.localeCompare(b.key));
    const lane = related.findIndex(other => other.key === item.key);
    const laneCount = Math.max(1, related.length);
    const startSession = Math.ceil(item.start! / 2);
    const endSession = Math.ceil(item.end! / 2);
    const locationGroups = selectedGroup ? locationsByWeek(scheduleItems, selectedGroup, item) : [];
    const locationChanged = locationGroups.length > 1;
    const previewIndex = previewKeys.indexOf(item.optionKey);
    const style = {
      gridColumn: String(item.day! + 1),
      gridRow: `${startSession + 1} / ${endSession + 2}`,
      left: `calc(${lane * 100 / laneCount}% + 5px)`,
      width: `calc(${100 / laneCount}% - 10px)`,
    };
    return <button type="button" className={`course-block ${item.source} preview-${previewIndex} ${isConflict ? "has-conflict" : ""}`} style={style} key={item.key} onClick={() => locationChanged && setLocationDetailKey(value => value === item.key ? null : item.key)} aria-expanded={locationChanged ? locationDetailKey === item.key : undefined}>
      <strong>{item.name}</strong>
      <span>{item.teacher || "教师未提供"}</span>
      <span>{item.location || "地点待定"}</span>
      <small>第{item.start}–{item.end}节</small>
      {locationChanged && <em>地点有变化</em>}
    </button>;
  };

  const detailedItem = boardItems.find(item => item.key === locationDetailKey);
  const detailedLocations = detailedItem && selectedGroup ? locationsByWeek(scheduleItems, selectedGroup, detailedItem) : [];

  return <section className="schedule-panel">
    <div className="schedule-title">
      <div><h2>七日课表</h2><p>{selectedGroup?.label ?? `第${selectedWeek}周`} · 当前按第 {selectedWeek} 周显示课程与地点。</p></div>
      <div className="schedule-legend"><span className="legend current" />当前课程 <span className="legend candidate" />预览课程 <span className="legend conflict" />时间冲突</div>
    </div>

    <section className="progress-band" aria-label="学分进度">
      <div className="progress-band-heading"><strong>学分进度</strong><small>与本次选课查询类别同步</small></div>
      {syncedProgress.length ? <div className="progress-band-inner">{syncedProgress.map(item => {
        const filter = progressFilterByKey[item.key];
        const confirmedCredits = item.completed_credits ?? 0;
        const requiredCredits = item.required_credits ?? 0;
        const projectedCredits = filter ? projectedCourseCredits(previewOptions, completedCourses, filter) : 0;
        const expectedCredits = confirmedCredits + projectedCredits;
        const remainingCredits = Math.max(0, requiredCredits - expectedCredits);
        return <article className="progress-card" key={item.key}>
          <header><strong>{item.label}</strong><span>{formatCredits(confirmedCredits)}/{formatCredits(requiredCredits)} 学分</span></header>
          {requiredCredits > 0 && <progress max={requiredCredits} value={Math.min(expectedCredits, requiredCredits)} />}
          <p>{remainingCredits > 0 ? `还差 ${formatCredits(remainingCredits)} 学分` : "总学分已满足"}{projectedCredits > 0 && ` · 预览后 ${formatCredits(expectedCredits)}`}{item.key === "cultural_quality" && ` · 四史${fourHistoriesComplete ? "已完成" : "待完成"} · D 类 2 学分待核对`}</p>
          <p className="progress-courses">{item.courses.length ? item.courses.map(course => <span key={`${course.code}-${course.name}`}>{course.name} · {formatCredits(course.credits)}</span>) : <span>暂无已确认课程</span>}</p>
        </article>;
      })}</div> : <div className="progress-band-inner"><article className="progress-card unavailable"><strong>尚未同步已修课程</strong><span>当前只能计算本次预览，不能判断真实缺口。</span></article></div>}
      {graduationProgress?.status === "incomplete" && <small>已修课程数据不完整，当前缺口只能作为下限参考。</small>}
    </section>

    <nav className="week-selector" aria-label="选择相同课表周次">
      <div className="week-selector-label"><strong>相同课表周次</strong><button className="text-button" onClick={() => setCalibrationOpen(value => !value)}>{currentWeek ? `本周：第${currentWeek}周` : "校准本周"}</button></div>
      <div className="week-groups">{weekGroups.map(group => {
        const active = group.signature === selectedGroup?.signature;
        const isCurrent = currentWeek !== null && group.weeks.includes(currentWeek);
        return <button key={`${group.signature}-${group.weeks[0]}`} ref={element => { if (element) groupButtons.current.set(group.signature, element); }} className={`week-group ${active ? "active" : "secondary"} ${group.empty ? "empty-week" : ""}`} onClick={() => setSelectedWeek(isCurrent ? currentWeek! : group.weeks[0])}><span>{group.label}</span>{isCurrent && <small>本周</small>}</button>;
      })}</div>
    </nav>

    {calibrationOpen && <section className="week-calibration" aria-label="校准当前教学周"><div><strong>{calibration && currentWeek === null ? "本周校准已超出当前学期范围" : "今天是第几教学周？"}</strong><p>按学期保存在本机，以后每周一自动递增。</p></div><select value={calibrationWeek} onChange={event => setCalibrationWeek(Number(event.target.value))} aria-label="当前教学周">{Array.from({ length: maxWeek }, (_, index) => index + 1).map(week => <option key={week} value={week}>第 {week} 周</option>)}</select><button onClick={saveCalibration}>保存本周</button></section>}

    <section className="schedule-diff">
      <button className="diff-summary" onClick={() => setDiffOpen(value => !value)} aria-expanded={diffOpen} disabled={!transitions.length}>
        <span>{transitions.length ? `${transitions.length} 次阶段变化` : selectedGroup?.weeks[0] === 1 ? "学期开始阶段" : "与上一阶段相同"}</span>
        {transitions.length > 0 && <strong>新增 {aggregateDiff.added} · 结束 {aggregateDiff.ended} · 调整 {aggregateDiff.adjusted}</strong>}
      </button>
      {diffOpen && <div className="diff-details">{transitions.map(diff => <article key={diff.week}><strong>第 {diff.week} 周起</strong>{diff.added.length > 0 && <p>新增：{diff.added.join("、")}</p>}{diff.ended.length > 0 && <p>结束：{diff.ended.join("、")}</p>}{diff.timeAdjusted.length > 0 && <p>时间调整：{diff.timeAdjusted.join("、")}</p>}{diff.teacherAdjusted.length > 0 && <p>老师调整：{diff.teacherAdjusted.join("、")}</p>}</article>)}</div>}
    </section>


    <div className="schedule-workspace">
      <div>
        <div className="timetable-scroll"><div className="timetable" style={{ "--periods": sessionCount } as React.CSSProperties}>
          <div className="time-head">时间 / 节次</div>{weekdays.map((day, index) => <div className="day-head" style={{ gridColumn: index + 2 }} key={day}>周{day}<small>星期{day}</small></div>)}
          {Array.from({ length: sessionCount }, (_, index) => <React.Fragment key={index}><div className="period-label" style={{ gridRow: index + 2 }}><strong>第{index * 2 + 1}–{index * 2 + 2}节</strong><small>{sessionTimes[index] ?? "时间待定"}</small></div>{weekdays.map((_, day) => <div className="time-cell" style={{ gridColumn: day + 2, gridRow: index + 2 }} key={day} />)}</React.Fragment>)}
          {boardItems.map(renderBlock)}
        </div></div>
        {detailedItem && <section className="location-detail"><div><strong>{detailedItem.name} · 地点安排</strong><button className="text-button" onClick={() => setLocationDetailKey(null)}>关闭</button></div>{detailedLocations.map(group => <p key={group.weeks.join("-")}><span>{formatWeekGroup(group.weeks)}</span>{group.locations.join("、")}</p>)}</section>}
      </div>

      <aside className="course-browser">
        <div className="browser-heading"><div><h3>待选课程</h3><small>按第 {selectedWeek} 周检测冲突 · 最多预览 {maximumPreviews} 门</small></div><span>{filteredOptions.length} 门 · {safePage}/{pages} 页</span></div>
        <input aria-label="搜索待选课程" value={query} onChange={event => setQuery(event.target.value)} placeholder="搜索课程名、代码或教师" />
        <div className="course-filters"><select value={requirementFilter} onChange={event => setRequirementFilter(event.target.value)} aria-label="培养要求">{requirementFilterOptions.map(value => <option key={value}>{value === "全部" ? "全部培养要求" : value}</option>)}</select><select value={conflictFilter} onChange={event => setConflictFilter(event.target.value)} aria-label="冲突状态"><option>全部</option><option>无冲突</option><option>有冲突</option></select></div>
        <div className="requirement-hints" aria-label="培养要求提醒"><span><b>体育</b> 每学年选 1 门</span><span><b>创新创业</b> 社会实践另计</span><span><b>英语</b> 大一优先</span></div>
        <div className="course-list">{pagedOptions.length ? pagedOptions.map(option => {
          const hasConflict = optionHasConflict(option);
          const previewIndex = previewKeys.indexOf(option.key);
          const previewed = previewIndex >= 0;
          const previewLimitReached = !previewed && previewKeys.length >= maximumPreviews;
          const planningCategory = planningFilterFor(option);
          const knownConflict = option.meetings.some(meeting => !meeting.unknown && current.some(other => !other.unknown && executionTimeOverlaps(meeting, other)));
          const executionBlocked = executionPending || option.unknown || knownConflict || !option.executionReady;
          const executionTitle = !option.executionReady ? "教学班缺少页面提供的执行身份" : option.unknown ? "上课时间未知，禁止执行" : knownConflict ? "与当前课表冲突" : executionPending ? "已有执行任务正在进行" : "确认并提交一次选课请求";
          return <article className={`candidate-row ${hasConflict ? "has-conflict" : ""}`} key={option.key}><div className="candidate-copy"><strong>{option.name}</strong><span>{planningCategory ?? option.category} · {option.credits ? `${formatCredits(option.credits)} 学分` : "学分待核对"} · {option.teacher || "教师未提供"}</span><small>{describeOptionMeetings(option.meetings)}</small></div><div className="candidate-actions">{previewed && <i className={`preview-dot preview-${previewIndex}`} aria-label={`预览颜色 ${previewIndex + 1}`} />}<b>{knownConflict ? "冲突" : option.unknown ? "待定" : "可排"}</b>{!option.unknown && <button className={previewed ? "secondary" : ""} disabled={previewLimitReached} title={previewLimitReached ? "最多同时预览3门课程" : undefined} onClick={() => togglePreview(option.key)}>{previewed ? "取消" : "预览"}</button>}<button className="execute-selection" disabled={executionBlocked} title={executionTitle} onClick={() => onExecuteSection?.(String(option.identity ?? option.key), option.name)}>选课</button></div></article>;
        }) : <p className="empty">没有符合条件的待选课程。</p>}</div>
        <nav className="pagination" aria-label="待选课程分页"><button className="secondary" disabled={safePage <= 1} onClick={() => setPage(safePage - 1)}>上一页</button><span>{safePage}</span><button className="secondary" disabled={safePage >= pages} onClick={() => setPage(safePage + 1)}>下一页</button></nav>
      </aside>
    </div>
    {[...current, ...candidateOptions.flatMap(option => option.meetings)].filter(item => item.unknown).length > 0 && <p className="unknown-hint">有课程的时间尚未能可靠解析；它们不会被显示为空闲时间，请在待选课程中核对原始上课信息。</p>}
  </section>;
}
