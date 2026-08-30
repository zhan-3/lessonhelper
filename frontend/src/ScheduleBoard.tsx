import React, { useEffect, useLayoutEffect, useMemo, useRef, useState } from "react";
import type { WorkbenchState } from "./api";
import {
  candidateExecutionStatus,
  candidateOption,
  compressWeeks,
  courseAlreadyCompleted,
  courseColor,
  deriveCurrentWeek,
  describeOptionMeetings,
  expandScheduleItems,
  formatWeekGroup,
  locationsByWeek,
  planningFilterFor,
  progressFilterByKey,
  progressKeysByQueryCode,
  projectedCourseCredits,
  selectedCandidateOptions,
  weekItems,
  type CandidateOption,
  type RequirementFilter,
  type ScheduleItem,
  type SelectionWindow,
  type WeekCalibration,
} from "./schedule";

const weekdays = ["一", "二", "三", "四", "五", "六", "日"];
const sessionTimes = ["08:00–09:45", "10:05–11:50", "14:00–15:45", "16:05–17:50", "18:40–20:25", "20:45–22:30"];
const pageSize = 12;
const maximumPreviews = 3;
const formatCredits = (credits: number) => Number.isInteger(credits) ? String(credits) : String(credits).replace(/0+$/, "").replace(/\.$/, "");
const courseColorStyle = (courseName: string) => ({ "--course-color": courseColor(courseName) }) as React.CSSProperties;

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

export function ScheduleBoard({ timetable, selection, graduationProgress, selectionWindows, studentGrade, executionHistory, previewKeys, onTogglePreview, onExecuteSection, executionPending }: {
  timetable: WorkbenchState["snapshots"]["timetable"];
  selection: WorkbenchState["snapshots"]["selection"];
  graduationProgress?: WorkbenchState["graduation_progress"];
  selectionWindows: SelectionWindow[];
  studentGrade: string;
  executionHistory: Array<Record<string, unknown>>;
  previewKeys: string[];
  onTogglePreview: (key: string) => void;
  onExecuteSection?: (sectionId: string, courseName: string) => void;
  executionPending?: boolean;
}) {
  const term = timetable?.term || selection?.term || "当前学期";
  const [selectedWeek, setSelectedWeek] = useState(1);
  const [query, setQuery] = useState("");
  const [requirementFilter, setRequirementFilter] = useState("全部");
  const [conflictFilter, setConflictFilter] = useState("全部");
  const [page, setPage] = useState(1);
  const [calibrationRevision, setCalibrationRevision] = useState(0);
  const [calibrationOpen, setCalibrationOpen] = useState(false);
  const [calibrationWeek, setCalibrationWeek] = useState(1);
  const [locationDetailKey, setLocationDetailKey] = useState<string | null>(null);
  const workspaceRef = useRef<HTMLDivElement>(null);
  const [browserHeight, setBrowserHeight] = useState<number | null>(null);
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
    const nextCalibration = readCalibration(term);
    const nextCurrentWeek = deriveCurrentWeek(nextCalibration, maxWeek);
    setSelectedWeek(nextCurrentWeek ?? 1);
    setCalibrationWeek(nextCurrentWeek ?? 1);
    setCalibrationOpen(!nextCalibration || nextCurrentWeek === null);
  }, [term, maxWeek]);

  const previewOptions = candidateOptions.filter(option => previewKeys.includes(option.key));
  const selectedOptions = selectedCandidateOptions(
    candidateOptions,
    (timetable?.payload.entries as Record<string, unknown>[] | undefined) ?? [],
    executionHistory,
  );
  const selectedIdentities = new Set(selectedOptions.map(option => `${option.courseCode || option.name}`.trim().toLowerCase()));
  const queuedOptions = previewOptions.filter(option => !selectedIdentities.has(`${option.courseCode || option.name}`.trim().toLowerCase()));
  const unknownCount = [...current, ...candidateOptions.flatMap(option => option.meetings)].filter(item => item.unknown).length;
  // 进度面板与待选课程同步：只展示本次选课查询覆盖的培养要求。
  const syncedProgress = (() => {
    const queries = (selection?.payload.queries as Array<{ category?: unknown }> | undefined) ?? [];
    const keys = new Set(queries.flatMap(item => progressKeysByQueryCode[String(item.category ?? "")] ?? []));
    return progressItems.filter(item => keys.has(item.key));
  })();
  const scheduleItems = [...current, ...previewOptions.flatMap(option => option.meetings)];
  // 预览只叠加课程，不重写用户已经建立的学期周次导航。
  const weekGroups = compressWeeks(current.length ? current : scheduleItems, maxWeek);
  const selectedGroup = weekGroups.find(group => group.weeks.includes(selectedWeek)) ?? weekGroups[0];
  const visibleCurrent = weekItems(current, selectedWeek);
  const visiblePreviews = weekItems(previewOptions.flatMap(option => option.meetings), selectedWeek);
  const boardItems = [...visibleCurrent, ...visiblePreviews];
  const sessionCount = Math.max(3, ...boardItems.map(item => Math.ceil((item.end ?? 0) / 2)));

  useLayoutEffect(() => {
    const compute = () => {
      if (window.innerWidth < 1121) { setBrowserHeight(null); return; }
      const title = workspaceRef.current?.querySelector(".schedule-title");
      const scroll = workspaceRef.current?.querySelector(".timetable-scroll");
      if (!title || !scroll) return;
      setBrowserHeight(Math.floor(title.getBoundingClientRect().height + scroll.getBoundingClientRect().height));
    };
    compute();
    window.addEventListener("resize", compute);
    return () => window.removeEventListener("resize", compute);
  }, [sessionCount]);

  useEffect(() => {
    setLocationDetailKey(null);
    const button = selectedGroup ? groupButtons.current.get(selectedGroup.signature) : undefined;
    button?.scrollIntoView({ behavior: "smooth", block: "nearest", inline: "nearest" });
  }, [selectedGroup?.signature]);

  const semesterConflict = (option: CandidateOption) => option.meetings.some(meeting => !meeting.unknown && current.some(other => !other.unknown && executionTimeOverlaps(meeting, other)));
  const filteredOptions = candidateOptions.filter(option => {
    const planningCategory = planningFilterFor(option);
    const searchable = `${option.name} ${option.courseCode} ${option.teacher} ${option.category} ${planningCategory ?? ""}`.toLowerCase();
    const hasConflict = semesterConflict(option);
    return !courseAlreadyCompleted(option, completedCourses) && (requirementFilter === "全部" || planningCategory === requirementFilter) && (!query || searchable.includes(query.toLowerCase())) &&
      (conflictFilter === "全部" || (conflictFilter === "有冲突" ? hasConflict : !hasConflict));
  }).sort((a, b) => Number(semesterConflict(a)) - Number(semesterConflict(b)) || (planningFilterFor(a) ?? a.category).localeCompare(planningFilterFor(b) ?? b.category, "zh-CN") || b.credits - a.credits || a.name.localeCompare(b.name, "zh-CN"));
  const pages = Math.max(1, Math.ceil(filteredOptions.length / pageSize));
  const safePage = Math.min(page, pages);
  const pagedOptions = filteredOptions.slice((safePage - 1) * pageSize, safePage * pageSize);

  useEffect(() => setPage(1), [selectedWeek, query, requirementFilter, conflictFilter]);

  const currentConflictKeys = new Set(visibleCurrent.filter((item, index) => visibleCurrent.some((other, otherIndex) => index !== otherIndex && timeOverlaps(item, other)) || visiblePreviews.some(other => timeOverlaps(item, other))).map(item => item.key));
  const previewConflictKeys = new Set(visiblePreviews.filter((item, index) => visibleCurrent.some(other => timeOverlaps(item, other)) || visiblePreviews.some((other, otherIndex) => index !== otherIndex && timeOverlaps(item, other))).map(item => item.key));
  const saveCalibration = () => {
    const next: WeekCalibration = { term, anchorDate: formatLocalDate(new Date()), week: calibrationWeek };
    window.localStorage.setItem(`workbench-current-week:${term}`, JSON.stringify(next));
    setCalibrationRevision(value => value + 1);
    setSelectedWeek(calibrationWeek);
    setCalibrationOpen(false);
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
    const style = {
      ...courseColorStyle(item.name),
      gridColumn: String(item.day! + 1),
      gridRow: `${startSession + 1} / ${endSession + 2}`,
      left: `calc(${lane * 100 / laneCount}% + 5px)`,
      width: `calc(${100 / laneCount}% - 10px)`,
    };
    const content = <><strong>{item.name}</strong><span>{item.teacher || "教师未提供"}</span><span>{item.location || "地点待定"}</span><small>第{item.start}–{item.end}节</small>{locationChanged && <em>查看分周地点</em>}</>;
    const className = `course-block ${item.source} ${isConflict ? "has-conflict" : ""}`;
    return locationChanged
      ? <button type="button" className={className} style={style} key={item.key} onClick={() => setLocationDetailKey(value => value === item.key ? null : item.key)} aria-expanded={locationDetailKey === item.key} aria-label={`${item.name}，查看分周地点`}>{content}</button>
      : <article className={className} style={style} key={item.key}>{content}</article>;
  };

  const detailedItem = boardItems.find(item => item.key === locationDetailKey);
  const detailedLocations = detailedItem && selectedGroup ? locationsByWeek(scheduleItems, selectedGroup, detailedItem) : [];

  return <section className="schedule-panel">
    <section className="progress-band" aria-label="毕业进度规划参考">
      <div className="progress-band-heading"><strong>毕业进度规划参考</strong><small>公开规则 + 本地快照 · 非学校正式毕业审核</small></div>
      {syncedProgress.length ? <div className="progress-band-inner">{syncedProgress.map(item => {
        const filter = progressFilterByKey[item.key];
        const confirmedCredits = item.completed_credits ?? 0;
        const requiredCredits = item.key === "cultural_quality" && studentGrade === "2021" ? 10 : (item.required_credits ?? 0);
        const selectedCredits = filter ? projectedCourseCredits(selectedOptions, completedCourses, filter) : 0;
        const selectedFacts = selectedOptions.map(option => ({ code: option.courseCode, name: option.name }));
        const queuedCredits = filter ? projectedCourseCredits(queuedOptions, [...completedCourses, ...selectedFacts], filter) : 0;
        const expectedCredits = confirmedCredits + selectedCredits + queuedCredits;
        const remainingCredits = Math.max(0, requiredCredits - expectedCredits);
        const width = (credits: number, before = 0) => requiredCredits > 0 ? `${Math.max(0, Math.min(credits, requiredCredits - before)) / requiredCredits * 100}%` : "0%";
        const culturalRule = studentGrade === "2021"
          ? "2021级：总计≥10学分、覆盖≥4模块；讲座最多计1学分"
          : ["2022", "2023", "2024"].includes(studentGrade)
            ? `总计≥8学分；D类≥2、四史≥1门（均计入总额，四史${fourHistoriesComplete ? "已完成" : "待核对"}）`
            : "D类、四史和讲座是总额内的子约束；当前年级请按个人培养方案核对";
        const rule = item.key === "cultural_quality" ? culturalRule
          : item.key === "innovation_and_practice" || item.key === "innovation" ? "创新创业+社会实践合计≥6，社会实践≥1；创新创业单项最低值待按个人方案核对"
            : item.key === "outside_major_elective" ? (studentGrade === "2025" ? "跨专业发展≥10；项目驱动超出10学分的部分可按规定转入创新实践" : ["2022", "2023", "2024"].includes(studentGrade) ? "跨专业发展≥10，通常须选定一个体系并在体系内修满" : "外专业/跨专业口径随年级变化，请按个人培养方案核对") : "";
        return <article className="progress-card" key={item.key}>
          <header><strong>{item.key === "outside_major_elective" ? "跨专业发展课程" : item.label}</strong><span>预计 {formatCredits(expectedCredits)}/{formatCredits(requiredCredits)} 学分</span></header>
          {requiredCredits > 0 && <div className="credit-meter" role="progressbar" aria-label={`${item.label}预计学分`} aria-valuemin={0} aria-valuemax={requiredCredits} aria-valuenow={Math.min(expectedCredits, requiredCredits)}><span className="confirmed" style={{ width: width(confirmedCredits) }} /><span className="selected" style={{ width: width(selectedCredits, confirmedCredits) }} /><span className="queued" style={{ width: width(queuedCredits, confirmedCredits + selectedCredits) }} /></div>}
          <p>{remainingCredits > 0 ? `还差 ${formatCredits(remainingCredits)} 学分` : "总学分已满足"} · <b className="credit-confirmed">已确认 {formatCredits(confirmedCredits)}</b>{selectedCredits > 0 && <> · <b className="credit-selected">本学期已选 +{formatCredits(selectedCredits)}</b></>}{queuedCredits > 0 && <> · <b className="credit-queued">队列 +{formatCredits(queuedCredits)}</b></>}</p>
          {rule && <p className="requirement-rule">{rule}</p>}
          <p className="progress-courses">{item.courses.map(course => <span className="confirmed-course" key={`${course.code}-${course.name}`}>{course.name} · {formatCredits(course.credits)}</span>)}{filter && selectedOptions.filter(option => planningFilterFor(option) === filter && !courseAlreadyCompleted(option, completedCourses)).map(option => <span className="selected-course" key={`selected-${option.key}`}>{option.name} · {formatCredits(option.credits)}</span>)}{filter && queuedOptions.filter(option => planningFilterFor(option) === filter && !courseAlreadyCompleted(option, [...completedCourses, ...selectedFacts])).map(option => <span className="queued-course" key={`queued-${option.key}`}>{option.name} · {formatCredits(option.credits)}</span>)}{!item.courses.length && !selectedCredits && !queuedCredits && <span>暂无已确认或预计课程</span>}</p>
        </article>;
      })}</div> : <div className="progress-band-inner"><article className="progress-card unavailable"><strong>尚未同步已修课程</strong><span>同步后可区分已确认、本学期已选与队列估值。</span></article></div>}
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

    <div className="schedule-workspace" ref={workspaceRef}>
      <div className="timetable-column">
        <div className="schedule-title">
          <div><h2>七日课表</h2><p>{selectedGroup?.label ?? `第${selectedWeek}周`} · 当前按第 {selectedWeek} 周显示课程与地点。</p></div>
          <div className="schedule-legend"><span className="legend current" />当前课程 <span className="legend candidate" />预览课程 <span className="legend conflict" />时间冲突</div>
        </div>
        <div className="timetable-scroll"><div className="timetable" style={{ "--periods": sessionCount } as React.CSSProperties}>
          <div className="time-head">时间 / 节次</div>{weekdays.map((day, index) => <div className="day-head" style={{ gridColumn: index + 2 }} key={day}>周{day}<small>星期{day}</small></div>)}
          {Array.from({ length: sessionCount }, (_, index) => <React.Fragment key={index}><div className="period-label" style={{ gridRow: index + 2 }}><strong>第{index * 2 + 1}–{index * 2 + 2}节</strong><small>{sessionTimes[index] ?? "时间待定"}</small></div>{weekdays.map((_, day) => <div className="time-cell" style={{ gridColumn: day + 2, gridRow: index + 2 }} key={day} />)}</React.Fragment>)}
          {boardItems.map(renderBlock)}
        </div></div>
        {detailedItem && <section className="location-detail"><div><strong>{detailedItem.name} · 地点安排</strong><button className="text-button" onClick={() => setLocationDetailKey(null)}>关闭</button></div>{detailedLocations.map(group => <p key={group.weeks.join("-")}><span>{formatWeekGroup(group.weeks)}</span>{group.locations.join("、")}</p>)}</section>}
      </div>

      <aside className="course-browser" style={browserHeight ? { height: browserHeight } : undefined}>
        <div className="browser-heading"><div><h3>查找待选课程</h3><small>先加入预览，确认整学期无冲突后再选课</small></div><span>{filteredOptions.length} 门 · {safePage}/{pages} 页</span></div>
        <input aria-label="搜索待选课程" value={query} onChange={event => setQuery(event.target.value)} placeholder="搜索课程名、代码或教师" />
        <div className="course-filters"><select value={requirementFilter} onChange={event => setRequirementFilter(event.target.value)} aria-label="培养要求">{requirementFilterOptions.map(value => <option key={value}>{value === "全部" ? "全部培养要求" : value}</option>)}</select><select value={conflictFilter} onChange={event => setConflictFilter(event.target.value)} aria-label="冲突状态"><option>全部</option><option>无冲突</option><option>有冲突</option></select></div>
        <div className="requirement-hints" aria-label="培养要求提醒"><span><b>体育</b> 每学年选 1 门</span><span><b>创新创业</b> 社会实践另计</span><span><b>英语</b> 大一优先</span></div>
        <div className="course-list" style={browserHeight ? { maxHeight: "none" } : undefined}>{pagedOptions.length ? pagedOptions.map(option => {
          const previewIndex = previewKeys.indexOf(option.key);
          const previewed = previewIndex >= 0;
          const planningCategory = planningFilterFor(option);
          const knownConflict = semesterConflict(option);
          const executionStatus = candidateExecutionStatus(option, selectionWindows, studentGrade);
          const executionBlocked = executionPending || option.unknown || knownConflict || !executionStatus.canExecute;
          const scheduleStatus = option.unknown ? "时间待确认" : knownConflict ? "学期内有冲突" : "学期内无冲突";
          const executionReason = option.unknown ? "暂不可选：上课时间待确认" : knownConflict ? "暂不可选：与当前课表冲突" : executionPending ? "暂不可选：正在执行其他任务" : executionStatus.canExecute ? executionStatus.reason : `暂不可选：${executionStatus.reason}`;
          return <article className={`candidate-row ${knownConflict ? "has-conflict" : ""} ${previewed ? "is-previewed" : ""}`} key={option.key}>
            <div className="candidate-copy"><div className="candidate-title">{previewed && <i className="preview-dot" style={courseColorStyle(option.name)} aria-hidden="true" />}<strong>{option.name}</strong></div><span>{planningCategory ?? option.category} · {option.credits ? `${formatCredits(option.credits)} 学分` : "学分待核对"} · {option.teacher || "教师未提供"}</span><small>{describeOptionMeetings(option.meetings)}</small><div className="candidate-status" id={`status-${option.key}`}><b className={knownConflict || option.unknown ? "blocked" : "ready"}>{scheduleStatus}</b><span>{executionReason}</span></div></div>
            <div className="candidate-actions">{!option.unknown && <button className={previewed ? "secondary" : ""} aria-describedby={`status-${option.key}`} onClick={() => onTogglePreview(option.key)}>{previewed ? "移出队列" : "加入队列"}</button>}<button className="execute-selection" disabled={executionBlocked} aria-describedby={`status-${option.key}`} onClick={() => onExecuteSection?.(String(option.identity ?? option.key), option.name)}>{executionPending ? "处理中" : "选课"}</button></div>
          </article>;
        }) : <p className="empty">没有符合条件的待选课程。</p>}</div>
        {unknownCount > 0 && <p className="unknown-note">有课程时间未能可靠解析，不会显示为空闲时间；请在列表中核对原始上课信息。</p>}
        <nav className="pagination" aria-label="待选课程分页"><button className="page-arrow" aria-label="上一页" disabled={safePage <= 1} onClick={() => setPage(safePage - 1)}>‹</button><button className="page-arrow" aria-label="下一页" disabled={safePage >= pages} onClick={() => setPage(safePage + 1)}>›</button></nav>
      </aside>
    </div>
  </section>;
}
