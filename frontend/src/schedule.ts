export type ScheduleItem = Record<string, unknown> & {
  source: "current" | "candidate";
  optionKey: string;
  key: string;
  day: number | null;
  start: number | null;
  end: number | null;
  weeks: number[];
  unknown: boolean;
  name: string;
  category: string;
  teacher: string;
  location: string;
};

export type CandidateOption = Record<string, unknown> & {
  key: string;
  name: string;
  category: string;
  teacher: string;
  courseCode: string;
  queryCode: string;
  credits: number;
  meetings: ScheduleItem[];
  unknown: boolean;
  executionReady: boolean;
};

export type WeekGroup = {
  signature: string;
  weeks: number[];
  label: string;
  empty: boolean;
};

export type ScheduleDiff = {
  week: number;
  added: string[];
  ended: string[];
  timeAdjusted: string[];
  teacherAdjusted: string[];
};

export type WeekCalibration = {
  term: string;
  anchorDate: string;
  week: number;
};

export type SelectionWindow = {
  action?: string;
  grades?: string[];
  category_codes?: string[];
  opens_at?: string;
  closes_at?: string;
  category_text?: string;
  method?: string;
};

/** Only actionable selection windows for the current student's grade. */
export function selectionWindowsForGrade(windows: SelectionWindow[], grade: string): SelectionWindow[] {
  if (!grade) return [];
  return windows
    .filter(window => window.action === "selection" && (window.grades ?? []).includes(grade))
    .sort((left, right) => String(left.opens_at ?? "").localeCompare(String(right.opens_at ?? "")));
}

type ParsedNoticeDateTime = {
  year: string;
  month: string;
  day: string;
  time: string;
};

export type SelectionWindowDisplay = {
  date: string;
  time: string;
  fullRange: string;
};

const parseNoticeDateTime = (value?: string): ParsedNoticeDateTime | null => {
  const match = String(value ?? "").match(/^(\d{4})-(\d{2})-(\d{2})[ T](\d{2}:\d{2})/);
  if (!match) return null;
  return { year: match[1], month: match[2], day: match[3], time: match[4] };
};

const noticeDate = (value: ParsedNoticeDateTime) =>
  new Date(Number(value.year), Number(value.month) - 1, Number(value.day), ...value.time.split(":").map(Number) as [number, number]);

const weekdayLabel = (value: ParsedNoticeDateTime) => `周${"日一二三四五六"[noticeDate(value).getDay()]}`;

/** Compact, timezone-stable labels for the confirmed-notice board. */
export function selectionWindowDisplay(window: SelectionWindow): SelectionWindowDisplay {
  const opens = parseNoticeDateTime(window.opens_at);
  const closes = parseNoticeDateTime(window.closes_at);
  if (!opens) return { date: "日期待定", time: "时间待定", fullRange: "开放时间待定" };

  const sameDay = closes && opens.year === closes.year && opens.month === closes.month && opens.day === closes.day;
  const date = !closes || sameDay
    ? `${opens.month}.${opens.day} ${weekdayLabel(opens)}`
    : `${opens.month}.${opens.day}–${closes.month}.${closes.day}`;
  const time = closes ? `${opens.time}–${closes.time}` : `${opens.time} 起`;
  const fullRange = closes
    ? `${opens.year}-${opens.month}-${opens.day} ${opens.time} 至 ${closes.year}-${closes.month}-${closes.day} ${closes.time}`
    : `${opens.year}-${opens.month}-${opens.day} ${opens.time} 起`;
  return { date, time, fullRange };
}

/** The board already implies selection, so remove a duplicated source suffix. */
export const selectionCategoryLabel = (window: SelectionWindow) =>
  String(window.category_text || "课程类别待确认").replace(/\s*——\s*选课\s*$/, "");

export type CandidateExecutionStatus = {
  canExecute: boolean;
  reason: string;
};

/** Explain execution availability from the confirmed category window first. */
export function candidateExecutionStatus(
  option: Pick<CandidateOption, "queryCode" | "executionReady">,
  windows: SelectionWindow[],
  grade: string,
  now: Date = new Date(),
): CandidateExecutionStatus {
  const matching = windows
    .filter(window =>
      window.action === "selection" &&
      window.method === "academic_system" &&
      (window.grades ?? []).includes(grade) &&
      (window.category_codes ?? []).includes(option.queryCode)
    )
    .map(window => ({
      window,
      opens: parseNoticeDateTime(window.opens_at),
      closes: parseNoticeDateTime(window.closes_at),
    }))
    .filter(item => item.opens && item.closes)
    .sort((left, right) => noticeDate(left.opens!).getTime() - noticeDate(right.opens!).getTime());

  if (!matching.length) {
    return { canExecute: false, reason: "没有匹配当前年级和课程类别的选课窗口" };
  }

  const nowTime = now.getTime();
  const active = matching.find(item =>
    noticeDate(item.opens!).getTime() <= nowTime && nowTime <= noticeDate(item.closes!).getTime()
  );
  if (active) {
    return option.executionReady
      ? { canExecute: true, reason: "可以提交选课" }
      : { canExecute: false, reason: "开放时段内未取得教学班标识，请刷新待选课程" };
  }

  const next = matching.find(item => noticeDate(item.opens!).getTime() > nowTime);
  if (next) {
    const display = selectionWindowDisplay(next.window);
    return { canExecute: false, reason: `当前不在选课时间；下次开放 ${display.date.split(" ")[0]} ${display.time}` };
  }
  return { canExecute: false, reason: "当前不在选课时间；本轮选课已结束" };
}

/** Stable dark-mode color for one course across weeks, sections, and previews. */
export function courseColor(courseName: string): string {
  const normalized = courseName.trim().toLocaleLowerCase("zh-CN").replace(/\s+/g, "");
  let hash = 2166136261;
  for (const character of normalized || "未命名课程") {
    hash = Math.imul(hash ^ character.codePointAt(0)!, 16777619);
  }
  const unsigned = hash >>> 0;
  const hue = unsigned % 360;
  const saturation = 40 + ((unsigned >>> 9) % 9);
  const lightness = 28 + ((unsigned >>> 18) % 4);
  return `hsl(${hue} ${saturation}% ${lightness}%)`;
}

export const requirementFilters = ["全部", "创新创业", "文化素质", "跨专业", "体育", "英语"] as const;
export type RequirementFilter = typeof requirementFilters[number];

/** Map the backend query category code (payload sections' query_code) to the decision-level filter. */
const queryCodeFilters: Record<string, Exclude<RequirementFilter, "全部">> = {
  ty: "体育",
  cxyx: "创新创业",
  cxsy: "创新创业",
  cxcy: "创新创业",
  szhx: "文化素质",
  xsxk: "跨专业",
};

export const queryFilterFor = (code: string): Exclude<RequirementFilter, "全部"> | null =>
  queryCodeFilters[code] ?? null;

/** Graduation-progress requirement keys covered by each selection query code. */
export const progressKeysByQueryCode: Record<string, string[]> = {
  szhx: ["cultural_quality"],
  // Official public rules verify a combined innovation/entrepreneurship +
  // social-practice requirement, not a universal innovation-only minimum.
  cxyx: ["innovation_and_practice"],
  cxsy: ["innovation_and_practice"],
  cxcy: ["innovation_and_practice"],
  xsxk: ["outside_major_elective"],
};

/** Which candidate filter each graduation-progress requirement projects onto. */
export const progressFilterByKey: Record<string, Exclude<RequirementFilter, "全部">> = {
  cultural_quality: "文化素质",
  innovation: "创新创业",
  innovation_and_practice: "创新创业",
  outside_major_elective: "跨专业",
};

/** Map volatile source labels to the small set of decisions students make. */
export function requirementFilterFor(category: string): Exclude<RequirementFilter, "全部"> | null {
  const value = category.replace(/\s/g, "");
  if (/创新创业|创新实验|创新研修|创业课程/.test(value)) return "创新创业";
  if (/文化素质|素质教育|文理通识/.test(value)) return "文化素质";
  if (/外专业|跨专业/.test(value)) return "跨专业";
  // 选课白名单里没有本专业查询，"专业基础/核心/选修课"标签只可能来自外专业查询。
  // 这是旧快照（缺 query_code）的兜底；新快照里 query_code 优先，不受此影响。
  if (/专业(基础|核心|选修|实践|大类)/.test(value)) return "跨专业";
  if (/体育/.test(value)) return "体育";
  if (/英语|外语/.test(value)) return "英语";
  return null;
}

/** Prefer the query source code (authoritative), then fall back to the raw label. */
export function planningFilterFor(option: { queryCode?: string; category: string }): Exclude<RequirementFilter, "全部"> | null {
  return queryFilterFor(option.queryCode ?? "") ?? requirementFilterFor(option.category);
}

/** 与所选周无关的完整时间+地点描述：同名同师、不同时段的教学班靠它区分。 */
export function describeOptionMeetings(meetings: ScheduleItem[]): string {
  const known = meetings.filter(item => !item.unknown);
  if (!known.length) return "上课时间待定";
  const unique = new Map<string, string>();
  for (const item of known) {
    const day = item.day !== null && chineseWeekdays[item.day - 1] ? chineseWeekdays[item.day - 1] : "?";
    const slot = `周${day}第${item.start ?? "?"}–${item.end ?? "?"}节`;
    const place = item.location ? ` ${item.location}` : "";
    const weeks = item.weeks.length ? `（${formatWeekGroup(item.weeks)}）` : "";
    const text = `${slot}${place}${weeks}`;
    if (!unique.has(text)) unique.set(text, text);
  }
  return [...unique.values()].join("；");
}

const chineseWeekdays = "一二三四五六日";

const firstNumber = (value: unknown): number | null => {
  const match = String(value ?? "").match(/\d+/);
  return match ? Number(match[0]) : null;
};

const weekdayFrom = (value: unknown): number | null => {
  const text = String(value ?? "");
  const numeric = text.match(/(?:星期|周)\s*([1-7])/);
  if (numeric) return Number(numeric[1]);
  const chinese = text.match(/(?:星期|周)\s*([一二三四五六日天])/);
  if (!chinese) return null;
  return chinese[1] === "天" ? 7 : chineseWeekdays.indexOf(chinese[1]) + 1;
};

export function expandWeekSpec(spec: string): number[] {
  const parity = spec.includes("单") ? "odd" : spec.includes("双") ? "even" : "all";
  const clean = spec.replace(/[单双周\[\]]/g, "");
  const weeks = clean.split(/[,，]/).flatMap(token => {
    const range = token.trim().match(/^(\d+)\s*[-~至]\s*(\d+)$/);
    if (range) {
      const start = Number(range[1]);
      const end = Number(range[2]);
      return Array.from({ length: Math.max(0, end - start + 1) }, (_, index) => start + index);
    }
    const value = Number(token.trim());
    return Number.isFinite(value) ? [value] : [];
  });
  return Array.from(new Set(weeks.filter(week => parity === "all" || (parity === "odd" ? week % 2 === 1 : week % 2 === 0)))).sort((a, b) => a - b);
}

const weeksFromRaw = (raw: Record<string, unknown>, text: string): number[] => {
  const supplied = raw.week_numbers ?? raw.weeks;
  if (Array.isArray(supplied)) return supplied.map(Number).filter(Number.isFinite).sort((a, b) => a - b);
  const matches = Array.from(text.matchAll(/\[([^\]]+?)周\]/g));
  if (matches.length) return Array.from(new Set(matches.flatMap(match => expandWeekSpec(match[1])))).sort((a, b) => a - b);
  const start = firstNumber(raw.week_start);
  const end = firstNumber(raw.week_end);
  if (start && end) return Array.from({ length: end - start + 1 }, (_, index) => start + index);
  return String(supplied ?? "").split(/[,，\s]+/).map(Number).filter(Number.isFinite).sort((a, b) => a - b);
};

const teacherAndLocation = (raw: Record<string, unknown>, suppliedLocation: string) => {
  const explicitTeacher = String(raw.teacher ?? "").trim();
  const location = suppliedLocation.trim();
  if (explicitTeacher || !location) return { teacher: explicitTeacher, location };
  const locationMatch = location.match(/([A-Za-z0-9一-龥]+(?:楼|公寓|体育场|馆|中心|实验室)[-－]?[A-Za-z0-9一-龥]*)$/);
  if (!locationMatch || locationMatch.index === 0) return { teacher: "", location };
  const teacher = location.slice(0, locationMatch.index).replace(/\[[^\]]+\]/g, "").replace(/^[,，、\s]+|[,，、\s]+$/g, "").trim();
  return { teacher, location: locationMatch[1] };
};

const makeMeeting = (
  raw: Record<string, unknown>,
  source: ScheduleItem["source"],
  optionKey: string,
  key: string,
  day: number | null,
  start: number | null,
  end: number | null,
  weeks: number[],
  location: string,
): ScheduleItem => {
  const details = teacherAndLocation(raw, location);
  return {
    ...raw,
    source,
    optionKey,
    key,
    day: day && day >= 1 && day <= 7 ? day : null,
    start,
    end,
    weeks,
    unknown: raw.conflict_status === "unknown" || !day || !start || !end,
    name: String(raw.course_name ?? raw.name ?? raw.title ?? "未命名课程"),
    category: String(raw.category ?? "未分类"),
    teacher: details.teacher,
    location: details.location,
  };
};

export function expandScheduleItems(raw: Record<string, unknown>, source: ScheduleItem["source"], index: number): ScheduleItem[] {
  const text = String(raw.time ?? raw.schedule ?? "");
  const optionKey = String(raw.identity ?? `${raw.course_code ?? ""}|${raw.course_name ?? raw.name ?? ""}|${raw.teacher ?? ""}|${index}`);
  const explicitDay = firstNumber(raw.weekday ?? raw.day);
  const explicitStart = firstNumber(raw.start_period ?? raw.start);
  const explicitEnd = firstNumber(raw.end_period ?? raw.end) ?? explicitStart;
  if (explicitDay && explicitStart && explicitEnd) {
    return [makeMeeting(raw, source, optionKey, `${source}-${index}-0`, explicitDay, explicitStart, explicitEnd, weeksFromRaw(raw, text), String(raw.location ?? raw.classroom ?? ""))];
  }

  const meetings: ScheduleItem[] = [];
  const pattern = /\[([^\]]+?)周\]\s*星期\s*([一二三四五六日天1-7])\s*第\s*([\d,，、\-~至]+)\s*节(?:◇([^,，◇]*))?/g;
  let match: RegExpExecArray | null;
  while ((match = pattern.exec(text)) !== null) {
    const periods = Array.from(match[3].matchAll(/\d+/g), value => Number(value[0]));
    const day = /^\d$/.test(match[2]) ? Number(match[2]) : match[2] === "天" ? 7 : chineseWeekdays.indexOf(match[2]) + 1;
    meetings.push(makeMeeting(raw, source, optionKey, `${source}-${index}-${meetings.length}`, day, periods.length ? Math.min(...periods) : null, periods.length ? Math.max(...periods) : null, expandWeekSpec(match[1]), match[4] ?? ""));
  }
  if (meetings.length) return meetings;

  const day = weekdayFrom(text);
  const periods = Array.from(text.matchAll(/第\s*(\d+)[^节]*?(\d+)?\s*节/g)).flatMap(matchValue => [Number(matchValue[1]), Number(matchValue[2])].filter(Number.isFinite));
  return [makeMeeting(raw, source, optionKey, `${source}-${index}-0`, day, periods.length ? Math.min(...periods) : null, periods.length ? Math.max(...periods) : null, weeksFromRaw(raw, text), String(raw.location ?? raw.classroom ?? ""))];
}

export function candidateOption(raw: Record<string, unknown>, index: number): CandidateOption {
  const meetings = expandScheduleItems(raw, "candidate", index);
  const creditMatch = String(raw.credits ?? raw.credit ?? "").match(/\d+(?:\.\d+)?/);
  return {
    ...raw,
    key: meetings[0].optionKey,
    name: meetings[0].name,
    category: meetings[0].category,
    teacher: meetings[0].teacher,
    courseCode: String(raw.course_code ?? ""),
    queryCode: String(raw.query_code ?? ""),
    credits: creditMatch ? Number(creditMatch[0]) : 0,
    meetings,
    unknown: meetings.every(meeting => meeting.unknown),
    executionReady: raw.execution_ready === true && String(raw.action_rwh ?? "") === String(raw.identity ?? ""),
  };
}

export const courseIdentity = (course: { courseCode?: string; code?: string; name: string }) =>
  String(course.courseCode || course.code || course.name).trim().toLowerCase();

export const courseAlreadyCompleted = (
  option: { courseCode?: string; name: string },
  completed: Array<{ code?: string; name: string }>,
) => completed.some(course =>
  Boolean(option.courseCode && course.code && option.courseCode.trim() === course.code.trim()) ||
  option.name.trim().toLowerCase() === course.name.trim().toLowerCase()
);

export function projectedCourseCredits(
  options: CandidateOption[],
  completed: Array<{ code?: string; name: string }>,
  filter: Exclude<RequirementFilter, "全部">,
): number {
  const projected = new Map<string, number>();
  for (const option of options) {
    const identity = courseIdentity(option);
    if (planningFilterFor(option) === filter && !courseAlreadyCompleted(option, completed)) {
      projected.set(identity, option.credits);
    }
  }
  return Array.from(projected.values()).reduce((total, credits) => total + credits, 0);
}

/** Candidate facts that are actually selected, from timetable or a verified execution result. */
export function selectedCandidateOptions(
  options: CandidateOption[],
  currentCourses: Array<Record<string, unknown>>,
  executions: Array<Record<string, unknown>>,
): CandidateOption[] {
  const selectedSections = new Set(executions
    .filter(item => String(item.result ?? "") === "selected")
    .map(item => String(item.section_id ?? ""))
    .filter(Boolean));
  const currentCodes = new Set(currentCourses.map(item => String(item.course_code ?? "").trim()).filter(Boolean));
  const currentNames = new Set(currentCourses.map(item => String(item.course_name ?? item.name ?? "").trim().toLowerCase()).filter(Boolean));
  const selected = new Map<string, CandidateOption>();
  for (const option of options) {
    const inTimetable = Boolean(option.courseCode && currentCodes.has(option.courseCode.trim())) || currentNames.has(option.name.trim().toLowerCase());
    if (selectedSections.has(option.key) || inTimetable) selected.set(courseIdentity(option), option);
  }
  return [...selected.values()];
}

export const activeInWeek = (item: ScheduleItem, week: number) => !item.weeks.length || item.weeks.includes(week);

export function weekItems(items: ScheduleItem[], week: number): ScheduleItem[] {
  const seen = new Set<string>();
  return items.filter(item => activeInWeek(item, week) && !item.unknown).filter(item => {
    const key = `${item.name}|${item.teacher}|${item.day}|${item.start}|${item.end}|${item.location}`;
    if (seen.has(key)) return false;
    seen.add(key);
    return true;
  });
}

const signatureKey = (item: ScheduleItem) => `${item.name.trim()}|${item.teacher.trim()}|${item.day}|${item.start}|${item.end}`;

export function weekSignature(items: ScheduleItem[], week: number): string {
  return Array.from(new Set(weekItems(items, week).map(signatureKey))).sort().join("\n");
}

export function formatWeekGroup(weeks: number[]): string {
  if (!weeks.length) return "无周次";
  if (weeks.length > 1 && weeks.every((week, index) => index === 0 || week - weeks[index - 1] === 2)) {
    return `第${weeks[0]}–${weeks[weeks.length - 1]}${weeks[0] % 2 ? "单" : "双"}周`;
  }
  const ranges: string[] = [];
  let start = weeks[0];
  let end = start;
  for (const week of weeks.slice(1)) {
    if (week === end + 1) {
      end = week;
      continue;
    }
    ranges.push(start === end ? String(start) : `${start}–${end}`);
    start = week;
    end = week;
  }
  ranges.push(start === end ? String(start) : `${start}–${end}`);
  return `第${ranges.join("、")}周`;
}

export function compressWeeks(items: ScheduleItem[], maxWeek: number): WeekGroup[] {
  const bySignature = new Map<string, number[]>();
  for (let week = 1; week <= maxWeek; week += 1) {
    const signature = weekSignature(items, week);
    bySignature.set(signature, [...(bySignature.get(signature) ?? []), week]);
  }
  return Array.from(bySignature, ([signature, weeks]) => ({
    signature,
    weeks,
    label: signature ? formatWeekGroup(weeks) : `无课程 · ${formatWeekGroup(weeks)}`,
    empty: !signature,
  })).sort((a, b) => a.weeks[0] - b.weeks[0]);
}

const exactKey = (item: ScheduleItem) => `${item.name}|${item.teacher}|${item.day}|${item.start}|${item.end}`;
const timeKey = (item: ScheduleItem) => `${item.name}|${item.teacher}`;
const courseKey = (item: ScheduleItem) => item.name;

export function diffWeeks(items: ScheduleItem[], previousWeek: number, week: number): ScheduleDiff {
  const previous = weekItems(items, previousWeek);
  const next = weekItems(items, week);
  const previousExact = new Set(previous.map(exactKey));
  const nextExact = new Set(next.map(exactKey));
  let removed = previous.filter(item => !nextExact.has(exactKey(item)));
  let added = next.filter(item => !previousExact.has(exactKey(item)));
  const timeAdjusted: string[] = [];
  const teacherAdjusted: string[] = [];

  for (const oldItem of [...removed]) {
    const matchIndex = added.findIndex(newItem => timeKey(newItem) === timeKey(oldItem));
    if (matchIndex >= 0) {
      timeAdjusted.push(oldItem.name);
      removed = removed.filter(item => item !== oldItem);
      added.splice(matchIndex, 1);
    }
  }
  for (const oldItem of [...removed]) {
    const matchIndex = added.findIndex(newItem => courseKey(newItem) === courseKey(oldItem));
    if (matchIndex >= 0) {
      teacherAdjusted.push(oldItem.name);
      removed = removed.filter(item => item !== oldItem);
      added.splice(matchIndex, 1);
    }
  }
  return {
    week,
    added: Array.from(new Set(added.map(item => item.name))),
    ended: Array.from(new Set(removed.map(item => item.name))),
    timeAdjusted: Array.from(new Set(timeAdjusted)),
    teacherAdjusted: Array.from(new Set(teacherAdjusted)),
  };
}

export function transitionsIntoGroup(items: ScheduleItem[], group: WeekGroup): ScheduleDiff[] {
  return group.weeks.filter(week => week > 1 && !group.weeks.includes(week - 1)).map(week => diffWeeks(items, week - 1, week));
}

export function locationsByWeek(items: ScheduleItem[], group: WeekGroup, target: ScheduleItem): Array<{ weeks: number[]; locations: string[] }> {
  const signature = signatureKey(target);
  const grouped = new Map<string, number[]>();
  for (const week of group.weeks) {
    const locations = weekItems(items, week).filter(item => signatureKey(item) === signature).map(item => item.location || "地点待定").sort();
    const key = locations.join("、");
    grouped.set(key, [...(grouped.get(key) ?? []), week]);
  }
  return Array.from(grouped, ([locationKey, weeks]) => ({ weeks, locations: locationKey.split("、") }));
}

const monday = (date: Date) => {
  const result = new Date(date.getFullYear(), date.getMonth(), date.getDate());
  const day = result.getDay() || 7;
  result.setDate(result.getDate() - day + 1);
  return result;
};

export function deriveCurrentWeek(calibration: WeekCalibration | null, maxWeek: number, now = new Date()): number | null {
  if (!calibration) return null;
  const anchor = new Date(`${calibration.anchorDate}T12:00:00`);
  if (Number.isNaN(anchor.getTime())) return null;
  const elapsed = Math.floor((monday(now).getTime() - monday(anchor).getTime()) / 604800000);
  const week = calibration.week + elapsed;
  return week >= 1 && week <= maxWeek ? week : null;
}
