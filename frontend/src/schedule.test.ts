import { describe, expect, it } from "vitest";
import {
  compressWeeks,
  deriveCurrentWeek,
  describeOptionMeetings,
  diffWeeks,
  expandScheduleItems,
  expandWeekSpec,
  formatWeekGroup,
  requirementFilterFor,
  planningFilterFor,
  queryFilterFor,
  projectedCourseCredits,
  candidateOption,
  weekSignature,
  type ScheduleItem,
} from "./schedule";

const item = (overrides: Partial<ScheduleItem> = {}): ScheduleItem => ({
  source: "current",
  optionKey: "course-a",
  key: "course-a-0",
  day: 1,
  start: 1,
  end: 2,
  weeks: [1, 2, 3],
  unknown: false,
  name: "高等数学",
  category: "专业基础课",
  teacher: "张老师",
  location: "N楼-101",
  ...overrides,
});

describe("schedule parsing", () => {
  it("maps raw course labels to the smaller planning filters", () => {
    expect(requirementFilterFor("创新实验课")).toBe("创新创业");
    expect(requirementFilterFor("文理通识-文化素质教育课")).toBe("文化素质");
    expect(requirementFilterFor("外专业课程")).toBe("跨专业");
    // 白名单里没有本专业查询：专业*标签只可能来自外专业查询。
    expect(requirementFilterFor("专业核心课")).toBe("跨专业");
    expect(requirementFilterFor("专业基础课（包括大类平台课）")).toBe("跨专业");
    expect(requirementFilterFor("公共选修课")).toBeNull();
  });

  it("marks only exact page-provided rwh identities executable", () => {
    expect(candidateOption({ identity: "RWH-1", action_rwh: "RWH-1", execution_ready: true, name: "课程" }, 0).executionReady).toBe(true);
    expect(candidateOption({ identity: "fallback", action_rwh: "", execution_ready: false, name: "课程" }, 0).executionReady).toBe(false);
    expect(candidateOption({ identity: "RWH-1", action_rwh: "OTHER", execution_ready: true, name: "课程" }, 0).executionReady).toBe(false);
  });

  it("projects distinct uncompleted course credits", () => {
    const options = [
      candidateOption({ identity: "section-a", course_code: "ART1", name: "艺术史", category: "文化素质", credits: "2.0" }, 0),
      candidateOption({ identity: "section-b", course_code: "ART1", name: "艺术史", category: "文化素质", credits: "2.0" }, 1),
      candidateOption({ identity: "section-c", course_code: "HIST", name: "四史专题", category: "文化素质", credits: "2.0" }, 2),
    ];
    expect(projectedCourseCredits(options, [{ code: "HIST", name: "四史专题" }], "文化素质")).toBe(2);
  });

  it("maps query source codes to planning filters ahead of raw labels", () => {
    const ownLabeled = candidateOption({ identity: "om-1", course_code: "CS1", name: "外校数据结构", category: "专业核心课", query_code: "xsxk", credits: "3.0" }, 0);
    expect(queryFilterFor("xsxk")).toBe("跨专业");
    expect(queryFilterFor("unknown")).toBeNull();
    expect(planningFilterFor(ownLabeled)).toBe("跨专业");
    expect(planningFilterFor(candidateOption({ identity: "ty-1", course_code: "PE1", name: "体育（3）", category: "体育", credits: "0.5" }, 0))).toBe("体育");
    expect(planningFilterFor(candidateOption({ identity: "raw-1", course_code: "ZZ1", name: "某课", category: "公共选修课", credits: "2.0" }, 0))).toBeNull();
  });

  it("projects other-major credits through the query source code", () => {
    const options = [
      candidateOption({ identity: "om-a", course_code: "ECO1", name: "经济学原理", category: "专业基础课（包括大类平台课）", query_code: "xsxk", credits: "3.0" }, 0),
      candidateOption({ identity: "om-b", course_code: "ECO1", name: "经济学原理", category: "专业基础课（包括大类平台课）", query_code: "xsxk", credits: "3.0" }, 1),
    ];
    expect(projectedCourseCredits(options, [], "跨专业")).toBe(3);
  });

  it("expands ranges and parity", () => {
    expect(expandWeekSpec("1-12双")).toEqual([2, 4, 6, 8, 10, 12]);
    expect(expandWeekSpec("1-5,7-9")).toEqual([1, 2, 3, 4, 5, 7, 8, 9]);
  });

  it("describes candidate meetings independently of the selected week", () => {
    const earlier = expandScheduleItems({ identity: "t1", name: "体育（3）", teacher: "王老师", time: "[1-8周]星期一第5,6节◇体育馆-101" }, "candidate", 0);
    const later = expandScheduleItems({ identity: "t2", name: "体育（3）", teacher: "王老师", time: "[10-17周]星期三第3,4节◇体育馆-202" }, "candidate", 0);
    const first = describeOptionMeetings(earlier);
    const second = describeOptionMeetings(later);
    expect(first).toContain("周一第5–6节 体育馆-101");
    expect(first).toContain("第1–8周");
    expect(second).toContain("周三第3–4节 体育馆-202");
    expect(second).toContain("第10–17周");
    expect(first).not.toBe(second);
    expect(describeOptionMeetings(expandScheduleItems({ identity: "t3", name: "某课", teacher: "某师", time: "" }, "candidate", 0))).toBe("上课时间待定");
  });

  it("expands every meeting in a candidate course", () => {
    const meetings = expandScheduleItems({
      identity: "candidate-1",
      name: "离散数学",
      teacher: "李老师",
      time: "[1-5,7-9周]星期二第5,6节◇M楼-201,[10-17周]星期四第7,8节◇N楼-301",
    }, "candidate", 0);
    expect(meetings).toHaveLength(2);
    expect(meetings[0]).toMatchObject({ day: 2, start: 5, end: 6, location: "M楼-201", weeks: [1, 2, 3, 4, 5, 7, 8, 9] });
    expect(meetings[1]).toMatchObject({ day: 4, start: 7, end: 8, location: "N楼-301", weeks: [10, 11, 12, 13, 14, 15, 16, 17] });
  });

  it("separates teacher annotations from imported timetable locations", () => {
    const [meeting] = expandScheduleItems({
      course_name: "离散数学",
      weekday: 2,
      start_period: 5,
      end_period: 6,
      week_numbers: [1, 2],
      location: "，黄俊恒[1-5，7-9] M楼-201",
    }, "current", 0);
    expect(meeting).toMatchObject({ teacher: "黄俊恒", location: "M楼-201" });
  });
});

describe("week compression", () => {
  it("ignores location changes but keeps non-contiguous equal weeks together", () => {
    const entries = [
      item({ key: "a-1", weeks: [1, 2, 3, 4, 5], location: "N楼-101" }),
      item({ key: "a-2", weeks: [7, 8, 9], location: "M楼-202" }),
      item({ key: "b", name: "大学物理", teacher: "王老师", weeks: [6] }),
    ];
    const groups = compressWeeks(entries, 9);
    expect(groups.find(group => group.weeks.includes(1))?.weeks).toEqual([1, 2, 3, 4, 5, 7, 8, 9]);
    expect(groups.find(group => group.weeks.includes(6))?.weeks).toEqual([6]);
    expect(formatWeekGroup([1, 2, 3, 4, 5, 7, 8, 9])).toBe("第1–5、7–9周");
  });

  it("includes teacher changes in the signature", () => {
    const entries = [item({ weeks: [1], teacher: "张老师" }), item({ key: "a-2", weeks: [2], teacher: "李老师" })];
    expect(weekSignature(entries, 1)).not.toBe(weekSignature(entries, 2));
  });
});

describe("week differences and calibration", () => {
  it("classifies a time adjustment without counting add and end", () => {
    const entries = [item({ weeks: [1], start: 1, end: 2 }), item({ key: "a-2", weeks: [2], start: 3, end: 4 })];
    expect(diffWeeks(entries, 1, 2)).toMatchObject({ added: [], ended: [], timeAdjusted: ["高等数学"] });
  });

  it("advances the calibrated week every Monday and rejects overflow", () => {
    const calibration = { term: "2026秋季", anchorDate: "2026-08-26", week: 1 };
    expect(deriveCurrentWeek(calibration, 18, new Date("2026-09-02T12:00:00"))).toBe(2);
    expect(deriveCurrentWeek(calibration, 2, new Date("2026-09-16T12:00:00"))).toBeNull();
  });
});
