// @vitest-environment jsdom

import React from "react";
import { cleanup, render, screen, waitFor, within } from "@testing-library/react";
import userEvent from "@testing-library/user-event";
import { afterEach, describe, expect, it, vi } from "vitest";
import type { RequirementBaseline, WorkbenchState } from "./api";
import { App } from "./main";

const guide: RequirementBaseline = {
  version: "guide-2026",
  title: "校园培养方案解读（2026 年版）",
  authority: "extracted-guide",
  requires_explicit_selection: true,
  source_summary: "项目内提纯指南",
  applicability: "不自动适用于所有年级和专业",
  coverage: "现有选修与认定学分总额",
  disclaimer: "不是学校正式毕业审核结论。",
  category_mapping: {},
  requirements: [],
};

const reference: RequirementBaseline = {
  version: "basic-graduation-reference-v1",
  title: "基础毕业要求参考基线 v1",
  authority: "reference",
  requires_explicit_selection: true,
  source_summary: "项目内提纯指南，加上用户确认的非官方参考图人工补充。",
  applicability: "用户主动选择；跨专业规则的其他年级须核对个人培养方案。",
  coverage: "基础选修及认定学分，不含专业必修和完整培养方案审核。",
  disclaimer: "非学校正式培养方案，也非学校正式毕业审核结论。",
  manual_supplements: ["创新创业单项至少 4 学分尚未证实为所有年级统一规定。"],
  category_mapping: { 创新研修课: "innovation" },
  requirements: [
    { key: "innovation", label: "创新创业", minimum: 4, unit: "credits", evidence: "grade_and_recognition" },
  ],
};

const noSelectedBaseline: WorkbenchState["selected_requirement_baseline"] = null;

const state = (selected: boolean) => ({
  login_configuration: { state: "configured", configured: true, masked_username: "2025******" },
  requirement_baselines: [guide, reference],
  selected_requirement_baseline: selected ? { ...reference, selected_at: "2026-01-01T00:00:00Z" } : noSelectedBaseline,
  recognized_credits: [],
  course_labels: [],
  labelable_courses: [],
  outside_major_track: null,
  profile: { grade: "2025" },
  confirmed_notice: null,
  snapshots: { selection: null, timetable: null, progress: null },
  snapshot_changes: { selection: null, timetable: null, progress: null },
  latest_plan: null,
  graduation_progress: {
    status: "historical",
    report: null,
    reason: "请选择要求基线后重新同步毕业进度",
    historical_report: {
      baseline_version: "guide-2026",
      data_complete: true,
      progress: [{
        key: "cultural_quality", label: "文化素质课程", required_credits: 8,
        completed_credits: 5, remaining_credits: 3, courses: [],
      }],
    },
  },
  stale: { selection: true, timetable: true, progress: true },
  snapshot_status: {
    selection: { status: "missing", reason: "尚无本地快照", source_at: "" },
    timetable: { status: "missing", reason: "尚无本地快照", source_at: "" },
    progress: { status: "historical", reason: "尚未选择要求基线", source_at: "" },
  },
  academic_session: { state: "disconnected" },
  active_task: null,
  execution_history: [],
  csrf_token: "csrf-test",
} as unknown as WorkbenchState);

const readyState = () => ({
  ...state(true),
  graduation_progress: {
    status: "ready",
    report: {
      baseline_version: "basic-graduation-reference-v1",
      data_complete: true,
      coverage: {
        grade_records: "complete", recognized_credits: "missing",
        course_classification: "missing", outside_major_track: "missing",
      },
      progress: [
        ["major_elective", "本专业选修", 3, "not_satisfied"],
        ["innovation", "创新创业", 4, "unknown"],
        ["social_practice", "社会实践", 1, "unknown"],
        ["innovation_and_practice", "创新创业 + 社会实践", 6, "unknown"],
        ["cultural_quality", "文化素质课程", 8, "unknown"],
        ["cultural_quality_d", "文化素质 D 类", 2, "unknown"],
        ["four_histories", "四史课程", 1, "unknown", "courses"],
        ["outside_major_elective", "外专业课程", 10, "unknown"],
      ].map(([key, label, minimum, condition_status, unit = "credits"]) => ({
        key, label, unit, source: key === "innovation" ? "manual-supplement" : "extracted-guide",
        parent: "", constraint: key === "outside_major_elective" ? "single_track" : "",
        rule_detail: key === "innovation" ? "人工补充参考规则" : "",
        minimum, confirmed_amount: 0, confirmed_gap: minimum, condition_status,
        condition_detail: condition_status === "not_satisfied" ? "未满足" : "未知",
        reason: condition_status === "not_satisfied" ? "confirmed_below_minimum" : "classification_evidence_missing",
        reason_detail: condition_status === "not_satisfied" ? "已确认贡献低于最低值" : "缺少明确分类依据",
        required_credits: minimum, completed_credits: 0, remaining_credits: minimum, courses: [],
      })),
    },
  },
} as unknown as WorkbenchState);

afterEach(() => {
  cleanup();
  vi.unstubAllGlobals();
});

describe("requirement baseline selection", () => {
  it("renders every baseline requirement even when selection queries cover none", async () => {
    vi.stubGlobal("fetch", vi.fn(async (input: RequestInfo | URL) => {
      const url = String(input);
      if (url === "/api/state") return new Response(JSON.stringify(readyState()));
      if (url === "/api/notices/candidates") return new Response(JSON.stringify({ notices: [] }));
      throw new Error(`unexpected request: ${url}`);
    }));
    Element.prototype.scrollIntoView = vi.fn();

    render(<App />);
    await screen.findByText("毕业进度规划参考");
    for (const label of [
      "本专业选修", "创新创业", "社会实践", "创新创业 + 社会实践",
      "文化素质课程", "文化素质 D 类", "四史课程", "跨专业发展课程",
    ]) expect(screen.getAllByText(label).length).toBeGreaterThan(0);
    expect(screen.getAllByText(/条件状态：/)).toHaveLength(8);
  });

  it("adds and removes user-declared course labels and an outside-major track", async () => {
    const local = readyState() as WorkbenchState;
    const courseIdentity = "合成 文化课程";
    const courseUrl = `/api/course-labels/${encodeURIComponent(courseIdentity)}`;
    local.labelable_courses = [
      { code: "", name: "合成   文化课程", category: "文化素质", credits: 2 },
      { code: "O01", name: "合成外专业课程", category: "外专业", credits: 10 },
    ];
    let labels: WorkbenchState["course_labels"] = [];
    let selectedTrack: WorkbenchState["outside_major_track"] = null;
    const requests: Array<{ url: string; init?: RequestInit }> = [];
    vi.stubGlobal("fetch", vi.fn(async (input: RequestInfo | URL, init?: RequestInit) => {
      const url = String(input);
      requests.push({ url, init });
      if (url === "/api/state") return new Response(JSON.stringify({ ...local, course_labels: labels, outside_major_track: selectedTrack }));
      if (url === "/api/notices/candidates") return new Response(JSON.stringify({ notices: [] }));
      if (url === courseUrl && init?.method === "PUT") {
        labels = [{ course_identity: courseIdentity, baseline_version: reference.version, updated_at: "2026-01-01T00:00:00Z", ...JSON.parse(String(init.body)) }];
        return new Response(JSON.stringify(labels[0]));
      }
      if (url === courseUrl && init?.method === "DELETE") {
        labels = []; return new Response(null, { status: 204 });
      }
      if (url === "/api/outside-major-track" && init?.method === "PUT") {
        selectedTrack = { baseline_version: reference.version, selected_at: "2026-01-01T00:00:00Z", ...JSON.parse(String(init.body)) };
        return new Response(JSON.stringify(selectedTrack));
      }
      if (url === "/api/outside-major-track" && init?.method === "DELETE") {
        selectedTrack = null; return new Response(null, { status: 204 });
      }
      throw new Error(`unexpected request: ${url}`);
    }));
    Element.prototype.scrollIntoView = vi.fn();
    const user = userEvent.setup();
    render(<App />);
    await screen.findByText("选课规划工作台");
    await user.click(screen.getByRole("button", { name: /数据操作/ }));

    await user.selectOptions(screen.getByLabelText("选择已完成课程"), courseIdentity);
    await user.click(screen.getByLabelText("D 类"));
    await user.click(screen.getByLabelText("四史"));
    await user.type(screen.getByLabelText("课程所属外专业体系"), "track-a");
    await user.click(screen.getByRole("button", { name: "保存课程标签" }));
    expect(await screen.findByText(/D 类 · 四史 · 体系：track-a/)).toBeTruthy();
    const labelRequest = requests.find(request => request.url === courseUrl && request.init?.method === "PUT");
    expect(labelRequest?.init?.headers).toMatchObject({ "X-CSRF-Token": "csrf-test" });

    await user.type(screen.getByLabelText("当前外专业课程体系"), "track-a");
    await user.click(screen.getByRole("button", { name: "保存体系" }));
    expect(await screen.findByRole("button", { name: "清除体系" })).toBeTruthy();
    await user.click(screen.getByRole("button", { name: "清除体系" }));
    expect(await screen.findByText("尚未选择体系，外专业条件保持未知。")).toBeTruthy();
    await user.click(screen.getByRole("button", { name: "删除" }));
    await waitFor(() => expect(labels).toHaveLength(0));
  });

  it("does not double-count declarations linked to enrolled or queued courses", async () => {
    const overlap = readyState() as WorkbenchState;
    const innovation = overlap.graduation_progress.report!.progress.find(item => item.key === "innovation")!;
    Object.assign(innovation, {
      declared_amount: 4, estimated_amount: 4, estimated_gap: 0, manual_review_required: true,
      declarations: [
        { identity: "declared-enrolled", note: "关联已选", credits: 2, linked_course_identity: "ENROLLED", contributes: true },
        { identity: "declared-queued", note: "关联队列", credits: 2, linked_course_identity: "QUEUED", contributes: true },
      ],
    });
    overlap.recognized_credits = [];
    overlap.snapshots.timetable = {
      id: "timetable-1", kind: "timetable", term: "2026-1", source: "test",
      source_at: "2026-01-01", payload: { entries: [], enrolled_courses: [
        { code: "ENROLLED", name: "本学期创新", category: "创新研修课", nature: "任选", credits: 2 },
      ] },
    };
    overlap.snapshots.selection = {
      id: "selection-1", kind: "selection", term: "2026-1", source: "test",
      source_at: "2026-01-01", payload: { sections: [
        { identity: "section-queued", course_code: "QUEUED", course_name: "队列创新", category: "创新研修课", credits: 2 },
      ] },
    };
    overlap.latest_plan = { goals: [{ goal_id: "goal-1", course_identity: "QUEUED", rank: 1, preferences: [{ section_id: "section-queued", rank: 1 }] }] };
    vi.stubGlobal("fetch", vi.fn(async (input: RequestInfo | URL) => {
      const url = String(input);
      if (url === "/api/state") return new Response(JSON.stringify(overlap));
      if (url === "/api/notices/candidates") return new Response(JSON.stringify({ notices: [] }));
      throw new Error(`unexpected request: ${url}`);
    }));
    Element.prototype.scrollIntoView = vi.fn();

    render(<App />);
    const meter = await screen.findByRole("progressbar", { name: "创新创业预计学分" });
    const card = meter.closest("article")!;
    expect(card.textContent).toContain("4 / 4 学分");
    expect(within(card).getByText(/用户申报 0/)).toBeTruthy();
    expect(within(card).getByText(/本学期已选 2/)).toBeTruthy();
    expect(within(card).getByText(/队列预览 2/)).toBeTruthy();
    expect(within(card).getByText(/已确认缺口：4 学分/)).toBeTruthy();
    expect(within(card).getAllByText(/已关联本学期或队列课程，不重复计入/)).toHaveLength(2);
    expect(within(card).getByText(/可能重叠，请人工核验/)).toBeTruthy();
  });

  it("creates, edits, and deletes a local recognized-credit declaration", async () => {
    let declarations: WorkbenchState["recognized_credits"] = [];
    const requests: Array<{ url: string; init?: RequestInit }> = [];
    vi.stubGlobal("fetch", vi.fn(async (input: RequestInfo | URL, init?: RequestInit) => {
      const url = String(input);
      requests.push({ url, init });
      if (url === "/api/state") return new Response(JSON.stringify({ ...state(true), recognized_credits: declarations }));
      if (url === "/api/notices/candidates") return new Response(JSON.stringify({ notices: [] }));
      if (url === "/api/recognized-credits" && init?.method === "POST") {
        const body = JSON.parse(String(init.body));
        declarations = [{ ...body, baseline_version: reference.version, linked_course_identity: "", created_at: "2026-01-01T00:00:00Z", updated_at: "2026-01-01T00:00:00Z" }];
        return new Response(JSON.stringify(declarations[0]), { status: 201 });
      }
      if (url.startsWith("/api/recognized-credits/") && init?.method === "PUT") {
        declarations = [{ ...declarations[0], ...JSON.parse(String(init.body)), updated_at: "2026-01-02T00:00:00Z" }];
        return new Response(JSON.stringify(declarations[0]));
      }
      if (url.startsWith("/api/recognized-credits/") && init?.method === "DELETE") {
        declarations = [];
        return new Response(null, { status: 204 });
      }
      throw new Error(`unexpected request: ${url}`);
    }));
    Element.prototype.scrollIntoView = vi.fn();
    const user = userEvent.setup();
    render(<App />);
    await screen.findByText("选课规划工作台");
    await user.click(screen.getByRole("button", { name: /数据操作/ }));

    await user.selectOptions(screen.getByLabelText("申报类别"), "social_practice");
    await user.clear(screen.getByLabelText("申报学分"));
    await user.type(screen.getByLabelText("申报学分"), "2");
    await user.type(screen.getByLabelText("申报说明"), "合成社会实践认定");
    await user.click(screen.getByRole("button", { name: "新增申报" }));
    expect(await screen.findByText("合成社会实践认定")).toBeTruthy();
    const createRequest = requests.find(request => request.url === "/api/recognized-credits" && request.init?.method === "POST");
    expect(createRequest?.init?.headers).toMatchObject({ "X-CSRF-Token": "csrf-test" });

    await user.click(screen.getByRole("button", { name: "编辑" }));
    await user.clear(screen.getByLabelText("申报学分"));
    await user.type(screen.getByLabelText("申报学分"), "1");
    await user.click(screen.getByRole("button", { name: "保存修改" }));
    await waitFor(() => expect(requests.some(request => request.url.startsWith("/api/recognized-credits/") && request.init?.method === "PUT")).toBe(true));
    expect(await screen.findByText(/1 学分/)).toBeTruthy();

    await user.click(screen.getByRole("button", { name: "删除" }));
    expect(await screen.findByText("暂无申报。不支持用手填学分证明四史门数。")).toBeTruthy();
  });

  it("shows applicability and historical progress before explicitly confirmed local selection", async () => {
    let selected = false;
    const requests: Array<{ url: string; init?: RequestInit }> = [];
    vi.stubGlobal("fetch", vi.fn(async (input: RequestInfo | URL, init?: RequestInit) => {
      const url = String(input);
      requests.push({ url, init });
      if (url === "/api/state") return new Response(JSON.stringify(state(selected)));
      if (url === "/api/notices/candidates") return new Response(JSON.stringify({ notices: [] }));
      if (url === "/api/requirement-baseline-selection") {
        selected = true;
        return new Response(JSON.stringify({ selected_requirement_baseline: reference }));
      }
      throw new Error(`unexpected request: ${url}`);
    }));

    Element.prototype.scrollIntoView = vi.fn();
    const user = userEvent.setup();
    render(<App />);
    await screen.findByText("选课规划工作台");
    await user.click(screen.getByRole("button", { name: /数据操作/ }));

    const option = screen.getByText(reference.title).closest("article");
    expect(option).not.toBeNull();
    expect(within(option!).getByText(reference.applicability)).toBeTruthy();
    expect(within(option!).getByText(reference.disclaimer)).toBeTruthy();
    await user.click(within(option!).getByRole("button", { name: "选择" }));

    const dialog = screen.getByRole("alertdialog");
    expect(within(dialog).getByText(/跨专业规则的其他年级须核对/)).toBeTruthy();
    expect(within(dialog).getByText(/创新创业单项至少 4 学分尚未证实/)).toBeTruthy();
    expect(within(dialog).getByText(/非学校正式培养方案/)).toBeTruthy();
    await user.click(within(dialog).getByRole("button", { name: "确认选择" }));

    await waitFor(() => expect(requests.some(request =>
      request.url === "/api/requirement-baseline-selection" &&
      request.init?.method === "POST" &&
      request.init.body === JSON.stringify({
        version: "basic-graduation-reference-v1",
        confirmation: "basic-graduation-reference-v1",
      })
    )).toBe(true));
    expect(await screen.findByText("查看旧基线进度")).toBeTruthy();
    await user.click(screen.getByText("查看旧基线进度"));
    expect(screen.getByText(/文化素质课程：已确认 5 \/ 8 学分/)).toBeTruthy();
  });
});
