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
