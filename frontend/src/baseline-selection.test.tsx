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
    { key: "innovation", label: "创新创业", minimum: 4, unit: "credits" },
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

afterEach(() => {
  cleanup();
  vi.unstubAllGlobals();
});

describe("requirement baseline selection", () => {
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
