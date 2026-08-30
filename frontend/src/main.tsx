import React, { useCallback, useEffect, useRef, useState } from "react";
import { createRoot } from "react-dom/client";
import type { CandidateNotice, Task, WorkbenchState } from "./api";
import { ScheduleBoard } from "./ScheduleBoard";
import { candidateExecutionStatus, expandScheduleItems, selectionCategoryLabel, selectionWindowDisplay, selectionWindowsForGrade, type SelectionWindow } from "./schedule";
import "./style.css";

const jsonHeaders = (token: string) => ({ "Content-Type": "application/json", "X-CSRF-Token": token });

type PlanPreference = { section_id: string; rank: number };
type PlanGoal = { goal_id: string; course_identity: string; rank: number; preferences: PlanPreference[] };
type PlanSection = Record<string, unknown>;

type PlanCourse = { identity: string; name: string; category: string; sections: PlanSection[] };

const planReasonLabels: Record<string, string> = {
  goals_missing: "尚未添加任何课程目标",
  timetable_snapshot_missing: "缺少课表快照，请先刷新课表",
  timetable_snapshot_incomplete: "课表快照不完整，请重新刷新",
  timetable_term_mismatch: "课表与当前学期不匹配",
  timetable_profile_mismatch: "课表与学生画像不匹配",
  timetable_snapshot_stale: "课表快照已过期，请刷新课表",
  timetable_source_time_invalid: "课表快照时间无效，请重新刷新",
  timetable_snapshot_unusable: "课表快照不可用，请重新刷新",
  selection_snapshot_missing: "缺少待选课程快照，请先刷新待选课程",
  selection_snapshot_incomplete: "待选课程快照不完整，请重新刷新",
  selection_term_mismatch: "待选课程与当前学期不匹配",
  selection_profile_mismatch: "待选课程与学生画像不匹配",
  selection_notice_mismatch: "待选课程与已确认通知不匹配，请强制刷新待选课程",
  selection_snapshot_stale: "待选课程快照已过期，请强制刷新",
  selection_source_time_invalid: "待选课程快照时间无效",
  selection_snapshot_unusable: "待选课程快照不可用",
  conflict_unknown: "存在时间未知的教学班，无法确认是否冲突",
};
const planReasonLabel = (reason: string) =>
  planReasonLabels[reason] ?? (reason.startsWith("section_missing:") ? `教学班 ${reason.slice("section_missing:".length)} 不在当前待选课程中，请重新规划` : reason);
const planConflictKindLabels: Record<string, string> = {
  current_timetable: "与当前课表冲突",
  candidate_sections: "与规划内其他教学班冲突",
  conflict_unknown: "时间未知，无法确认是否冲突",
};
const sectionTimeText = (section: PlanSection) => String(section.time ?? section.schedule ?? "时间待定");
const sectionIdentity = (section: PlanSection) => String(section.identity ?? section.section_id ?? "");
const sectionLabel = (section: PlanSection) => {
  const teacher = String(section.teacher ?? "").trim();
  return [teacher && `教师 ${teacher}`, sectionTimeText(section)].filter(Boolean).join(" · ");
};

type QueueResult = { goal: string; sectionId: string; sectionLabel: string; outcome: "submitted" | "skipped" | "failed"; message: string };

const timeOverlaps = (left: { day: number | null; start: number | null; end: number | null }, right: { day: number | null; start: number | null; end: number | null }) =>
  left.day !== null && left.day === right.day && left.start !== null && right.start !== null && left.end !== null && right.end !== null && left.start <= right.end && right.start <= left.end;

const waitForTask = (task: Task): Promise<Task> => new Promise(resolve => {
  const timer = window.setInterval(async () => {
    try {
      const response = await fetch(`/api/tasks/${task.id}`);
      if (!response.ok) { window.clearInterval(timer); resolve(task); return; }
      const next = await response.json() as Task;
      if (["succeeded", "failed", "cancelled"].includes(next.state)) { window.clearInterval(timer); resolve(next); }
    } catch { window.clearInterval(timer); resolve(task); }
  }, 600);
});

type PendingConfirm = { title: string; detail?: string; confirmLabel?: string; onConfirm: () => void };

type QueueStep = { goal: string; attempts: Array<{ sectionId: string; label: string }> };

function ConfirmDialog({ pending, onConfirm, onCancel }: { pending: PendingConfirm; onConfirm: () => void; onCancel: () => void }) {
  const confirmRef = useRef<HTMLButtonElement>(null);
  useEffect(() => {
    confirmRef.current?.focus();
    const closeOnEscape = (event: KeyboardEvent) => {
      if (event.key === "Escape") onCancel();
    };
    window.addEventListener("keydown", closeOnEscape);
    return () => window.removeEventListener("keydown", closeOnEscape);
  }, [onCancel]);
  return <div className="confirm-layer">
    <button className="confirm-backdrop" onClick={onCancel} aria-label="取消" />
    <div className="confirm-dialog" role="alertdialog" aria-modal="true" aria-labelledby="confirm-title">
      <h3 id="confirm-title">{pending.title}</h3>
      {pending.detail && <div className="confirm-detail">{pending.detail}</div>}
      <div className="confirm-actions"><button className="secondary" onClick={onCancel}>取消</button><button ref={confirmRef} onClick={onConfirm}>{pending.confirmLabel ?? "确认"}</button></div>
    </div>
  </div>;
}

function App() {
  const [state, setState] = useState<WorkbenchState | null>(null);
  const [candidates, setCandidates] = useState<CandidateNotice[]>([]);
  const [task, setTask] = useState<Task | null>(null);
  const [goals, setGoals] = useState<PlanGoal[]>([]);
  const [pendingCourse, setPendingCourse] = useState("");
  const [queueRunning, setQueueRunning] = useState(false);
  const [queueResults, setQueueResults] = useState<QueueResult[]>([]);
  const [dragGoalIndex, setDragGoalIndex] = useState<number | null>(null);
  const [pendingConfirm, setPendingConfirm] = useState<PendingConfirm | null>(null);
  const lastSelectionTerm = useRef<string | null>(null);
  const [plan, setPlan] = useState<Record<string, unknown> | null>(null);
  const [message, setMessage] = useState("");
  const [username, setUsername] = useState("");
  const [password, setPassword] = useState("");
  const [savingLogin, setSavingLogin] = useState(false);
  const [dataDrawerOpen, setDataDrawerOpen] = useState(false);
  const [windowsOpen, setWindowsOpen] = useState(false);
  const drawerTriggerRef = useRef<HTMLButtonElement>(null);
  const drawerCloseRef = useRef<HTMLButtonElement>(null);
  const goalsHydrated = useRef(false);

  const askConfirm = (title: string, detail: string | undefined, confirmLabel: string | undefined, onConfirm: () => void) => setPendingConfirm({ title, detail, confirmLabel, onConfirm });

  const load = useCallback(async () => {
    const [response, notices] = await Promise.all([
      fetch("/api/state"),
      fetch("/api/notices/candidates"),
    ]);
    const [next, noticeResult] = await Promise.all([
      response.json() as Promise<WorkbenchState>,
      notices.json() as Promise<{ notices?: CandidateNotice[] }>,
    ]);
    setState(next);
    if (next.active_task) setTask(next.active_task);
    setPlan(next.latest_plan);
    if (next.latest_plan?.goals && !goalsHydrated.current) {
      setGoals(next.latest_plan.goals as PlanGoal[]);
      goalsHydrated.current = true;
    }
    setCandidates(noticeResult.notices ?? []);
  }, []);

  useEffect(() => { load().catch((error: unknown) => setMessage(String(error))); }, [load]);

  useEffect(() => {
    if (!dataDrawerOpen) return;
    const closeOnEscape = (event: KeyboardEvent) => {
      if (event.key === "Escape") setDataDrawerOpen(false);
    };
    document.body.classList.add("drawer-open");
    drawerCloseRef.current?.focus();
    window.addEventListener("keydown", closeOnEscape);
    return () => {
      document.body.classList.remove("drawer-open");
      drawerTriggerRef.current?.focus();
      window.removeEventListener("keydown", closeOnEscape);
    };
  }, [dataDrawerOpen]);

  useEffect(() => {
    if (state?.snapshots.selection?.term && state.snapshots.selection.term !== lastSelectionTerm.current) {
      lastSelectionTerm.current = state.snapshots.selection.term;
      setGoals([]);
      setQueueResults([]);
    }
  }, [state?.snapshots.selection?.term]);

  const configureLogin = async (event: React.FormEvent) => {
    event.preventDefault();
    if (!state || savingLogin) return;
    setSavingLogin(true);
    setMessage("");
    try {
      const response = await fetch("/api/login-configuration", {
        method: "POST",
        headers: jsonHeaders(state.csrf_token),
        body: JSON.stringify({ username, password }),
      });
      const result = await response.json();
      if (!response.ok) {
        setMessage(result.error ?? "无法保存自动登录配置");
        return;
      }
      setPassword("");
      setMessage("登录配置已保存；需要访问教务时请点击“连接教务”。");
      await load();
    } finally {
      setSavingLogin(false);
    }
  };

  const clearLogin = () => {
    if (!state) return;
    askConfirm("清除自动登录？", "清除自动登录并重置当前学生的画像、课程快照和规划；官方通知会保留。", undefined, () => void clearLoginConfirmed());
  };

  const clearLoginConfirmed = async () => {
    if (!state) return;
    const response = await fetch("/api/login-configuration", {
      method: "DELETE",
      headers: jsonHeaders(state.csrf_token),
    });
    if (response.ok) {
      setTask(null);
      setUsername("");
      setPassword("");
      await load();
    } else setMessage("无法清除自动登录配置");
  };

  const submitTask = async (operation: string) => {
    if (!state) return;
    const response = await fetch("/api/tasks", {
      method: "POST",
      headers: jsonHeaders(state.csrf_token),
      body: JSON.stringify({ operation }),
    });
    const created = await response.json();
    if (!response.ok) {
      setMessage(created.error ?? "任务提交失败");
      return;
    }
    setTask(created);
  };

  const run = (operation: string) => {
    if (!state) return;
    if (operation === "refresh-selection") {
      const status = state.snapshot_status?.selection;
      const windows = (state.confirmed_notice?.windows as Array<{ category_codes?: string[] }> | undefined) ?? [];
      const categories = windows.flatMap(window => window.category_codes ?? []);
      askConfirm("强制刷新待选课程？", `上次完整刷新：${formatSourceTime(status?.source_at)}\n查询类别：${new Set(categories).size}\n\n这会访问学校系统。`, undefined, () => void submitTask("refresh-selection"));
      return;
    }
    void submitTask(operation);
  };

  useEffect(() => {
    if (!task || !state || ["succeeded", "failed", "cancelled"].includes(task.state)) return;
    const timer = window.setInterval(async () => {
      const response = await fetch(`/api/tasks/${task.id}`);
      if (response.ok) {
        const next = await response.json() as Task;
        setTask(next);
        if (["succeeded", "failed", "cancelled"].includes(next.state)) await load();
      }
    }, 600);
    return () => window.clearInterval(timer);
  }, [task, state, load]);

  const discoverOfficialNotices = async () => {
    if (!state) return;
    const response = await fetch("/api/notices/discover", {
      method: "POST",
      headers: jsonHeaders(state.csrf_token),
      body: JSON.stringify({ index_url: "https://jwc.hitwh.edu.cn/ks/list.htm" }),
    });
    const result = await response.json();
    setMessage(response.ok ? `已从教务处获取并解析 ${result.notices?.length ?? 0} 份候选通知` : (result.error ?? "通知获取失败"));
    if (response.ok) await load();
  };

  const confirm = async (id: string) => {
    if (!state) return;
    const response = await fetch(`/api/notices/${id}/confirm`, {
      method: "POST",
      headers: jsonHeaders(state.csrf_token),
    });
    const result = await response.json();
    setMessage(response.ok ? "通知已确认；需要时请强制刷新待选课程" : (result.error ?? "确认失败"));
    await load();
  };

  const cancel = async () => {
    if (task && state) {
      await fetch(`/api/tasks/${task.id}`, {
        method: "DELETE",
        headers: { "X-CSRF-Token": state.csrf_token },
      });
    }
  };

  const finishObservation = async () => {
    if (!task || !state) return;
    const response = await fetch(`/api/tasks/${task.id}/finish`, {
      method: "POST",
      headers: jsonHeaders(state.csrf_token),
    });
    if (!response.ok) setMessage("当前任务无法完成监听");
  };

  const executeSection = (sectionId: string, courseName: string) => {
    if (!state?.snapshots.selection) return;
    askConfirm("确认提交一次选课请求？", `${courseName}\n教学班：${sectionId}\n\n系统不会自动重试。`, undefined, () => void submitSection(sectionId));
  };

  const submitSection = async (sectionId: string) => {
    if (!state?.snapshots.selection) return;
    const response = await fetch("/api/executions/selection", {
      method: "POST",
      headers: jsonHeaders(state.csrf_token),
      body: JSON.stringify({
        section_id: sectionId,
        snapshot_id: state.snapshots.selection.id,
        confirmation: sectionId,
      }),
    });
    const result = await response.json();
    if (!response.ok) {
      setMessage(result.error ?? "选课请求未提交");
      return;
    }
    setTask(result as Task);
    setMessage("已提交一次选课任务；不会自动重试。");
  };

  const resolveUnknown = async (identity: string) => {
    if (!state) return;
    await fetch(`/api/executions/${identity}/resolve`, { method: "POST", headers: jsonHeaders(state.csrf_token) });
    await load();
  };

  const clearHistory = () => {
    if (!state) return;
    askConfirm("清除全部本地选课执行历史？", undefined, undefined, () => void clearHistoryConfirmed());
  };

  const clearHistoryConfirmed = async () => {
    if (!state) return;
    await fetch("/api/executions", { method: "DELETE", headers: jsonHeaders(state.csrf_token) });
    await load();
  };

  const savePlan = async () => {
    if (!state || !goals.length) return;
    const response = await fetch("/api/plans", {
      method: "POST",
      headers: jsonHeaders(state.csrf_token),
      body: JSON.stringify({ goals }),
    });
    const result = await response.json();
    if (!response.ok) {
      setMessage(result.error ?? "规划保存失败");
      return;
    }
    setPlan(result);
    setMessage(result.status === "ready" ? "只读规划已保存" : "规划已保存，但当前仍被数据条件阻断");
  };

  const addGoal = () => {
    if (!pendingCourse) return;
    const course = planCourses.find(candidate => candidate.identity === pendingCourse);
    if (!course) return;
    const first = course.sections[0];
    setGoals(list => [...list, {
      goal_id: `goal-${Date.now()}-${list.length}`,
      course_identity: pendingCourse,
      rank: list.length + 1,
      preferences: first ? [{ section_id: sectionIdentity(first), rank: 1 }] : [],
    }]);
    setPendingCourse("");
  };

  const toggleQueueSection = (sectionId: string) => {
    const section = planSections.find(candidate => sectionIdentity(candidate) === sectionId);
    if (!section) return;
    const courseIdentity = String(section.course_code ?? section.course_name ?? section.name ?? "未命名课程");
    setGoals(list => {
      const goalIndex = list.findIndex(goal => goal.course_identity === courseIdentity);
      if (goalIndex < 0) {
        return [...list, { goal_id: `goal-${Date.now()}-${list.length}`, course_identity: courseIdentity, rank: list.length + 1, preferences: [{ section_id: sectionId, rank: 1 }] }];
      }
      const goal = list[goalIndex];
      if (goal.preferences[0]?.section_id === sectionId) {
        return list.filter((_, index) => index !== goalIndex);
      }
      const next = [...list];
      next[goalIndex] = { ...goal, preferences: [{ section_id: sectionId, rank: 1 }, ...goal.preferences.filter(preference => preference.section_id !== sectionId)] };
      return next;
    });
  };

  const changeGoalCourse = (index: number, identity: string) => {
    setGoals(list => {
      const next = [...list];
      const first = planCourses.find(candidate => candidate.identity === identity)?.sections[0];
      next[index] = {
        ...next[index],
        course_identity: identity,
        preferences: first ? [{ section_id: sectionIdentity(first), rank: 1 }] : [],
      };
      return next;
    });
  };

  const removeGoal = (index: number) => setGoals(list => list.filter((_, itemIndex) => itemIndex !== index));

  const moveGoal = (index: number, delta: number) => setGoals(list => {
    const target = index + delta;
    if (target < 0 || target >= list.length) return list;
    const next = [...list];
    [next[index], next[target]] = [next[target], next[index]];
    return next;
  });

  const moveGoalTo = (from: number, to: number) => setGoals(list => {
    if (from === to || from < 0 || from >= list.length || to < 0 || to >= list.length) return list;
    const next = [...list];
    const [moved] = next.splice(from, 1);
    next.splice(to, 0, moved);
    return next;
  });

  const addPreference = (goalIndex: number, sectionId: string) => {
    if (!sectionId) return;
    setGoals(list => {
      const next = [...list];
      const goal = next[goalIndex];
      if (!goal || goal.preferences.some(preference => preference.section_id === sectionId)) return list;
      next[goalIndex] = { ...goal, preferences: [...goal.preferences, { section_id: sectionId, rank: goal.preferences.length + 1 }] };
      return next;
    });
  };

  const removePreference = (goalIndex: number, preferenceIndex: number) => setGoals(list => {
    const next = [...list];
    const goal = next[goalIndex];
    if (!goal) return list;
    next[goalIndex] = { ...goal, preferences: goal.preferences.filter((_, itemIndex) => itemIndex !== preferenceIndex) };
    return next;
  });

  const movePreference = (goalIndex: number, preferenceIndex: number, delta: number) => setGoals(list => {
    const next = [...list];
    const goal = next[goalIndex];
    if (!goal) return list;
    const target = preferenceIndex + delta;
    if (target < 0 || target >= goal.preferences.length) return list;
    const preferences = [...goal.preferences];
    [preferences[preferenceIndex], preferences[target]] = [preferences[target], preferences[preferenceIndex]];
    next[goalIndex] = { ...goal, preferences };
    return next;
  });

  const planConflictLabel = (conflict: Record<string, unknown>) => {
    const sectionId = String(conflict.section_id ?? "");
    const section = planSections.find(candidate => sectionIdentity(candidate) === sectionId);
    const name = String(section?.course_name ?? section?.name ?? sectionId);
    const kind = planConflictKindLabels[String(conflict.kind ?? "")] ?? String(conflict.kind ?? "未知");
    const withName = conflict.with_id ? `（与 ${sectionName(String(conflict.with_id))}）` : "";
    return `${name}：${kind}${withName}`;
  };
  const sectionName = (id: string) => {
    const section = planSections.find(candidate => sectionIdentity(candidate) === id);
    return String(section?.course_name ?? section?.name ?? id);
  };

  const sectionExecutionBlocked = (section: PlanSection): string | null => {
    if (String(section.execution_ready) !== "true" || String(section.action_rwh ?? "") !== sectionIdentity(section)) return "教学班缺少可执行身份，请先刷新待选课程";
    const status = candidateExecutionStatus({ queryCode: String(section.query_code ?? ""), executionReady: true }, selectionWindows, grade);
    if (!status.canExecute) return status.reason;
    const meetings = expandScheduleItems(section as Record<string, unknown>, "candidate", 0).filter(item => !item.unknown);
    if (!meetings.length) return "上课时间待确认";
    const overlapsCurrent = meetings.some(meeting => currentScheduleItems.some(other => !other.unknown && timeOverlaps(meeting, other) && (!meeting.weeks.length || !other.weeks.length || meeting.weeks.some(week => other.weeks.includes(week)))));
    if (overlapsCurrent) return "与当前课表冲突";
    return null;
  };

  const runQueue = () => {
    if (!state || queueRunning || remoteBusy || !goals.length) return;
    const snapshotId = state.snapshots.selection.id;
    const steps: QueueStep[] = goals.map(goal => {
      const course = planCourses.find(candidate => candidate.identity === goal.course_identity);
      const available = course?.sections ?? [];
      return {
        goal: course?.name ?? goal.course_identity,
        attempts: goal.preferences
          .map(preference => available.find(section => sectionIdentity(section) === preference.section_id))
          .filter((section): section is PlanSection => Boolean(section))
          .map(section => ({ sectionId: sectionIdentity(section), label: sectionLabel(section) })),
      };
    });
    const preview = steps.map((step, index) => `${index + 1}. ${step.goal}${step.attempts.length ? `：依次尝试 ${step.attempts.length} 个教学班` : "（无可执行教学班）"}`).join("\n");
    askConfirm("按优先级提交选课", `${steps.length} 个课程目标：\n\n${preview}\n\n每个教学班只提交一次、不会自动重试；某目标成功后自动跳到下一个目标。基于当前待选课程快照执行。`, undefined, () => void executeQueue(steps, snapshotId));
  };

  const executeQueue = async (steps: QueueStep[], snapshotId: string) => {
    if (!state) return;
    const results: QueueResult[] = [];
    setQueueResults([]);
    setQueueRunning(true);
    setMessage("");
    try {
      for (const step of steps) {
        if (!step.attempts.length) {
          results.push({ goal: step.goal, sectionId: "", sectionLabel: "", outcome: "skipped", message: "该课程没有可执行的教学班" });
          setQueueResults([...results]);
          continue;
        }
        for (const attempt of step.attempts) {
          const section = planSections.find(candidate => sectionIdentity(candidate) === attempt.sectionId);
          const blockedReason = section ? sectionExecutionBlocked(section) : "教学班已不在快照中";
          if (blockedReason) {
            results.push({ goal: step.goal, sectionId: attempt.sectionId, sectionLabel: attempt.label, outcome: "skipped", message: blockedReason });
            setQueueResults([...results]);
            continue;
          }
          const response = await fetch("/api/executions/selection", {
            method: "POST",
            headers: jsonHeaders(state.csrf_token),
            body: JSON.stringify({ section_id: attempt.sectionId, snapshot_id: snapshotId, confirmation: attempt.sectionId }),
          });
          const created = await response.json();
          if (!response.ok) {
            results.push({ goal: step.goal, sectionId: attempt.sectionId, sectionLabel: attempt.label, outcome: "failed", message: created.error ?? "提交失败" });
            setQueueResults([...results]);
            break;
          }
          setTask(created as Task);
          const finished = await waitForTask(created as Task);
          const succeeded = finished.state === "succeeded";
          results.push({ goal: step.goal, sectionId: attempt.sectionId, sectionLabel: attempt.label, outcome: succeeded ? "submitted" : "failed", message: finished.error ?? (succeeded ? "已提交" : `任务结束：${finished.state}`) });
          setQueueResults([...results]);
          if (succeeded) break;
        }
      }
      await load();
    } finally {
      setQueueRunning(false);
    }
  };

  if (!state) return <main className="shell"><p>正在读取本地工作台…</p></main>;
  if (!state.login_configuration.configured) return <main className="login-shell">
    <section className="login-intro">
      <p className="eyebrow">LOCAL ACADEMIC WORKBENCH</p>
      <h1>先连接你的教务身份</h1>
      <p>保存一次，之后工作台会在学校统一认证页面自动登录。成绩、课表和选课查询仍只在本机完成。</p>
      <dl><div><dt>存储位置</dt><dd>仅本机 .private 目录</dd></div><div><dt>加密方式</dt><dd>Windows DPAPI · 当前用户可解密</dd></div><div><dt>操作边界</dt><dd>查询只读；选课需逐次明确确认，不提供退课</dd></div></dl>
    </section>
    <section className="login-form-panel" aria-labelledby="login-title">
      <div><span>首次设置</span><h2 id="login-title">自动登录配置</h2><p>请输入学校统一身份认证使用的学号和密码。</p></div>
      {message && <p className="login-error" role="alert">{message}</p>}
      {state.login_configuration.state === "invalid" && <p className="login-error" role="alert">{state.login_configuration.message ?? "原登录配置无法读取，请重新保存。"}</p>}
      <form onSubmit={configureLogin} autoComplete="off">
        <label>学号<input name="academic-username" value={username} onChange={event => setUsername(event.target.value)} inputMode="numeric" autoComplete="off" required autoFocus /></label>
        <label>密码<input name="academic-password" type="password" value={password} onChange={event => setPassword(event.target.value)} autoComplete="new-password" required /></label>
        <button type="submit" disabled={savingLogin}>{savingLogin ? "正在安全保存…" : "保存登录配置"}</button>
      </form>
      <small>凭据不会写入数据库、日志或课程快照。学校要求验证码时，浏览器会停下等待你处理。</small>
    </section>
  </main>;
  const timetable = state.snapshots.timetable;
  const selection = state.snapshots.selection;
  const planSections = ((selection?.payload.sections as PlanSection[] | undefined) ?? []).filter(section => typeof section === "object" && section !== null);
  const planCourses: PlanCourse[] = (() => {
    const byCourse = new Map<string, PlanCourse>();
    for (const section of planSections) {
      const identity = String(section.course_code ?? section.course_name ?? section.name ?? "未命名课程");
      const existing = byCourse.get(identity);
      if (existing) existing.sections.push(section);
      else byCourse.set(identity, { identity, name: String(section.course_name ?? section.name ?? identity), category: String(section.category ?? "未分类"), sections: [section] });
    }
    return [...byCourse.values()].sort((a, b) => a.name.localeCompare(b.name, "zh-CN"));
  })();
  const currentScheduleItems = ((timetable?.payload.entries as Record<string, unknown>[] | undefined) ?? []).flatMap((item, index) => expandScheduleItems(item, "current", index));

  const previewStatusFor = (goal: PlanGoal): { text: string; kind: "ok" | "conflict" | "unknown" } => {
    const top = goal.preferences[0];
    if (!top) return { text: "未选教学班", kind: "unknown" };
    const section = planSections.find(candidate => sectionIdentity(candidate) === top.section_id);
    if (!section) return { text: "教学班已不在快照中", kind: "unknown" };
    const meetings = expandScheduleItems(section as Record<string, unknown>, "candidate", 0).filter(item => !item.unknown);
    if (!meetings.length) return { text: "时间待确认", kind: "unknown" };
    const conflict = meetings.some(meeting => currentScheduleItems.some(other => !other.unknown && timeOverlaps(meeting, other) && (!meeting.weeks.length || !other.weeks.length || meeting.weeks.some(week => other.weeks.includes(week)))));
    return conflict ? { text: "学期内有冲突", kind: "conflict" } : { text: "学期内无冲突", kind: "ok" };
  };
  const terminalStates = ["succeeded", "failed", "cancelled", "interface_unconfirmed"];
  const activeTask = state.active_task ?? (task && !terminalStates.includes(task.state) ? task : null);
  const remoteBusy = Boolean(activeTask);
  const formatSourceTime = (iso?: string): string => {
    if (!iso) return "无记录";
    const then = new Date(iso).getTime();
    if (Number.isNaN(then)) return iso;
    const seconds = Math.max(0, Math.floor((Date.now() - then) / 1000));
    if (seconds < 60) return "刚刚";
    if (seconds < 3600) return `${Math.floor(seconds / 60)} 分钟前`;
    if (seconds < 86400) return `${Math.floor(seconds / 3600)} 小时前`;
    if (seconds < 172800) return "昨天";
    if (seconds < 604800) return `${Math.floor(seconds / 86400)} 天前`;
    const date = new Date(then);
    const sameYear = date.getFullYear() === new Date().getFullYear();
    const month = String(date.getMonth() + 1).padStart(2, "0");
    const day = String(date.getDate()).padStart(2, "0");
    return sameYear ? `${month}-${day}` : `${date.getFullYear()}-${month}-${day}`;
  };
  const statusLabel = (kind: "selection" | "timetable" | "progress") => {
    const value = state.snapshot_status?.[kind];
    const labels: Record<string, string> = { current: "当前", historical: "历史", incomplete: "不完整", missing: "缺失" };
    const status = labels[value?.status ?? "missing"] ?? value?.status ?? "缺失";
    const time = value?.source_at ? ` · ${formatSourceTime(value.source_at)}` : "";
    const reason = value?.reason ? ` · ${value.reason}` : "";
    return `${status}${time}${reason}`;
  };
  const grade = String(state.profile?.grade ?? "");
  const selectionWindows = selectionWindowsForGrade(
    (state.confirmed_notice?.windows as SelectionWindow[] | undefined) ?? [],
    grade,
  );
  const nextWindow = selectionWindows.find(window => new Date(String(window.closes_at ?? "").replace(" ", "T")).getTime() > Date.now()) ?? selectionWindows[selectionWindows.length - 1];
  const nextWindowDisplay = nextWindow ? selectionWindowDisplay(nextWindow) : null;
  const nextWindowLabel = nextWindowDisplay ? `${nextWindowDisplay.date} ${nextWindowDisplay.time}` : "";
  const pendingNotices = candidates.filter(notice => notice.status !== "confirmed");
  const taskProgressLabel = (value?: Record<string, unknown>) => {
    if (!value) return "";
    if (value.semester) return `正在读取 ${value.semester}${value.page ? ` · 第 ${value.page}/${value.page_count ?? "?"} 页` : ""}${value.records !== undefined ? ` · 已读取 ${value.records} 条课程记录` : ""}`;
    return String(value.message ?? value.target ?? "");
  };

  return <main className="shell">
    <header className="app-header">
      <div className="header-copy">
        <p className="eyebrow">GUARDED ACADEMIC WORKBENCH</p>
        <h1>选课规划工作台</h1>
      </div>
      <aside className="window-summary" aria-labelledby="window-summary-title">
        <button className="window-summary-toggle" onClick={() => setWindowsOpen(value => !value)} aria-expanded={windowsOpen} aria-haspopup="true">
          <span className="window-summary-copy"><small id="window-summary-title">{grade ? `${grade} 级` : ""}选课开放</small><strong>{selectionWindows.length ? `${selectionWindows.length} 个时段${nextWindowLabel ? ` · 下一 ${nextWindowLabel}` : ""}` : grade ? "暂无适用时段" : "完善年级后显示"}</strong></span>
          <span className="window-summary-chevron" aria-hidden="true">{windowsOpen ? "▲" : "▼"}</span>
        </button>
        {windowsOpen && <div className="window-summary-panel">
          {pendingNotices.length > 0 && <div className="window-summary-notices">{pendingNotices.map(notice => <div className="window-summary-notice" key={notice.version_id}><span><strong>{notice.title || "未命名通知"}</strong><small>{notice.term || "学期待确认"}</small></span><button className="secondary" onClick={() => confirm(notice.version_id)} disabled={remoteBusy}>确认</button></div>)}</div>}
          {selectionWindows.length ? <ul className="window-summary-list">{selectionWindows.map((window, index) => {
            const display = selectionWindowDisplay(window);
            return <li key={`${window.opens_at}-${index}`} title={display.fullRange}><time dateTime={window.opens_at?.replace(" ", "T")}>{display.date}</time><b>{display.time}</b><span>{selectionCategoryLabel(window)}</span></li>;
          })}</ul> : <p className="selection-window-empty">{grade ? `已确认通知中暂无 ${grade} 级适用选课时段` : "完善学生年级后显示适用选课时段"}</p>}
          <div className="window-summary-foot"><button className="secondary" onClick={discoverOfficialNotices} disabled={remoteBusy}>更新通知</button></div>
        </div>}
      </aside>
    </header>
    {message && <p className="notice" role="status">{message}</p>}
    <ScheduleBoard timetable={timetable} selection={selection} graduationProgress={state.graduation_progress} selectionWindows={(state.confirmed_notice?.windows as SelectionWindow[] | undefined) ?? []} studentGrade={grade} executionHistory={(state.execution_history as Array<Record<string, unknown>> | undefined) ?? []} previewKeys={goals.flatMap(goal => goal.preferences.slice(0, 1).map(preference => preference.section_id))} onTogglePreview={toggleQueueSection} onExecuteSection={executeSection} executionPending={remoteBusy || queueRunning} />

    <button ref={drawerTriggerRef} className={`data-drawer-trigger ${remoteBusy ? "is-busy" : ""}`} onClick={() => setDataDrawerOpen(true)} aria-haspopup="dialog" aria-expanded={dataDrawerOpen}>
      <span className="data-trigger-mark" aria-hidden="true" />
      <span><strong>{remoteBusy ? "数据同步中" : "数据操作"}</strong><small>刷新与快照状态</small></span>
    </button>
    {dataDrawerOpen && <div className="data-drawer-layer">
      <button className="data-drawer-backdrop" onClick={() => setDataDrawerOpen(false)} aria-label="关闭数据操作面板" />
      <aside className="data-drawer" role="dialog" aria-modal="true" aria-labelledby="data-drawer-title">
        <header className="data-drawer-heading"><div><small>教务数据中心</small><h2 id="data-drawer-title">刷新与连接</h2><p>远程读取会保留最近一次完整快照，不会因失败清空现有数据。</p></div><button ref={drawerCloseRef} className="drawer-close" onClick={() => setDataDrawerOpen(false)} aria-label="关闭数据操作面板">关闭</button></header>
        <section className="drawer-actions" aria-label="数据操作">
          <button onClick={() => run("connect")} disabled={remoteBusy}><span>连接教务</span><small>验证当前教务会话</small></button>
          <button onClick={() => run("refresh-timetable")} disabled={remoteBusy}><span>刷新课表</span><small>{statusLabel("timetable")}</small></button>
          <button onClick={() => run("refresh-progress")} disabled={remoteBusy}><span>刷新毕业进度</span><small>{statusLabel("progress")}</small></button>
          <button onClick={() => run("refresh-selection")} disabled={remoteBusy || !state.confirmed_notice}><span>刷新待选课程</span><small>{state.confirmed_notice ? statusLabel("selection") : "需先确认官方通知"}</small></button>
          {state.capabilities?.development_diagnostics && <button className="secondary" onClick={() => run("observe-navigation")} disabled={remoteBusy}><span>诊断监听</span><small>开发诊断工具</small></button>}
        </section>
        {activeTask && <section className="drawer-task" aria-live="polite"><small>当前任务</small><strong>{activeTask.operation} · {activeTask.state}</strong><span>{taskProgressLabel(activeTask.progress) || `开始于 ${activeTask.created_at ?? "—"}`}</span><div>{activeTask.operation === "observe-navigation" && <button onClick={finishObservation}>完成监听</button>}{activeTask.task_kind !== "execution" && <button className="danger" onClick={cancel}>取消任务</button>}</div></section>}
        {task && !activeTask && <section className="drawer-task" aria-live="polite"><small>最近任务</small><strong>{task.operation} · {task.state}</strong>{task.error && <span>{task.error}{task.task_kind === "observation" && task.operation !== "connect" && "；旧快照仍然保留"}</span>}{taskProgressLabel(task.progress) && <span>{taskProgressLabel(task.progress)}</span>}</section>}
        <section className="drawer-snapshots" aria-label="快照状态"><article title={state.snapshot_status?.timetable?.source_at}><small>课表</small><strong>{statusLabel("timetable")}</strong></article><article title={state.snapshot_status?.selection?.source_at}><small>待选课程</small><strong>{statusLabel("selection")}</strong></article><article title={state.snapshot_status?.progress?.source_at}><small>毕业进度</small><strong>{statusLabel("progress")}</strong></article></section>
        <section className="drawer-session" aria-label="教务会话"><small>教务会话</small><strong>{state.login_configuration.masked_username} · {state.academic_session.state}</strong><span>浏览器 {state.academic_session.browser ?? "未知"} · WebVPN {state.academic_session.webvpn ?? "未知"}{state.academic_session.last_verified_at && ` · 验证于 ${state.academic_session.last_verified_at}`}</span><div><button onClick={clearLogin}>重新配置登录</button></div></section>
        {state.execution_history && state.execution_history.length > 0 && <section className="drawer-history" aria-label="选课执行记录"><small>选课执行记录</small>{state.execution_history.map(item => <div className="drawer-history-item" key={item.id}><strong>{item.course_name || item.section_id}</strong><span>{item.created_at} · {item.result}{item.message ? ` · ${item.message}` : ""}</span>{item.result === "unknown" && !item.resolved && <button onClick={() => resolveUnknown(item.id)}>已核实，解除阻断</button>}</div>)}<button className="danger" onClick={clearHistory}>清除执行记录</button></section>}
      </aside>
    </div>}
    {pendingConfirm && <ConfirmDialog pending={pendingConfirm} onConfirm={() => { const action = pendingConfirm.onConfirm; setPendingConfirm(null); action(); }} onCancel={() => setPendingConfirm(null)} />}

    <section className="queue-panel" aria-labelledby="queue-title">
      <header className="queue-heading">
        <div><small>课表预览 · 选课执行</small><h2 id="queue-title">预览与抢课队列</h2><p>{goals.length ? `正在预览 ${goals.length} 门课程：拖拽排序，第 1 位最先提交。` : "加入队列的课程会叠加到课表上比较时间；这里安排抢课顺序。"}</p></div>
        <div className="queue-actions">{goals.length > 0 && <button className="text-button" onClick={() => { setGoals([]); setQueueResults([]); }}>清空队列</button>}<button className="run-queue" onClick={runQueue} disabled={!goals.length || queueRunning || remoteBusy}>{queueRunning ? "正在按队列提交…" : "按队列抢课"}</button></div>
      </header>
      {!planSections.length && <p className="queue-empty">暂无待选课程快照：先在右下角“数据操作”中刷新待选课程，再建立队列。</p>}
      {goals.map((goal, goalIndex) => {
        const available = planCourses.find(course => course.identity === goal.course_identity)?.sections ?? [];
        const status = previewStatusFor(goal);
        return <article className={`plan-goal ${dragGoalIndex === goalIndex ? "dragging" : ""}`} key={goal.goal_id}
          onDragOver={event => { event.preventDefault(); event.dataTransfer.dropEffect = "move"; }}
          onDrop={event => { event.preventDefault(); const from = Number(event.dataTransfer.getData("text/plain")); if (Number.isFinite(from)) moveGoalTo(from, goalIndex); setDragGoalIndex(null); }}>
          <header><span className="drag-handle" draggable onDragStart={event => { setDragGoalIndex(goalIndex); event.dataTransfer.effectAllowed = "move"; event.dataTransfer.setData("text/plain", String(goalIndex)); }} onDragEnd={() => setDragGoalIndex(null)} aria-hidden="true">⋮⋮</span><span className="rank-badge">{goalIndex + 1}</span>
            <select value={goal.course_identity} onChange={event => changeGoalCourse(goalIndex, event.target.value)} aria-label={`队列第 ${goalIndex + 1} 位的课程`}>{planCourses.map(course => <option key={course.identity} value={course.identity}>{course.name}{course.category ? `（${course.category}）` : ""}</option>)}</select>
            <div className="goal-controls"><button className="secondary" disabled={goalIndex === 0} onClick={() => moveGoal(goalIndex, -1)} aria-label="上移队列优先级">↑</button><button className="secondary" disabled={goalIndex >= goals.length - 1} onClick={() => moveGoal(goalIndex, 1)} aria-label="下移队列优先级">↓</button><button className="danger" onClick={() => removeGoal(goalIndex)} aria-label="移出队列">移出</button></div>
          </header>
          <p className={`preview-status ${status.kind}`}><b>预览：{status.text}</b><span>首位教学班叠加在课表上</span></p>
          {goal.preferences.map((preference, preferenceIndex) => {
            const section = available.find(candidate => sectionIdentity(candidate) === preference.section_id);
            return <div className="plan-pref" key={preference.section_id}><span className="rank-badge small">{preferenceIndex + 1}</span><span className="plan-pref-copy">{section ? sectionLabel(section) : `教学班 ${preference.section_id}`}</span>{preferenceIndex === 0 && <span className="preview-tag">预览</span>}<div className="goal-controls"><button className="secondary" disabled={preferenceIndex === 0} onClick={() => movePreference(goalIndex, preferenceIndex, -1)} aria-label="上移教学班偏好">↑</button><button className="secondary" disabled={preferenceIndex >= goal.preferences.length - 1} onClick={() => movePreference(goalIndex, preferenceIndex, 1)} aria-label="下移教学班偏好">↓</button><button className="danger" onClick={() => removePreference(goalIndex, preferenceIndex)} aria-label="移除教学班偏好">移除</button></div></div>;
          })}
          {available.length > goal.preferences.length && <div className="plan-add-pref"><select value="" onChange={event => addPreference(goalIndex, event.target.value)} aria-label={`为队列第 ${goalIndex + 1} 位添加教学班`}><option value="">添加教学班偏好…</option>{available.filter(section => !goal.preferences.some(preference => preference.section_id === sectionIdentity(section))).map(section => <option key={sectionIdentity(section)} value={sectionIdentity(section)}>{sectionLabel(section)}</option>)}</select></div>}
        </article>;
      })}
      {planCourses.length > 0 && <div className="plan-add-goal"><select value={pendingCourse} onChange={event => setPendingCourse(event.target.value)} aria-label="选择要加入队列的课程"><option value="">选择要加入队列的课程…</option>{planCourses.map(course => <option key={course.identity} value={course.identity}>{course.name}{course.category ? `（${course.category}）` : ""}</option>)}</select><button onClick={addGoal} disabled={!pendingCourse}>加入队列</button></div>}
      <div className="plan-save-row"><button onClick={savePlan} disabled={!goals.length}>保存队列</button><small>列表顺序即抢课优先级：第 1 位最先提交。</small></div>
      {queueResults.length > 0 && <ol className="queue-results" aria-live="polite">{queueResults.map((result, index) => <li key={index} className={`queue-result ${result.outcome}`}><b>{result.outcome === "submitted" ? "已提交" : result.outcome === "skipped" ? "跳过" : "失败"}</b><span>{result.goal}{result.sectionLabel ? ` · ${result.sectionLabel}` : ""}</span><small>{result.message}</small></li>)}</ol>}
      {plan && <div className="plan-result-view">
        <div className={`plan-result-status ${plan.status === "ready" ? "ready" : "blocked"}`}><strong>{plan.status === "ready" ? "队列可用" : "队列被数据条件阻断"}</strong><span>{plan.status === "ready" ? "按上方顺序提交；冲突教学班会被自动跳过。" : "请先处理以下条件后重新保存："}</span></div>
        {Array.isArray(plan.blocked_reasons) && plan.blocked_reasons.length > 0 && <ul className="plan-reasons">{plan.blocked_reasons.map(reason => <li key={String(reason)}>{planReasonLabel(String(reason))}</li>)}</ul>}
        {Array.isArray(plan.conflicts) && plan.conflicts.length > 0 && <ul className="plan-conflicts">{plan.conflicts.map((conflict, index) => <li key={index}>{planConflictLabel(conflict as Record<string, unknown>)}</li>)}</ul>}
        <details><summary>查看队列数据</summary><pre className="plan-result">{JSON.stringify(plan, null, 2)}</pre></details>
      </div>}
    </section>
  </main>;
}

createRoot(document.getElementById("root")!).render(<App />);
