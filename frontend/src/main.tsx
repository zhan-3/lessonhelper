import React, { useCallback, useEffect, useMemo, useState } from "react";
import { createRoot } from "react-dom/client";
import type { CandidateNotice, Task, WorkbenchState } from "./api.generated";
import "./style.css";

const jsonHeaders = (token: string) => ({ "Content-Type": "application/json", "X-CSRF-Token": token });

type ScheduleItem = Record<string, unknown> & { source: "current" | "candidate"; key: string; day: number | null; start: number | null; end: number | null; weeks: number[]; parity: string; unknown: boolean; name: string };
const number = (value: unknown): number | null => { const match = String(value ?? "").match(/\d+/); return match ? Number(match[0]) : null; };
const numbers = (value: unknown): number[] => String(value ?? "").split(/[,\s]+/).map(Number).filter(Number.isFinite);
function scheduleItem(raw: Record<string, unknown>, source: ScheduleItem["source"], index: number): ScheduleItem {
  const text = String(raw.time ?? raw.schedule ?? "");
  const range = text.match(/(\d+)\s*[-~至]\s*(\d+)/);
  const day = number(raw.weekday ?? raw.day ?? (text.match(/(?:星期|周)\s*([1-7])/)?.[1]));
  const start = number(raw.start_period ?? raw.start ?? range?.[1]);
  const end = number(raw.end_period ?? raw.end ?? range?.[2] ?? start);
  const weeks = numbers(raw.week_numbers ?? raw.weeks);
  const parity = String(raw.week_parity ?? raw.parity ?? "all").toLowerCase();
  const unknown = raw.conflict_status === "unknown" || !day || !start || !end || /待定|未定|pending|tbd/i.test(text);
  return { ...raw, source, key: `${source}-${index}`, day: day && day >= 1 && day <= 7 ? day : null, start, end, weeks, parity, unknown, name: String(raw.course_name ?? raw.name ?? raw.title ?? "未命名课程") };
}
function ScheduleBoard({ timetable, selection }: { timetable: WorkbenchState["snapshots"]["timetable"]; selection: WorkbenchState["snapshots"]["selection"] }) {
  const [mobileDay, setMobileDay] = useState(1);
  const [mobileWeek, setMobileWeek] = useState(() => window.matchMedia("(min-width: 701px)").matches);
  const current = useMemo(() => ((timetable?.payload.entries as Record<string, unknown>[] | undefined) ?? []).map((item, i) => scheduleItem(item, "current", i)), [timetable]);
  const candidates = useMemo(() => ((selection?.payload.sections as Record<string, unknown>[] | undefined) ?? []).map((item, i) => scheduleItem(item, "candidate", i)), [selection]);
  const all = [...current, ...candidates];
  const overlaps = (a: ScheduleItem, b: ScheduleItem) => {
    const weeksOverlap = !a.weeks.length || !b.weeks.length || a.weeks.some(week => b.weeks.includes(week));
    const parityOverlap = a.parity === "all" || b.parity === "all" || a.parity === b.parity;
    return a.day !== null && a.day === b.day && a.start !== null && b.start !== null && a.end !== null && b.end !== null && a.start <= b.end && b.start <= a.end && weeksOverlap && parityOverlap;
  };
  const currentConflict = new Set(current.filter((item, i) => current.some((other, j) => i !== j && overlaps(item, other)) || candidates.some(other => overlaps(item, other))).map(item => item.key));
  const candidateConflict = new Set(candidates.filter((item, i) => candidates.some((other, j) => i !== j && overlaps(item, other))).map(item => item.key));
  const candidateCurrentConflict = new Set(candidates.filter(item => current.some(other => overlaps(item, other))).map(item => item.key));
  const days = [1, 2, 3, 4, 5, 6, 7];
  const visibleDays = mobileWeek ? days : [mobileDay];
  const renderItem = (item: ScheduleItem) => { const currentConflictClass = item.source === "current" ? currentConflict.has(item.key) : candidateCurrentConflict.has(item.key); const candidateConflictClass = item.source === "candidate" && candidateConflict.has(item.key); return <article className={`schedule-card ${item.source} ${currentConflictClass ? "current-conflict" : ""} ${candidateConflictClass ? "candidate-conflict" : ""} ${item.unknown ? "unknown" : ""}`} key={item.key}><strong>{item.name}</strong><span>{item.start && item.end ? `${item.start}-${item.end} 节` : "时间待定"}</span>{item.unknown && <em>冲突未知</em>}</article>; };
  return <section className="schedule-panel"><div className="schedule-heading"><div><h2>七日课表与候选叠加</h2><p className="muted">灰色为待选课程，红色为当前课表冲突，橙色为候选课程间冲突。</p></div><div className="schedule-switch"><button className={!mobileWeek ? "secondary" : ""} onClick={() => setMobileWeek(false)}>日视图</button><button className={mobileWeek ? "secondary" : ""} onClick={() => setMobileWeek(true)}>整周</button></div></div>{all.some(item => item.unknown) && <p className="unknown-hint">存在时间待定课程：系统不会把它当作空闲，请确认时间后再规划。</p>}<nav className="day-tabs" aria-label="选择日期">{days.map(day => <button key={day} className={mobileDay === day ? "active" : ""} onClick={() => setMobileDay(day)}>星期{day}</button>)}</nav><div className={`week schedule-grid ${mobileWeek ? "show-week" : "single-day"}`}>{visibleDays.map(day => <div className="day-column" key={day}><h3>星期{day}</h3>{all.filter(item => item.day === day).map(renderItem)}{!all.some(item => item.day === day) && <p className="empty-day">暂无课程</p>}</div>)}</div>{all.some(item => item.day === null) && <div className="unknown-list"><strong>未排入日期</strong>{all.filter(item => item.day === null).map(renderItem)}</div>}</section>;
}

function App() {
  const [state, setState] = useState<WorkbenchState | null>(null);
  const [candidates, setCandidates] = useState<CandidateNotice[]>([]);
  const [task, setTask] = useState<Task | null>(null);
  const [url, setUrl] = useState("");
  const [text, setText] = useState("");
  const [goals, setGoals] = useState('[{"goal_id":"goal-1","course_identity":"COURSE","rank":1,"preferences":[{"section_id":"SECTION","rank":1}]}]');
  const [plan, setPlan] = useState<Record<string, unknown> | null>(null);
  const [message, setMessage] = useState("");

  const load = useCallback(async () => {
    const response = await fetch("/api/state");
    const next = await response.json() as WorkbenchState;
    setState(next);
    setPlan(next.latest_plan);
    const notices = await fetch("/api/notices/candidates");
    setCandidates((await notices.json()).notices ?? []);
  }, []);
  useEffect(() => { load().catch((error: unknown) => setMessage(String(error))); }, [load]);

  const run = async (operation: string) => {
    if (!state) return;
    const response = await fetch("/api/tasks", { method: "POST", headers: jsonHeaders(state.csrf_token), body: JSON.stringify({ operation }) });
    const created = await response.json();
    if (!response.ok) { setMessage(created.error ?? "任务提交失败"); return; }
    setTask(created);
  };
  useEffect(() => {
    if (!task || !state || ["succeeded", "failed", "cancelled"].includes(task.state)) return;
    const timer = window.setInterval(async () => {
      const response = await fetch(`/api/tasks/${task.id}`);
      if (response.ok) {
        const next = await response.json() as Task;
        setTask(next);
        if (["succeeded", "failed", "cancelled"].includes(next.state)) { await load(); }
      }
    }, 600);
    return () => window.clearInterval(timer);
  }, [task, state, load]);

  const inspectNotice = async () => {
    if (!state) return;
    const response = await fetch("/api/notices/candidates", { method: "POST", headers: jsonHeaders(state.csrf_token), body: JSON.stringify({ source_url: url, text }) });
    const result = await response.json();
    setMessage(response.ok ? "已发现候选通知，请确认后使用" : (result.error ?? "通知检查失败"));
    if (response.ok) { setUrl(""); setText(""); await load(); }
  };
  const confirm = async (id: string) => {
    if (!state) return;
    const response = await fetch(`/api/notices/${id}/confirm`, { method: "POST", headers: jsonHeaders(state.csrf_token) });
    const result = await response.json();
    setMessage(response.ok ? "通知已确认" : (result.error ?? "确认失败"));
    await load();
  };
  const cancel = async () => {
    if (!task || !state) return;
    await fetch(`/api/tasks/${task.id}`, { method: "DELETE", headers: { "X-CSRF-Token": state.csrf_token } });
  };
  const finishObservation = async () => {
    if (!task || !state) return;
    const response = await fetch(`/api/tasks/${task.id}/finish`, { method: "POST", headers: jsonHeaders(state.csrf_token) });
    if (!response.ok) setMessage("当前任务无法完成监听");
  };
  const savePlan = async () => {
    if (!state) return;
    let parsed: unknown;
    try { parsed = JSON.parse(goals); } catch { setMessage("规划必须是有效的 JSON 数组"); return; }
    if (!Array.isArray(parsed)) { setMessage("规划必须是目标数组"); return; }
    const response = await fetch("/api/plans", { method: "POST", headers: jsonHeaders(state.csrf_token), body: JSON.stringify({ goals: parsed }) });
    const result = await response.json();
    if (!response.ok) { setMessage(result.error ?? "规划保存失败"); return; }
    setPlan(result);
    setMessage(result.status === "ready" ? "只读规划已保存" : "规划已保存，但当前仍被数据条件阻断");
  };
  if (!state) return <main className="shell"><p>正在读取本地工作台…</p></main>;
  const timetable = state.snapshots.timetable;
  const selection = state.snapshots.selection;
  return <main className="shell">
    <header><p className="eyebrow">READ-ONLY ACADEMIC WORKBENCH</p><h1>选课规划工作台</h1><span className="session">教务会话 · {state.academic_session.state}</span></header>
    {message && <p className="notice">{message}</p>}
    <section className="toolbar"><button onClick={() => run("connect")}>连接教务会话</button><button onClick={() => run("observe-navigation")}>开始手动监听</button><button onClick={() => run("refresh-selection")} disabled={!state.confirmed_notice}>刷新选课班</button><button onClick={() => run("refresh-timetable")}>刷新课表</button><button className="secondary" onClick={load}>刷新状态</button></section>
    {(state.stale.selection || state.stale.timetable) && <p className="warning">存在过期快照：刷新失败时仍保留旧数据，规划不会把过期数据标记为已就绪。</p>}
    {task && <section className="task"><strong>任务：{task.state}</strong>{task.progress && <span> · {JSON.stringify(task.progress)}</span>}{task.operation === "observe-navigation" && !["succeeded", "failed", "cancelled"].includes(task.state) && <button onClick={finishObservation}>完成监听</button>}{!["succeeded", "failed", "cancelled"].includes(task.state) && <button className="danger" onClick={cancel}>取消</button>}</section>}
    <section className="facts"><article><small>当前画像</small><strong>{String(state.profile?.grade ?? "尚未配置")} 年级</strong></article><article><small>已确认通知</small><strong>{String(state.confirmed_notice?.title ?? "尚未确认")}</strong></article><article><small>当前学期</small><strong>{timetable?.term || selection?.term || "—"}</strong></article></section>
    <section className="panel"><h2>主动检查官方选课通知</h2><p>仅解析批准的官方来源，不会执行选课操作。</p><input value={url} onChange={(e: React.ChangeEvent<HTMLInputElement>) => setUrl(e.target.value)} placeholder="https://jwc.hitwh.edu.cn/…" /><textarea value={text} onChange={(e: React.ChangeEvent<HTMLTextAreaElement>) => setText(e.target.value)} placeholder="也可以粘贴通知正文" rows={4} /><button onClick={inspectNotice}>检查并生成候选</button></section>
    <section className="panel"><h2>候选通知</h2>{candidates.length ? candidates.map(n => <article className="candidate" key={n.version_id}><strong>{n.title || "未命名通知"}</strong><span>{n.term || "待补充"} · {n.status}</span>{n.status !== "confirmed" && <button onClick={() => confirm(n.version_id)}>确认此通知</button>}</article>) : <p className="empty">暂无候选通知</p>}</section>
    <section className="panel"><h2>本地只读规划</h2><p>按课程目标优先级和教学班偏好填写 JSON；这里只保存本地规划，不会发送选课请求。</p><textarea value={goals} onChange={(e: React.ChangeEvent<HTMLTextAreaElement>) => setGoals(e.target.value)} rows={7} /><button onClick={savePlan}>保存规划</button>{plan && <pre className="plan-result">{JSON.stringify(plan, null, 2)}</pre>}</section>
    <ScheduleBoard timetable={timetable} selection={selection} />
    <section><h2>待选课程</h2><p>{selection ? `${(selection.payload.sections as any[])?.length ?? 0} 个课程班 · ${selection.source_at}` : "暂无待选课程快照"}</p></section>
  </main>;
}
createRoot(document.getElementById("root")!).render(<App />);
