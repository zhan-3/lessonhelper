import React, { useCallback, useEffect, useState } from "react";
import { createRoot } from "react-dom/client";
import type { CandidateNotice, Task, WorkbenchState } from "./api.generated";
import "./style.css";

const jsonHeaders = (token: string) => ({ "Content-Type": "application/json", "X-CSRF-Token": token });

function App() {
  const [state, setState] = useState<WorkbenchState | null>(null);
  const [candidates, setCandidates] = useState<CandidateNotice[]>([]);
  const [task, setTask] = useState<Task | null>(null);
  const [url, setUrl] = useState("");
  const [text, setText] = useState("");
  const [message, setMessage] = useState("");

  const load = useCallback(async () => {
    const response = await fetch("/api/state");
    const next = await response.json() as WorkbenchState;
    setState(next);
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
  if (!state) return <main className="shell"><p>正在读取本地工作台…</p></main>;
  const timetable = state.snapshots.timetable;
  const selection = state.snapshots.selection;
  return <main className="shell">
    <header><p className="eyebrow">READ-ONLY ACADEMIC WORKBENCH</p><h1>选课规划工作台</h1><span className="session">教务会话 · {state.academic_session.state}</span></header>
    {message && <p className="notice">{message}</p>}
    <section className="toolbar"><button onClick={() => run("connect")}>连接教务会话</button><button onClick={() => run("refresh-selection")} disabled={!state.confirmed_notice}>刷新选课班</button><button onClick={() => run("refresh-timetable")}>刷新课表</button><button className="secondary" onClick={load}>刷新状态</button></section>
    {task && <section className="task"><strong>任务：{task.state}</strong>{task.progress && <span> · {JSON.stringify(task.progress)}</span>}{!["succeeded", "failed", "cancelled"].includes(task.state) && <button className="danger" onClick={cancel}>取消</button>}</section>}
    <section className="facts"><article><small>当前画像</small><strong>{String(state.profile?.grade ?? "尚未配置")} 年级</strong></article><article><small>已确认通知</small><strong>{String(state.confirmed_notice?.title ?? "尚未确认")}</strong></article><article><small>当前学期</small><strong>{timetable?.term || selection?.term || "—"}</strong></article></section>
    <section className="panel"><h2>主动检查官方选课通知</h2><p>仅解析批准的官方来源，不会执行选课操作。</p><input value={url} onChange={(e: React.ChangeEvent<HTMLInputElement>) => setUrl(e.target.value)} placeholder="https://jwc.hitwh.edu.cn/…" /><textarea value={text} onChange={(e: React.ChangeEvent<HTMLTextAreaElement>) => setText(e.target.value)} placeholder="也可以粘贴通知正文" rows={4} /><button onClick={inspectNotice}>检查并生成候选</button></section>
    <section className="panel"><h2>候选通知</h2>{candidates.length ? candidates.map(n => <article className="candidate" key={n.version_id}><strong>{n.title || "未命名通知"}</strong><span>{n.term || "待补充"} · {n.status}</span>{n.status !== "confirmed" && <button onClick={() => confirm(n.version_id)}>确认此通知</button>}</article>) : <p className="empty">暂无候选通知</p>}</section>
    <section><h2>每周课表</h2>{timetable ? <div className="week">{[1,2,3,4,5,6,7].map(day => <div key={day}><b>星期{day}</b>{((timetable.payload.entries as any[]) || []).filter(x => x.weekday === day).map((x, i) => <p key={i}>{x.start_period}-{x.end_period} {x.course_name}</p>)}</div>)}</div> : <p className="empty">暂无完整课表快照</p>}</section>
    <section><h2>待选课程</h2><p>{selection ? `${(selection.payload.sections as any[])?.length ?? 0} 个课程班 · ${selection.source_at}` : "暂无待选课程快照"}</p></section>
  </main>;
}
createRoot(document.getElementById("root")!).render(<App />);
