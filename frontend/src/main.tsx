import React, { useCallback, useEffect, useState } from "react";
import { createRoot } from "react-dom/client";
import type { CandidateNotice, Task, WorkbenchState } from "./api.generated";
import { ScheduleBoard } from "./ScheduleBoard";
import "./style.css";

const jsonHeaders = (token: string) => ({ "Content-Type": "application/json", "X-CSRF-Token": token });

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
    const [response, notices] = await Promise.all([
      fetch("/api/state"),
      fetch("/api/notices/candidates"),
    ]);
    const [next, noticeResult] = await Promise.all([
      response.json() as Promise<WorkbenchState>,
      notices.json() as Promise<{ notices?: CandidateNotice[] }>,
    ]);
    setState(next);
    setPlan(next.latest_plan);
    setCandidates(noticeResult.notices ?? []);
  }, []);

  useEffect(() => { load().catch((error: unknown) => setMessage(String(error))); }, [load]);

  const run = async (operation: string) => {
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

  const inspectNotice = async () => {
    if (!state) return;
    const response = await fetch("/api/notices/candidates", {
      method: "POST",
      headers: jsonHeaders(state.csrf_token),
      body: JSON.stringify({ source_url: url, text }),
    });
    const result = await response.json();
    setMessage(response.ok ? "已发现候选通知，请确认后使用" : (result.error ?? "通知检查失败"));
    if (response.ok) {
      setUrl("");
      setText("");
      await load();
    }
  };

  const confirm = async (id: string) => {
    if (!state) return;
    const response = await fetch(`/api/notices/${id}/confirm`, {
      method: "POST",
      headers: jsonHeaders(state.csrf_token),
    });
    const result = await response.json();
    setMessage(response.ok ? "通知已确认" : (result.error ?? "确认失败"));
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

  const savePlan = async () => {
    if (!state) return;
    let parsed: unknown;
    try {
      parsed = JSON.parse(goals);
    } catch {
      setMessage("规划必须是有效的 JSON 数组");
      return;
    }
    if (!Array.isArray(parsed)) {
      setMessage("规划必须是目标数组");
      return;
    }
    const response = await fetch("/api/plans", {
      method: "POST",
      headers: jsonHeaders(state.csrf_token),
      body: JSON.stringify({ goals: parsed }),
    });
    const result = await response.json();
    if (!response.ok) {
      setMessage(result.error ?? "规划保存失败");
      return;
    }
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
    <section className="panel"><h2>主动检查官方选课通知</h2><p>仅解析批准的官方来源，不会执行选课操作。</p><input value={url} onChange={event => setUrl(event.target.value)} placeholder="https://jwc.hitwh.edu.cn/…" /><textarea value={text} onChange={event => setText(event.target.value)} placeholder="也可以粘贴通知正文" rows={4} /><button onClick={inspectNotice}>检查并生成候选</button></section>
    <section className="panel"><h2>候选通知</h2>{candidates.length ? candidates.map(notice => <article className="candidate" key={notice.version_id}><strong>{notice.title || "未命名通知"}</strong><span>{notice.term || "待补充"} · {notice.status}</span>{notice.status !== "confirmed" && <button onClick={() => confirm(notice.version_id)}>确认此通知</button>}</article>) : <p className="empty">暂无候选通知</p>}</section>
    <section className="panel"><h2>本地只读规划</h2><p>按课程目标优先级和教学班偏好填写 JSON；这里只保存本地规划，不会发送选课请求。</p><textarea value={goals} onChange={event => setGoals(event.target.value)} rows={7} /><button onClick={savePlan}>保存规划</button>{plan && <pre className="plan-result">{JSON.stringify(plan, null, 2)}</pre>}</section>
    <ScheduleBoard timetable={timetable} selection={selection} graduationProgress={state.graduation_progress} />
  </main>;
}

createRoot(document.getElementById("root")!).render(<App />);
