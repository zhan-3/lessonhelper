import React, { useCallback, useEffect, useState } from "react";
import { createRoot } from "react-dom/client";
import type { CandidateNotice, Task, WorkbenchState } from "./api";
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
  const [username, setUsername] = useState("");
  const [password, setPassword] = useState("");
  const [savingLogin, setSavingLogin] = useState(false);

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
    setCandidates(noticeResult.notices ?? []);
  }, []);

  useEffect(() => { load().catch((error: unknown) => setMessage(String(error))); }, [load]);

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

  const clearLogin = async () => {
    if (!state || !window.confirm("清除自动登录并重置当前学生的画像、课程快照和规划？官方通知会保留。")) return;
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

  const run = async (operation: string) => {
    if (!state) return;
    if (operation === "refresh-selection") {
      const status = state.snapshot_status?.selection;
      const windows = (state.confirmed_notice?.windows as Array<{ category_codes?: string[] }> | undefined) ?? [];
      const categories = windows.flatMap(window => window.category_codes ?? []);
      if (!window.confirm(`强制刷新待选课程？\n上次完整刷新：${status?.source_at || "无"}\n查询类别：${new Set(categories).size}\n\n这会访问学校系统。`)) return;
    }
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

  const executeSection = async (sectionId: string, courseName: string) => {
    if (!state?.snapshots.selection) return;
    if (!window.confirm(`确认提交一次选课请求？\n${courseName}\n教学班：${sectionId}\n\n系统不会自动重试。`)) return;
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

  const clearHistory = async () => {
    if (!state || !window.confirm("清除全部本地选课执行历史？")) return;
    await fetch("/api/executions", { method: "DELETE", headers: jsonHeaders(state.csrf_token) });
    await load();
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
        <button type="submit" disabled={savingLogin}>{savingLogin ? "正在安全保存…" : "保存并连接学校"}</button>
      </form>
      <small>凭据不会写入数据库、日志或课程快照。学校要求验证码时，浏览器会停下等待你处理。</small>
    </section>
  </main>;
  const timetable = state.snapshots.timetable;
  const selection = state.snapshots.selection;
  const terminalStates = ["succeeded", "failed", "cancelled", "interface_unconfirmed"];
  const activeTask = state.active_task ?? (task && !terminalStates.includes(task.state) ? task : null);
  const remoteBusy = Boolean(activeTask);
  const statusLabel = (kind: "selection" | "timetable" | "progress") => {
    const value = state.snapshot_status?.[kind];
    const labels: Record<string, string> = { current: "当前", historical: "历史", incomplete: "不完整", missing: "缺失" };
    return `${labels[value?.status ?? "missing"] ?? value?.status}${value?.source_at ? ` · ${value.source_at}` : ""}${value?.reason ? ` · ${value.reason}` : ""}`;
  };

  return <main className="shell">
    <header><p className="eyebrow">GUARDED ACADEMIC WORKBENCH</p><h1>选课规划工作台</h1><span className="session">{state.login_configuration.masked_username} · 浏览器 {state.academic_session.browser ?? "未知"} · WebVPN {state.academic_session.webvpn ?? "未知"} · 教务 {state.academic_session.state}{state.academic_session.last_verified_at && ` · 验证于 ${state.academic_session.last_verified_at}`} · <button className="text-button" onClick={clearLogin}>重新配置</button></span></header>
    {message && <p className="notice">{message}</p>}
    <section className="toolbar"><button onClick={() => run("connect")} disabled={remoteBusy}>连接教务</button><button onClick={() => run("refresh-timetable")} disabled={remoteBusy}>刷新课表</button><button onClick={() => run("refresh-progress")} disabled={remoteBusy}>刷新毕业进度</button><button onClick={() => run("refresh-selection")} disabled={remoteBusy || !state.confirmed_notice}>强制刷新待选课程</button><button onClick={() => run("observe-navigation")} disabled={remoteBusy}>诊断监听</button><button className="secondary" onClick={load}>刷新本地页面</button></section>
    {activeTask && <section className="task"><strong>当前任务：{activeTask.operation} · {activeTask.state}</strong><span> · 开始于 {activeTask.created_at ?? "—"} · 超时 {activeTask.timeout_seconds ?? "—"} 秒</span>{activeTask.progress && <span> · {JSON.stringify(activeTask.progress)}</span>}{activeTask.operation === "observe-navigation" && <button onClick={finishObservation}>完成监听</button>}{activeTask.task_kind !== "execution" && <button className="danger" onClick={cancel}>取消</button>}</section>}
    {task && !activeTask && <section className="task"><strong>最近任务：{task.operation} · {task.state}</strong>{task.error && <span> · {task.error}</span>}{task.progress && <span> · {JSON.stringify(task.progress)}</span>}</section>}
    <section className="facts"><article><small>课表快照</small><strong>{statusLabel("timetable")}</strong></article><article><small>待选课程快照</small><strong>{statusLabel("selection")}</strong></article><article><small>毕业进度快照</small><strong>{statusLabel("progress")}</strong></article></section>
    <section className="panel"><h2>主动检查官方选课通知</h2><p>直接读取教务处教务通知列表，下载标题类似“关于2026年秋季学期各类课程选课时间安排的通知”的公开静态页面并解析，不启动教务浏览器。</p><button onClick={discoverOfficialNotices}>从教务处获取最新选课安排</button><details><summary>手工链接或正文退路</summary><input value={url} onChange={event => setUrl(event.target.value)} placeholder="https://jwc.hitwh.edu.cn/…" /><textarea value={text} onChange={event => setText(event.target.value)} placeholder="也可以粘贴通知正文" rows={4} /><button onClick={inspectNotice}>检查并生成候选</button></details></section>
    <section className="panel"><h2>候选通知</h2>{candidates.length ? candidates.map(notice => <article className="candidate" key={notice.version_id}><strong>{notice.title || "未命名通知"}</strong><span>{notice.term || "待补充"} · {notice.status}</span>{notice.status !== "confirmed" && <button onClick={() => confirm(notice.version_id)}>确认此通知</button>}</article>) : <p className="empty">暂无候选通知</p>}</section>
    <section className="panel"><h2>本地只读规划</h2><p>按课程目标优先级和教学班偏好填写 JSON；这里只保存本地规划，不会发送选课请求。</p><textarea value={goals} onChange={event => setGoals(event.target.value)} rows={7} /><button onClick={savePlan}>保存规划</button>{plan && <pre className="plan-result">{JSON.stringify(plan, null, 2)}</pre>}</section>
    {state.execution_history && state.execution_history.length > 0 && <section className="panel"><h2>选课执行历史</h2>{state.execution_history.map(item => <article className="candidate" key={item.id}><strong>{item.course_name || item.section_id}</strong><span>{item.created_at} · {item.result} · {item.message}</span>{item.result === "unknown" && !item.resolved && <button onClick={() => resolveUnknown(item.id)}>已核实，解除阻断</button>}</article>)}<button className="danger" onClick={clearHistory}>清除执行历史</button></section>}
    <ScheduleBoard timetable={timetable} selection={selection} graduationProgress={state.graduation_progress} onExecuteSection={executeSection} executionPending={activeTask?.operation === "execute-selection"} />
  </main>;
}

createRoot(document.getElementById("root")!).render(<App />);
