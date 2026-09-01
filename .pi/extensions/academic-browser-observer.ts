import type { ExtensionAPI } from "@earendil-works/pi-coding-agent";
import { Type } from "typebox";

const DEFAULT_BASE = "http://127.0.0.1:5000";
const BASE = process.env.ACADEMIC_OBSERVER_URL?.replace(/\/$/, "") ?? DEFAULT_BASE;
const BASE_URL = new URL(BASE);
if (!new Set(["127.0.0.1", "localhost", "::1"]).has(BASE_URL.hostname)) {
  throw new Error("ACADEMIC_OBSERVER_URL must be loopback-only");
}
const ORIGIN = BASE_URL.origin;
const MAX_TARGETS = 20;

type Envelope = {
  status: string;
  summary?: string;
  data?: Record<string, unknown>;
  warnings?: string[];
  next_actions?: string[];
};

function compact(value: Envelope): Envelope {
  const data = value.data ?? {};
  const targets = Array.isArray(data.targets) ? data.targets.slice(0, MAX_TARGETS) : undefined;
  const candidates = Array.isArray(data.candidates) ? data.candidates.slice(0, 10).map((value) => {
    const item = value as Record<string, unknown>;
    return {
      evidence_identity: item.evidence_identity,
      method: item.method,
      path_shape: item.path_shape,
      count: item.count,
      score: item.score,
      completeness: item.completeness,
      resource_types: item.resource_types,
      provenance_category: item.provenance_category,
      redirect_hop_count: item.redirect_hop_count,
      redirect_path_shapes: Array.isArray(item.redirect_path_shapes) ? item.redirect_path_shapes.slice(0, 11) : [],
      reasons: Array.isArray(item.reasons) ? item.reasons.slice(0, 5) : [],
      warnings: Array.isArray(item.warnings) ? item.warnings.slice(0, 3) : [],
    };
  }) : undefined;
  const changes = data.target_changes as { added?: unknown[]; removed?: unknown[] } | undefined;
  return {
    status: value.status,
    summary: value.summary,
    warnings: (value.warnings ?? []).slice(0, 10),
    next_actions: (value.next_actions ?? []).slice(0, 5),
    data: {
      connection: data.connection,
      detach_only: data.detach_only,
      trace_id: data.trace_id,
      target_count: data.target_count,
      targets,
      candidates,
      event_count: data.event_count,
      dropped_events: data.dropped_events,
      missing_evidence: Array.isArray(data.missing_evidence) ? data.missing_evidence.slice(0, 20) : undefined,
      target_changes: changes ? {
        added_count: changes.added?.length ?? 0,
        removed_count: changes.removed?.length ?? 0,
      } : undefined,
    },
  };
}

function result(envelope: Envelope) {
  let bounded = compact(envelope);
  let text = JSON.stringify(bounded);
  if (Buffer.byteLength(text, "utf8") > 16_384) {
    bounded = {
      ...bounded,
      warnings: [...(bounded.warnings ?? []), "model-facing output exceeded 16KB; target details omitted"],
      next_actions: [...(bounded.next_actions ?? []), "request a narrower bounded inspection"],
      data: { ...bounded.data, targets: undefined },
    };
    text = JSON.stringify(bounded);
  }
  return {
    content: [{ type: "text" as const, text }],
    details: bounded,
  };
}

export default function academicBrowserObserver(pi: ExtensionAPI) {
  let csrf = "";
  let activeTrace = "";
  let shuttingDown = false;

  async function token(signal?: AbortSignal): Promise<string> {
    if (csrf) return csrf;
    const response = await fetch(`${BASE}/api/state`, { signal });
    if (!response.ok) throw new Error(`observer state unavailable: HTTP ${response.status}`);
    const state = await response.json() as { csrf_token?: string };
    if (!state.csrf_token) throw new Error("observer CSRF token unavailable");
    csrf = state.csrf_token;
    return csrf;
  }

  async function call(path: string, method: "GET" | "POST", body?: unknown, signal?: AbortSignal, retried = false): Promise<Envelope> {
    const headers: Record<string, string> = {};
    if (method === "POST") {
      headers.Origin = ORIGIN;
      headers["X-CSRF-Token"] = await token(signal);
      headers["Content-Type"] = "application/json";
    }
    const response = await fetch(`${BASE}${path}`, {
      method, headers, signal,
      body: body === undefined ? undefined : JSON.stringify(body),
    });
    if (response.status === 403 && method === "POST" && !retried) {
      csrf = "";
      return call(path, method, body, signal, true);
    }
    let payload: Envelope & { error?: string };
    try {
      payload = await response.json() as Envelope & { error?: string };
    } catch {
      payload = { status: "failed", summary: `observer returned HTTP ${response.status}` };
    }
    if (!response.ok) {
      return { status: "failed", summary: payload.error ?? payload.summary ?? `HTTP ${response.status}`, next_actions: ["verify the loopback workbench is running"] };
    }
    return payload;
  }

  async function stopAndDetach(): Promise<void> {
    if (shuttingDown) return;
    shuttingDown = true;
    try { await call("/api/browser-observer/stop", "POST"); } catch { /* best effort */ }
    try { await call("/api/browser-observer/disconnect", "POST"); } catch { /* best effort */ }
    activeTrace = "";
    csrf = "";
  }

  async function execute(path: string, method: "GET" | "POST", body: unknown, signal?: AbortSignal) {
    try {
      const envelope = await call(path, method, body, signal);
      const trace = envelope.data?.trace_id;
      if (typeof trace === "string") activeTrace = trace;
      if (path.endsWith("/stop") || path.endsWith("/disconnect")) activeTrace = "";
      if (path.endsWith("/checkpoint") || path.endsWith("/stop")) {
        const safe = compact(envelope);
        pi.appendEntry("academic-browser-observer", {
          trace_id: safe.data?.trace_id,
          status: safe.status,
          event_count: safe.data?.event_count,
          dropped_events: safe.data?.dropped_events,
          warnings: safe.warnings,
          timestamp: new Date().toISOString(),
        });
      }
      return result(envelope);
    } catch (error) {
      if (signal?.aborted) {
        try {
          const cancellation = await call("/api/browser-observer/cancel", "POST");
          if (cancellation.status === "cancelled") {
            activeTrace = "";
            return result(cancellation);
          }
          return result({ status: "partial", summary: "cancellation could not be confirmed", warnings: [cancellation.summary ?? "cancel endpoint failed"], next_actions: ["call academic_browser_stop before continuing"] });
        } catch (cancelError) {
          return result({ status: "partial", summary: "cancellation could not be confirmed", warnings: [cancelError instanceof Error ? cancelError.message : "cancel endpoint failed"], next_actions: ["call academic_browser_stop before continuing"] });
        }
      }
      return result({ status: "failed", summary: error instanceof Error ? error.message : "observer request failed", next_actions: ["verify the loopback workbench and explicit CDP endpoint"] });
    }
  }

  async function begin(endpoint: string, signal?: AbortSignal) {
    try {
      const connected = await call("/api/browser-observer/connect", "POST", { endpoint }, signal);
      if (!new Set(["connected", "already_connected"]).has(connected.status)) return result(connected);
      const inventory = await call("/api/browser-observer/targets", "GET", undefined, signal);
      if (inventory.status !== "connected") {
        await call("/api/browser-observer/disconnect", "POST");
        return result({ ...inventory, status: "failed", next_actions: ["verify the explicit endpoint and retry academic_browser_begin"] });
      }
      const started = await call("/api/browser-observer/start", "POST", undefined, signal);
      if (!new Set(["observing", "already_observing"]).has(started.status)) {
        await call("/api/browser-observer/disconnect", "POST");
        return result(started);
      }
      const trace = started.data?.trace_id;
      if (typeof trace === "string") activeTrace = trace;
      return result({
        status: started.status,
        summary: "borrowed browser connected, inventoried, and observation started",
        warnings: [...(connected.warnings ?? []), ...(inventory.warnings ?? []), ...(started.warnings ?? [])],
        next_actions: ["let the authorized external actor perform exactly one bounded read-only operation", "then call academic_browser_finish"],
        data: { ...started.data, connection: "borrowed", detach_only: true, target_count: inventory.data?.target_count, targets: inventory.data?.targets },
      });
    } catch (error) {
      if (signal?.aborted) {
        try { await call("/api/browser-observer/cancel", "POST"); } catch { /* best effort */ }
        try { await call("/api/browser-observer/disconnect", "POST"); } catch { /* best effort */ }
        activeTrace = "";
        return result({ status: "cancelled", summary: "combined observation start cancelled and detached" });
      }
      return result({ status: "failed", summary: error instanceof Error ? error.message : "combined observation start failed", next_actions: ["verify the loopback workbench and explicit CDP endpoint"] });
    }
  }

  async function finish(signal?: AbortSignal) {
    let stopped: Envelope;
    try {
      stopped = await call("/api/browser-observer/stop", "POST", undefined, signal);
    } catch (error) {
      stopped = { status: signal?.aborted ? "partial" : "failed", summary: error instanceof Error ? error.message : "observation stop failed", warnings: ["stop could not be confirmed before detach"] };
    }
    let detached: Envelope;
    try {
      detached = await call("/api/browser-observer/disconnect", "POST");
    } catch (error) {
      detached = { status: "failed", summary: error instanceof Error ? error.message : "detach failed" };
    }
    activeTrace = "";
    const envelope: Envelope = {
      ...stopped,
      status: detached.status === "disconnected" ? stopped.status : "partial",
      summary: `${stopped.summary ?? "observation stopped"}; detach=${detached.status}`,
      warnings: [...(stopped.warnings ?? []), ...(detached.warnings ?? []), ...(detached.status === "disconnected" ? [] : ["borrowed detach could not be confirmed"])],
      next_actions: detached.status === "disconnected" ? stopped.next_actions : ["confirm the borrowed browser is still alive and retry disconnect"],
      data: { ...(stopped.data ?? {}), connection: detached.status },
    };
    const safe = compact(envelope);
    pi.appendEntry("academic-browser-observer", { trace_id: safe.data?.trace_id, status: safe.status, event_count: safe.data?.event_count, dropped_events: safe.data?.dropped_events, warnings: safe.warnings, timestamp: new Date().toISOString() });
    return result(envelope);
  }

  pi.registerTool({
    name: "academic_browser_begin", label: "Begin Borrowed Observation",
    description: "In one bounded call, connect to one explicit loopback CDP endpoint, inspect sanitized targets, and start read-only observation. Never scans, launches, navigates, or closes a browser.",
    executionMode: "sequential",
    parameters: Type.Object({ endpoint: Type.String() }),
    execute: (_id, params, signal) => begin(params.endpoint, signal),
  });
  pi.registerTool({
    name: "academic_browser_finish", label: "Finish Borrowed Observation",
    description: "Stop observation, return compact sanitized candidates, and detach without closing the borrowed browser.",
    executionMode: "sequential",
    parameters: Type.Object({}),
    execute: (_id, _params, signal) => finish(signal),
  });

  pi.registerTool({
    name: "academic_browser_connect", label: "Connect Borrowed Browser",
    description: "Connect to one explicitly supplied loopback CDP endpoint without scanning, launching, or closing a browser.",
    executionMode: "sequential",
    parameters: Type.Object({ endpoint: Type.String() }),
    execute: (_id, params, signal) => execute("/api/browser-observer/connect", "POST", { endpoint: params.endpoint }, signal),
  });
  pi.registerTool({
    name: "academic_browser_inspect", label: "Inspect Browser Targets",
    description: "Return a compact sanitized target inventory from the borrowed browser.",
    executionMode: "sequential",
    parameters: Type.Object({}),
    execute: (_id, _params, signal) => execute("/api/browser-observer/targets", "GET", undefined, signal),
  });
  pi.registerTool({
    name: "academic_browser_start", label: "Start Browser Observation",
    description: "Start one bounded read-only observation before an external actor operates.",
    executionMode: "sequential",
    parameters: Type.Object({}),
    execute: (_id, _params, signal) => execute("/api/browser-observer/start", "POST", undefined, signal),
  });
  pi.registerTool({
    name: "academic_browser_checkpoint", label: "Checkpoint Browser Observation",
    description: "Return only compact sanitized counts and topology changes since the previous checkpoint.",
    executionMode: "sequential",
    parameters: Type.Object({}),
    execute: (_id, _params, signal) => execute("/api/browser-observer/checkpoint", "GET", undefined, signal),
  });
  pi.registerTool({
    name: "academic_browser_stop", label: "Stop Browser Observation",
    description: "Stop the active trace and release listeners without closing the borrowed browser.",
    executionMode: "sequential",
    parameters: Type.Object({}),
    execute: (_id, _params, signal) => execute("/api/browser-observer/stop", "POST", undefined, signal),
  });
  pi.registerTool({
    name: "academic_browser_disconnect", label: "Detach Borrowed Browser",
    description: "Detach safely without closing the remote browser or existing tabs.",
    executionMode: "sequential",
    parameters: Type.Object({}),
    execute: (_id, _params, signal) => execute("/api/browser-observer/disconnect", "POST", undefined, signal),
  });

  pi.on("session_start", async (event) => {
    shuttingDown = false;
    csrf = "";
    activeTrace = "";
    if (event.reason === "reload") {
      // Fail closed even if the old extension instance could not finish its
      // shutdown hook. Live state is never reconstructed from session entries.
      try { await call("/api/browser-observer/stop", "POST"); } catch { /* best effort */ }
      try { await call("/api/browser-observer/disconnect", "POST"); } catch { /* best effort */ }
      csrf = "";
      pi.appendEntry("academic-browser-observer", { status: "disconnected_after_reload", timestamp: new Date().toISOString() });
    }
  });

  pi.on("session_shutdown", async () => {
    if (activeTrace || csrf) await stopAndDetach();
  });
}
