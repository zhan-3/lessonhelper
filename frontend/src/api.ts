import type { components } from "./api.generated";

export type TaskOperation = components["schemas"]["TaskRequest"]["operation"];
export type CandidateNotice = components["schemas"]["CandidateNotice"];
export type Snapshot = components["schemas"]["Snapshot"];
export type CompletedCourseFact = components["schemas"]["CompletedCourseFact"];
export type RequirementProgress = components["schemas"]["RequirementProgress"];
export type GraduationProgressReport = components["schemas"]["GraduationProgressReport"];
export type GraduationProgress = components["schemas"]["GraduationProgress"];
export type LoginConfiguration = components["schemas"]["LoginConfiguration"];
type GeneratedWorkbenchState = components["schemas"]["WorkbenchState"];
export type Task = components["schemas"]["Task"] & { created_at?: string; updated_at?: string; timeout_seconds?: number };
export type SnapshotStatus = { status: "current" | "historical" | "incomplete" | "missing"; reason: string; source_at: string };
export type ExecutionHistory = { id: string; created_at: string; section_id: string; course_name: string; category: string; result: string; message: string; resolved: number };
export type WorkbenchState = Omit<GeneratedWorkbenchState, "active_task" | "snapshot_status" | "execution_history" | "academic_session"> & {
  active_task?: Task | null;
  snapshot_status?: Record<"selection" | "timetable" | "progress", SnapshotStatus>;
  execution_history?: ExecutionHistory[];
  academic_session: {
    state: string; browser?: string; webvpn?: string; last_verified_at?: string;
  };
};
