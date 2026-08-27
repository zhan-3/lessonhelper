// Generated from course_selection/openapi.yaml. Run npm run generate:api after contract edits.
export type TaskOperation = "connect" | "refresh-selection" | "refresh-timetable" | "observe-navigation";
export interface Task { id:string; operation?:string; state:string; progress?:Record<string,unknown>; error?:string }
export interface CandidateNotice { version_id:string; status:string; title?:string; term?:string; query_eligible?:boolean; [key:string]:unknown }
export interface Snapshot { id:string; kind:string; term:string; source:string; source_at:string; payload:Record<string,unknown> }
export interface CompletedCourseFact { code:string; name:string; category:string; credits:number }
export interface RequirementProgress { key:string; label:string; required_credits:number; completed_credits:number; remaining_credits:number; courses:CompletedCourseFact[] }
export interface GraduationProgressReport { generated_at?:string; baseline_version?:string; data_complete:boolean; progress:RequirementProgress[] }
export interface GraduationProgress { status:"ready"|"incomplete"|"missing"|"invalid"; report:GraduationProgressReport|null }
export interface WorkbenchState { profile:Record<string,unknown>|null; confirmed_notice:Record<string,unknown>|null; snapshots:{selection:Snapshot|null;timetable:Snapshot|null}; snapshot_changes:{selection:Record<string,unknown>|null;timetable:Record<string,unknown>|null}; latest_plan:Record<string,unknown>|null; graduation_progress:GraduationProgress; stale:{selection:boolean;timetable:boolean}; academic_session:{state:string}; csrf_token:string }
