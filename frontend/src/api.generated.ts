// Generated from course_selection/openapi.yaml. Run npm run generate:api after contract edits.
export type TaskOperation = "connect" | "refresh-selection" | "refresh-timetable";
export interface Task { id:string; operation?:string; state:string; progress?:Record<string,unknown>; error?:string }
export interface CandidateNotice { version_id:string; status:string; title?:string; term?:string; query_eligible?:boolean; [key:string]:unknown }
export interface Snapshot { id:string; kind:string; term:string; source:string; source_at:string; payload:Record<string,unknown> }
export interface WorkbenchState { profile:Record<string,unknown>|null; confirmed_notice:Record<string,unknown>|null; snapshots:{selection:Snapshot|null;timetable:Snapshot|null}; latest_plan:Record<string,unknown>|null; stale:{selection:boolean;timetable:boolean}; academic_session:{state:string}; csrf_token:string }
