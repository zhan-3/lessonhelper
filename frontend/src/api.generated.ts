// Generated from course_selection/openapi.yaml. Run npm run generate:api after contract edits.
export type TaskOperation = "connect" | "refresh-selection" | "refresh-timetable";
export interface Snapshot { id:string; kind:string; term:string; source:string; source_at:string; payload:Record<string,unknown> }
export interface WorkbenchState { profile:Record<string,unknown>|null; confirmed_notice:Record<string,unknown>|null; snapshots:{selection:Snapshot|null;timetable:Snapshot|null}; stale:{selection:boolean;timetable:boolean}; academic_session:{state:string}; csrf_token:string }
