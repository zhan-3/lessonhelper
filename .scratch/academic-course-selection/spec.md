# Persistent Academic Selection Workbench

**Status:** ready-for-agent

## Problem Statement

The student currently has separate commands and local files for authentication, notice parsing, timetable import, and course-selection discovery. A real course-selection query has succeeded, but each command owns and closes its own browser, repeated runs commonly repeat authentication, the real selection result is not connected to the workbench, and timetable discovery does not yet produce a real timetable snapshot. Partial reads and raw diagnostic pages also make it too easy to confuse an incomplete attempt with current academic facts or to retain sensitive material unnecessarily.

The student needs one long-running local workbench that keeps a single authenticated academic session available, reads the real timetable and notice-approved candidate sections, preserves the last complete result when refreshes fail, and presents those facts in an app-like weekly timetable for conflict-aware planning. The workbench must remain local and read-only in this phase.

## Solution

Provide a single-instance local application with Flask as a replaceable JSON API adapter, React as the planning interface, SQLite as the sole writable source of truth, and one dedicated worker that owns a visible Playwright-bundled Chromium academic session. The browser starts lazily, remains alive until the service exits, restores or renews authentication when necessary, and performs serialized observation tasks without sending course-selection mutations.

Timetable and candidate-section refreshes run as asynchronous tasks. Each successful refresh atomically creates a complete, versioned academic snapshot; an incomplete or failed refresh attempt records its state but never replaces the previous snapshot. The workbench reads the timetable directly from the verified academic interface, retains spreadsheet import as a clearly labelled fallback, and reads every page for only the course categories allowed by the confirmed notice and student profile.

The React interface immediately renders the latest applicable snapshots, their source and age, the academic-session state, and asynchronous refresh progress. It displays the existing timetable as a weekly grid and overlays user-focused candidate sections in grey, with explicit conflict and unknown-conflict states. Official undergraduate notices resembling the university's “course-selection time arrangements” notices can be discovered on user request or an optional long interval, but every content version remains a candidate notice until the student confirms it.

## User Stories

1. As a student, I want one local start command, so that I do not have to coordinate separate backend, frontend, and browser commands.
2. As a student, I want a second launch attempt to reuse the existing workbench, so that I never accidentally run two browsers against the academic system.
3. As a student, I want the academic browser to start only when protected data is needed, so that opening cached data does not trigger login or network activity.
4. As a student, I want the academic browser to remain open while the workbench runs, so that successive refreshes reuse the same academic session.
5. As a student, I want the workbench to use Playwright's bundled visible Chromium, so that browser behaviour is predictable and I can complete interactive authentication when required.
6. As a student, I want saved authentication state to be restored when possible, so that I log in less often.
7. As a student, I want encrypted saved credentials to be used only on the exact university CAS page, so that they cannot be filled into an unrelated WebVPN or application page.
8. As a student, I want an expired session to be renewed inside the existing browser, so that a refresh does not close halfway through authentication.
9. As a student, I want a refresh to pause in a visible waiting-for-authentication state, so that I can complete a CAPTCHA or changed login flow without losing the task.
10. As a student, I want the paused refresh to continue after successful authentication, so that I do not need to start the operation again.
11. As a student, I want the browser to remain available after an authentication timeout, so that a later retry can recover without restarting the whole service.
12. As a student, I want the workbench to detect a browser I closed manually, so that the next refresh can rebuild the academic session cleanly.
13. As a student, I want only one academic task to operate the browser at a time, so that timetable and selection pages cannot corrupt each other's navigation state.
14. As a student, I want duplicate pending refreshes to be coalesced, so that repeated clicks do not create repeated university requests.
15. As a student, I want to cancel a queued refresh immediately, so that unnecessary work never reaches the university system.
16. As a student, I want a running refresh to cancel safely between requests, so that cancellation does not destroy the reusable academic session.
17. As a student, I want to see whether a task is queued, connecting, waiting for authentication, reading data, complete, cancelled, or failed, so that long operations are understandable.
18. As a student, I want page and category progress during candidate-section refreshes, so that a multi-page read does not appear frozen.
19. As a student, I want my student profile stored locally, so that notices and requirement baselines can be matched to my 2025 cohort without repeated input.
20. As a student, I want the workbench to read my real academic timetable directly, so that conflict planning does not depend on manual spreadsheet export.
21. As a student, I want a timetable refresh to include every course meeting, week range, parity, period, location, and stable course identity available from the academic interface, so that conflict checks reflect the real schedule.
22. As a student, I want spreadsheet timetable import to remain available, so that I have a fallback when the academic interface changes or is temporarily unavailable.
23. As a student, I want imported timetable data labelled as a user-imported snapshot, so that it is never mistaken for a live academic read.
24. As a student, I want timetable and candidate-section refreshes to be independent, so that a failure in one does not prevent a complete snapshot of the other.
25. As a student, I want a “refresh all” action, so that I can intentionally update both kinds of academic data in one workflow.
26. As a student, I want the last complete timetable to remain visible after a failed refresh, so that a network failure does not erase my schedule.
27. As a student, I want the last complete candidate-section list to remain visible after one category or page fails, so that partial data is not presented as the current complete list.
28. As a student, I want every snapshot to show its source time and refresh source, so that I know how current and authoritative it is.
29. As a student, I want old timetable snapshots to be marked stale after 24 hours, so that I notice when schedule data may need refreshing.
30. As a student, I want old candidate-section snapshots to be marked stale after 30 minutes, so that changing availability is not mistaken for current capacity.
31. As a student, I want planning to require an applicable, sufficiently fresh snapshot, so that a stale or mismatched result cannot silently drive decisions.
32. As a student, I want a snapshot bound to the student-profile version, academic term, and confirmed-notice version, so that data from another planning context is not reused as current data.
33. As a student, I want historical snapshots to remain viewable, so that I can understand what changed without treating history as current.
34. As a student, I want the workbench to retain recent structured snapshot history without raw pages, so that diagnostics do not unnecessarily retain tokens or identity data.
35. As a student, I want existing local JSON data imported once into the new database, so that the upgrade does not discard my profile, confirmed notice, timetable, or prior read-only result.
36. As a student, I want old JSON files preserved as read-only migration evidence, so that migration is recoverable without maintaining two writable stores.
37. As a student, I want database upgrades backed up and transactional, so that a failed application upgrade cannot leave a half-migrated workspace.
38. As a student, I want only notice windows matching my profile, selection action, academic-system method, and course category to enter the query whitelist, so that unrelated categories are never queried by guesswork.
39. As a student, I want drop windows, application processes, paper forms, email processes, and other cohorts excluded from the query whitelist, so that the system follows the official notice precisely.
40. As a student, I want every page of every whitelisted category read through the verified read-only query, so that a 20-row first page is not mistaken for the full list.
41. As a student, I want the list to distinguish a true empty result, a not-yet-open round, no matching round, authentication failure, and an incomplete read, so that zero records always has a trustworthy meaning.
42. As a student, I want candidate sections to include course code, name, credits, teacher, schedule, campus, department, audience, requirements, capacity, and selected state when the source exposes them, so that I can compare real options.
43. As a student, I want course identity and course-section identity kept separate, so that different sections of one course can be compared without becoming duplicate course goals.
44. As a student, I want source task identifiers scoped to their academic term, so that historical refreshes do not accidentally merge unrelated sections.
45. As a student, I want missing source identifiers to fall back to a documented composite identity, so that sections remain usable when the university omits an identifier.
46. As a student, I want composite capacity labels preserved without inventing a misleading total, so that gender or audience quotas remain accurate.
47. As a student, I want schedules that cannot be parsed marked as conflict unknown, so that “time pending” is never interpreted as free time.
48. As a student, I want applicable course-selection categories shown before, during, and after their windows, so that I can plan ahead and review history.
49. As a student, I want each notice window labelled not started, open, or ended, so that only currently actionable windows appear actionable.
50. As a student, I want the workbench to discover only official undergraduate notices similar to “course-selection time arrangements”, so that incidental articles containing the word “selection” do not change my workflow.
51. As a student, I want discovered notice text to contain parseable term, time, cohort, category, action, and method facts, so that incomplete articles cannot trigger course queries.
52. As a student, I want discovered notices to remain candidates until I confirm them, so that automated discovery cannot silently alter the query whitelist.
53. As a student, I want a changed notice body at the same URL to create a new immutable version, so that an official correction cannot overwrite what I previously confirmed.
54. As a student, I want to compare a changed candidate notice with the confirmed version, so that I can understand the correction before accepting it.
55. As a student, I want notice checks to run when I request them, so that the application does not continuously poll the university.
56. As a student, I want an optional seven-day notice check only while the workbench is already running, so that long-cycle discovery never wakes the application or keeps authentication alive.
57. As a student, I want public notice checks to avoid opening the academic browser, so that a public website read does not cause unnecessary authentication.
58. As a student, I want the weekly timetable to be the primary workbench view, so that existing commitments and candidate options are understood spatially.
59. As a student, I want focused candidate sections overlaid as translucent grey blocks, so that I can compare them with my current timetable without filling every empty cell.
60. As a student, I want conflicts with my current timetable outlined in red, so that hard conflicts are immediately visible.
61. As a student, I want conflicts between candidate sections outlined in orange, so that competing plans are distinguishable from current-course conflicts.
62. As a student, I want conflict-unknown sections called out separately, so that I can review them manually before planning.
63. As a student, I want to create ranked course goals, so that the planner understands which courses matter most.
64. As a student, I want to rank acceptable sections within one course goal, so that teacher and time preferences do not become separate course priorities.
65. As a student, I want the planning view to remain read-only in this phase, so that arranging preferences cannot accidentally submit a selection.
66. As a mobile student, I want a day-focused portrait view with swipe navigation, so that the timetable remains readable on a narrow screen.
67. As a desktop or landscape user, I want a seven-day weekly view, so that I can compare gaps and conflicts across the whole week.
68. As a student, I want to switch to a horizontally scrollable full-week view on mobile, so that compact access does not remove the overview.
69. As a student, I want the interface to keep showing an old snapshot alongside a visible stale or failed-refresh warning, so that an error does not turn the page blank.
70. As a student, I want the local service bound only to my machine, so that nearby devices cannot access academic data or future execution controls.
71. As a student, I want all state-changing local APIs protected against cross-origin submission, so that a website open in another tab cannot operate my workbench.
72. As a maintainer, I want the frontend API types generated from the backend contract, so that task and snapshot fields cannot drift silently.
73. As a maintainer, I want framework-independent use cases and repositories, so that replacing Flask does not require rewriting academic behaviour.
74. As a maintainer, I want observation tasks to block known and suspected selection mutations, so that real-interface verification remains safely read-only.
75. As a maintainer, I want sensitive request fields removed from structured diagnostics and raw HTML omitted entirely, so that local debugging does not retain credentials, tokens, student number, or student name.

## Implementation Decisions

- Keep Flask as a local JSON API adapter and production static-asset host; academic use cases, persistence, task orchestration, and domain models must not depend on Flask.
- Build the planning interface with React, TypeScript, and Vite. Development may run Vite and Flask separately; the production build is served by Flask so users retain one start command and do not need Node at runtime.
- Maintain an OpenAPI contract for the JSON API and generate the frontend's TypeScript API types from it.
- Bind the local service to loopback only. Protect local state-changing API requests with same-origin enforcement and a CSRF token.
- Enforce a workspace-level single-instance lock. A second invocation opens or points to the existing workbench and exits without starting another browser or server. Stale locks left by crashes must be safely recoverable.
- Introduce an application service boundary for connect, refresh-timetable, refresh-selection, refresh-all, cancel-task, inspect-task, inspect-snapshot, notice discovery, notice confirmation, and planning operations.
- Introduce an academic gateway boundary representing authenticated university reads. Production uses the Playwright-backed gateway; tests use a deterministic gateway through the same interface.
- A single dedicated worker thread creates and exclusively owns Playwright, the browser context, pages, and all academic-gateway calls. Flask requests communicate with it through task submission rather than touching Playwright objects.
- Launch Playwright's bundled visible Chromium lazily on the first protected operation. Do not silently fall back to system Chrome. Keep the browser alive until service shutdown or unrecoverable browser failure.
- Restore the saved browser state when possible. If WebVPN or CAS authentication has expired, renew it in the same browser. Credentials remain DPAPI-protected and may be filled only into the exact HTTPS university CAS credential page.
- Manual browser interaction is supported only while the session is waiting for authentication. Unexpected user navigation during an observation task aborts that refresh attempt without replacing a snapshot.
- Waiting for authentication pauses the task for up to ten minutes and exposes that state to the API. Successful authentication resumes the same task. Timeout fails the attempt but does not intentionally destroy a healthy browser context.
- If the browser is manually closed or crashes, mark the academic session disconnected. The next protected task rebuilds it and restores authentication rather than crashing the service.
- Run academic tasks serially. Coalesce equivalent queued refreshes and allow independent timetable and selection refreshes plus a combined refresh-all operation.
- Cancellation is cooperative: queued work cancels immediately; running work stops between network requests or pages. Browser termination is reserved for an unresponsive browser, not ordinary task cancellation.
- Task states include queued, connecting, waiting for authentication, reading, succeeded, failed, cancel requested, and cancelled. Reading progress includes operation-specific category and page details.
- Keep observation and execution capabilities distinct. All work in this spec is an observation task and blocks course selection, drop, save, registration, and other mutation requests. A future execution task will require a separately confirmed write contract and explicit user authorization.
- Use SQLite as the sole writable source of truth for student profiles, immutable notice versions and confirmations, academic snapshots, timetable entries, candidate sections, plans, refresh tasks, and structured task diagnostics.
- Use the standard SQLite interface behind repository modules rather than an ORM. Enable foreign keys, WAL mode, and a finite busy timeout.
- Use numbered SQL migrations and a schema-version table. Back up the local database before upgrades and run migrations transactionally; a failed migration rolls back and prevents the workbench from opening against a partially upgraded schema.
- Import existing supported JSON state once. Preserve imported files as read-only migration evidence and do not dual-write SQLite and JSON afterward.
- Keep credentials and browser authentication state outside SQLite. Do not store raw response HTML, Cookie values, authentication tokens, student number, student name, or passwords in snapshots or diagnostics.
- Model snapshot metadata relationally and timetable entries and candidate sections as queryable rows. Diagnostic details that are not used by planning may use structured JSON columns.
- Bind every applicable snapshot to a student-profile version and academic term. Candidate-section snapshots additionally bind to a specific confirmed-notice version and its query whitelist.
- Write a refresh attempt into a new snapshot version and publish it atomically only after every required source, category, and page succeeds. Failed and incomplete attempts record a sanitized task outcome and leave the previous snapshot current.
- Retain the current snapshot and the most recent twenty successful structured snapshots for each snapshot kind. Retention never includes raw university pages.
- Mark timetable snapshots stale after 24 hours and candidate-section snapshots stale after 30 minutes. Display stale data, but require an applicable sufficiently fresh snapshot before planning can be treated as ready.
- Do not implement general background polling or active keepalive. Notice discovery is user-triggered by default, with an optional seven-day check only while the service is already running. It must not launch a closed service, wake the academic browser, or preserve authentication.
- Prefer a public official-notice source without opening the academic browser. Reuse the existing academic session only when the configured official source actually requires WebVPN.
- Restrict notice discovery to configured official undergraduate teaching-notice listings and notices whose title and content resemble comprehensive course-selection time arrangements. Keyword presence alone is insufficient.
- Require parseable academic term, time window, cohort, category, action, and handling method before a candidate notice can drive any academic query.
- Identify immutable notice versions from source identity and content. A changed body at the same URL creates a new candidate version and never overwrites or silently supersedes the confirmed version. Confirmation remains a user decision.
- Derive the query whitelist only from confirmed selection windows matching the current student profile, selection action, academic-system method, and explicit category mappings. Exclude drop windows, applications, paper/email processes, and inapplicable cohorts.
- Read the real timetable through the authenticated academic interface and normalize it into timetable entries. Retain spreadsheet import as a user-imported-snapshot fallback with explicit source labelling.
- Read candidate sections through the verified read-only course query, explicitly setting the confirmed term and each whitelisted category. Follow every server-rendered page and report completeness per category and snapshot.
- Separate course identity from course-section identity. Prefer the university task identifier for a section within a term; if absent, use a normalized composite of course code, teacher, schedule, and campus. Never use a page row number as identity.
- Preserve source capacity labels. Populate scalar selected/capacity fields only when one unambiguous quota pair exists; do not flatten composite gender or audience quotas into a false total.
- Parse and expose selected state when the source provides it. Absence of evidence remains unknown rather than false.
- Classify login required, entry unreachable, round not open, no matching round, true empty result, interface unconfirmed, incomplete read, and ready result as distinct observable states.
- Represent unparseable, pending, or missing schedules as conflict unknown. They cannot be treated as free and require user review before inclusion in a final plan.
- Expose task, session, notice, snapshot, timetable, candidate-section, and planning resources through the JSON API. Page loads read stored state and do not implicitly trigger protected university requests.
- Make the weekly timetable the primary React view. Show academic term, academic-session status, source, data-as-of time, staleness, and refresh controls alongside task progress.
- Overlay only focused or selected candidate sections as translucent grey blocks. Use red treatment for conflicts with the current timetable, orange for conflicts among candidates, and a distinct warning for conflict unknown.
- Model planning in two levels: ranked course goals and ordered acceptable section preferences within each goal. Planning is local and read-only under this spec.
- Use a seven-day view on desktop and landscape screens. Use a day-focused swipeable view by default on portrait mobile, with an optional horizontally scrollable week view.
- Preserve existing data on refresh errors and keep it visible with explicit warning state. A failed refresh must never render as an empty timetable or candidate list.

## Testing Decisions

- The primary automated test seam is the Flask JSON API through application services, asynchronous task orchestration, repositories, and a temporary real SQLite database, with the university replaced by a deterministic academic-gateway adapter.
- Tests assert externally observable HTTP resources, task-state transitions, published snapshots, retained previous snapshots, and database-visible outcomes. They do not assert private methods, exact thread implementation, SQL statement text, or internal call counts unless the count is itself a university-safety contract.
- API-level tests cover lazy connection, single task serialization, duplicate refresh coalescing, cooperative cancellation, waiting for authentication and resume, browser-disconnected recovery signals, and independent versus combined refreshes.
- Snapshot tests cover full successful publication, atomic rollback on any failed page/category, retention of the previous snapshot, staleness, applicability to profile/term/notice versions, source labelling, and recent-history retention.
- Migration tests start from representative existing JSON workspaces and supported older database schema versions, then verify one-time import, transactional upgrade, rollback on failure, and preservation of migration evidence.
- Notice tests use official-notice-shaped fixed samples to cover strict source/title matching, complete multi-window parsing, cohort/action/method filtering, content-version changes, diff-ready immutable candidates, and the requirement for explicit confirmation.
- Timetable and candidate-section adapter tests use sanitized fixed HTML/network samples. They cover pagination, server field changes that must fail loudly, course and section identity, teacher/schedule parsing, composite quotas, selected-state extraction, every empty/not-open/no-match state, and complete redaction.
- Existing timetable import tests are prior art for spreadsheet parsing and term validation. Existing selection-entry and discovery tests are prior art for read-only request guards, response classification, sanitization, and legacy HTML parsing. Existing Flask test-client tests are prior art for the highest API seam.
- Frontend component tests cover the weekly/day views, grey candidate overlays, conflict colours, unknown-conflict warnings, stale/error banners, task progress, and generated API-type usage. Browser-level frontend tests exercise only critical user workflows rather than duplicating component details.
- Security tests verify loopback binding configuration, same-origin/CSRF enforcement for state changes, exact-host credential filling, mutation-request blocking for observation tasks, and absence of raw HTML and sensitive identity/authentication values from SQLite and diagnostics.
- The live acceptance test is opt-in and manual: start one workbench, connect one visible bundled Chromium, refresh the real timetable, refresh all notice-approved candidate categories, and refresh again. The same browser must remain open, no second browser may appear, reauthentication must not occur unless the server actually expires it, all reads must complete or preserve old snapshots, and no selection mutation may be sent.
- The first tracer-bullet acceptance requires one command to start the workbench, one user-triggered connection, real timetable and candidate-section snapshots in SQLite, and a React weekly view that overlays candidate sections. Complex automatic ranking and all course-selection mutations are not required for that first slice.

## Out of Scope

- Submitting, dropping, saving, registering, or otherwise changing course-selection state.
- Active session keepalive, high-frequency background polling, or launching the application solely to check notices.
- Automatic confirmation of a discovered or changed official notice.
- Treating a listed candidate section as proof that the student is eligible, the section has capacity, or submission will succeed.
- Guessing course categories, timetable fields, write endpoints, or request contracts when the university interface is not confirmed.
- Multi-user hosting, remote deployment, LAN/mobile access to the local service, account isolation, or cloud synchronization.
- Replacing local DPAPI credential protection or adding database encryption in the first implementation.
- Complex automatic course-goal ranking, optimization, or automated execution in the first tracer-bullet slice.
- Recreating the university's original spreadsheet as the primary timetable representation; original-file download may be added later as a diagnostic convenience.
- Saving raw university HTML as normal application state or diagnostics.

## Further Notes

- The current live discovery has already verified a real read-only candidate-course query and multi-page results for the 2025 cohort. That evidence should be used to build the production academic gateway, but existing per-run raw response files are migration/debug evidence rather than the future persistence model.
- The current timetable discovery recognizes the relevant academic area but has not yet produced a real normalized timetable snapshot; confirming and implementing that read contract is part of this spec.
- Current code closes browser contexts at command exit, and the existing persistent-session option persists profile data rather than a running singleton. The new worker lifecycle replaces that command-scoped ownership for the workbench.
- The existing read-only review found selected-state extraction, exact window matching, no-match versus true-empty classification, reusable query-client shape, and raw HTML retention incomplete. These gaps are requirements of this spec rather than accepted limitations.
- Domain language is defined in the root glossary. In particular, course identity, course-section identity, refresh attempt, academic snapshot, applicable snapshot, observation task, execution task, candidate notice, and notice version must not be collapsed into interchangeable terms.
- Architectural decisions are recorded separately for keeping the application core independent from Flask, using SQLite as the local source of truth, and using React for the planning interface.
