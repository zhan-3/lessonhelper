# Student Profile Interface Analysis and Local Login

**Status:** implemented-reviewed

## Review notes

Parallel code review (standards + spec axes) found three P1s; all addressed:

1. **Privacy leak (fixed)** — candidate artifacts could carry identity values via raw JSON object keys and URL fragments. `student_profile_observation.py` now redacts dynamic keys (`[dynamic-key]`), strips URL fragments, and uses a structural summarizer (`_safe_structure`) instead of raw key dumps; regression test `test_identity_values_used_as_keys_and_url_fragments_are_redacted`.
2. **Reset-login race / lost shell (fixed)** — login configuration no longer races the frontend `connect`; `POST /api/login-configuration` now atomically (under `run_when_idle`) configures, enqueues `reset-login` → `launch-shell` → `connect`, and returns the created connection task id. `DELETE` runs the same idle-guarded sequence. `PlaywrightAcademicGateway.reset_login()` now fails closed if the profile dir cannot be removed and refuses reconnect (`_login_reset_error` blocks all non-reset tasks); `_shell_url` is preserved across gateway replacement via the `launch-shell` context.
3. **Graduation progress ignoring snapshot applicability (fixed)** — `WorkbenchService.graduation_progress()` returns `not_applicable` when the progress snapshot's `profile_id` no longer matches the current profile or when `baseline_version != "guide-2026"`; OpenAPI + generated types updated (`not_applicable` status, `Snapshot`/`GraduationProgress` schemas).

Post-review verification: 129 Python tests, 9 frontend tests, `tsc -b`, Vite build, and Ruff all pass. Committed as `d95f5d0` on `develop`.

## Intent

Provide a minimal local login setup and a separately invoked, strictly read-only maintenance tool that discovers candidate official student-profile interfaces. The goal is to extract verifiable interface contracts, not clone or run the university frontend.

## Requirements

### Local login setup

- The production workbench presents a local setup screen before protected academic tasks can run.
- A student number and password are accepted only by the loopback service, protected by same-origin and CSRF checks, and encrypted at rest with Windows DPAPI.
- Passwords must not enter SQLite, API responses, logs, snapshots, diagnostics, or generated frontend assets.
- Routine state reads must not decrypt the saved password.
- The browser may autofill credentials only on the exact official HTTPS CAS credential page.
- Local HTTP responses use explicit loopback Host validation, CSP with `frame-ancestors 'none'`, no-referrer, no-sniff, and `Cache-Control: no-store` for APIs.
- Clearing or changing login credentials must not silently mix two students' browser sessions or academic data. Full identity switching remains blocked until an official identity contract exists.

### Student-profile interface analyzer

- Expose only `analyze-interface --target student-profile` in the first version.
- Reuse the persistent Playwright Chromium and `AcademicBrowserSession`; do not implement another credential or browser lifecycle.
- Verify WebVPN capabilities through `/user/info` and `/user/portal_groups` before treating the portal as ready.
- Resolve the academic application dynamically from the portal and accept only an HTTPS WebVPN proxy path.
- Use a strict observation policy: GET/HEAD/OPTIONS and exact official authentication POSTs are allowed; unknown POSTs are blocked.
- Observe only short-lived XHR/fetch JSON responses and produce candidate profile contracts from relevant field names.
- Candidate artifacts contain canonical/sanitized URLs, methods, statuses, matched field paths, and response structure. They never contain student name, student number, profile values, cookies, tickets, tokens, or raw identity responses.
- Do not modify SQLite, the current student profile, identity confirmation, or normal workbench state.
- Store artifacts only under `.private/interface-analysis/student-profile/<timestamp>/`; clean directories older than seven days on the next run.
- Do not implement JavaScript AST analysis, Source Map analysis, frontend mirroring, adapter generation, or automatic contract promotion in this tracer bullet.
- Real university access requires separate explicit authorization; offline implementation and tests must not access the university.

## Acceptance

- CLI rejects every target except `student-profile`.
- Relevant Chinese and English/pinyin field names produce a structure-only candidate.
- Unrelated application resource names do not produce candidates.
- Identity values and sensitive URL values are absent from serialized artifacts.
- Non-WebVPN and unrecognized portal entries are rejected.
- Unknown POSTs remain blocked by the reused strict policy.
- Seven-day cleanup is deterministic.
- Login configuration is DPAPI encrypted and absent from API output.
- Hostile Host headers are rejected and security headers are present.
- Python tests, frontend tests, typechecking/build, and lint for changed modules pass.
