## Project contract

This repository is a Windows-local, privacy-first experimental campus academic assistant. It is a development snapshot, not a stable release or a production booking service.

### Feature maturity

- Lab booking is incomplete and has not passed real-environment validation. Describe its commands and code as experimental, not usable or reliable.
- Course selection has substantial implementation and automated tests, but its real academic-system read and submission flows have not passed complete acceptance testing.
- Graduation progress is advisory. Present its output as an estimate that requires manual verification, never as an official graduation determination.
- Keep README and release-facing claims aligned with verified evidence. Automated tests prove code behavior, not compatibility with the live university systems.

### Safety boundary

- Keep credentials, cookies, browser profiles, student records, timetables, and booking results local and Git-ignored. Before sharing or committing, verify that `.private/`, `.env`, `storage_state.json`, `*.xls`, and `*.xlsx` are untracked.
- Preserve the distinction between observation tasks and execution tasks defined in `CONTEXT.md`.
- Require explicit user confirmation for a specific teaching section before any selection submission. Submit at most once and do not add automatic retries for unknown results.
- Prefer a visible browser for authentication and state-changing operations. A future background mode must begin from a user-established session, surface authentication requirements, and retain explicit stop and failure safeguards.

### Completion claims

Classify live-system features as `implemented`, `automated-test verified`, or `real-environment verified`. Claim the strongest status only when corresponding evidence exists. Record unresolved live verification as a release risk rather than inferring success from mocks, fixtures, or page selectors.

## Agent skills

### Academic system fieldwork

Use `.agents/skills/academic-system-fieldwork/` when exploring authenticated campus systems, diagnosing browser targets or network contracts, publishing live academic facts, or crossing a state-changing boundary.

### Issue tracker

Issues are tracked as local Markdown files under `.scratch/<feature>/`. See `docs/agents/issue-tracker.md`.

### Triage labels

Uses the default canonical triage labels. See `docs/agents/triage-labels.md`.

### Domain docs

This is a single-context repository using root `CONTEXT.md` and `docs/adr/`. See `docs/agents/domain.md`.
