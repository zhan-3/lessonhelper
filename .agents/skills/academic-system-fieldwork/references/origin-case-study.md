# Origin case study: HITWH local academic assistant

This skill was distilled from a privacy-reviewed mining pass over 22 Pi project sessions (8,064 entries), metadata from 86 Codex session files, selected project-linked Codex threads, and repository commits/tests. Raw conversations remain local because they contain sensitive history. The records below preserve changed decisions, not chat prose.

## Disproved assumptions

### “Leaving the login URL means authenticated”

The portal could render while protected XHR still redirected, ticket transitions could stall, and a short post-login binding transition could appear as logout. The implementation moved to layered authentication and protected destination probes.

Repository evidence: `course_progress/session.py`, `tests/test_course_progress.py`; commits `03fceeb`, `5d6e571`, `c36d9b9`.

### “The newest page or named iframe is the business target”

A collector attached to the wrong page and later waited for an iframe even though the protected application was top-level. Target resolution moved from position/name to protected origin, stable markers, and a valid read capability.

Repository evidence: `course_progress/session.py`, `tests/test_course_progress.py`.

### “A popup event means the application opened”

The portal created a live blank target whose navigation never committed. The workaround captured the authorized intended destination and navigated a controlled page, then validated the protected result.

Evidence tier: real-environment observation; generic tool support remains unimplemented.

### “Endpoint and method are the whole request contract”

A nominally correct probe failed because successful SPA traffic included additional request shape. Executable teaching-section identity also came from server/page action data, not course display fields. Contracts were expanded to headers, fields, redirects, provenance, identity, and dynamic-field acquisition.

Repository evidence: `docs/adr/0007-fix-academic-reads-to-identity-bearing-requests.md`, `course_selection/selection_execution.py`, and its tests.

### “HTTP success and parsed zero rows mean complete”

An early collection reported successful complete terms with zero records while the user could see records. Completion was split into transport, parse, domain acceptance, and all-segment publication gates.

Repository evidence: `docs/adr/0006-use-personal-timetable-snapshots-as-facts.md`, snapshot and deep-observation tests.

### “Discovery output can become application data”

Manual navigation and automatic discovery could find plausible routes but could not establish stable identity or completeness. Diagnostic and publishable result types were separated, and discovery was kept behind a development gate.

Repository evidence: `course_selection/deep_observation.py`, `tests/test_deep_observation.py`; commits `5e8b4fd`, `5e42a6b`.

### “Timeout after submit means failure”

A request may have reached the server before transport failed. Execution was redesigned around exact identity, fresh re-read, one POST, and durable unknown-result lockout.

Repository evidence: `course_selection/selection_execution.py`, `tests/test_selection_execution.py`; commits `3001c0c`, `4ee07dc`.

### “Closing an attached browser only detaches diagnostics”

A borrowed CDP diagnostic closed the user's browser. Competing supervisors also caused navigation storms. Ownership became explicit: one profile/port owner, borrowed sessions disconnect only.

Evidence tier: real-environment observation plus lifecycle implementation/tests.

### “Green tests mean live readiness”

A reviewed live path still referenced removed state after tests passed, and state-changing flows remained unverified against the real system. Feature reporting now separates implementation, automated tests, and real-environment evidence.

Repository evidence: `AGENTS.md`, README project status, acceptance tests, and session-derived review history.

## Remaining gaps

- Persistent redacted CDP tracing and full target topology are still operational tool gaps rather than finished reusable primitives.
- Lab booking is not real-environment verified.
- Course-selection execution is automated-test verified but lacks complete live acceptance evidence.
- Authentication and endpoint behavior can drift; discovery should propose a new version, never silently repair production facts.
