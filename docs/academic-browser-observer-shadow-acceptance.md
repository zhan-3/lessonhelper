# Academic Browser Observer: read-only Shadow acceptance

This procedure is **not** evidence that the live academic system has already been verified. Run it only after the synthetic eval gates pass and the user authorizes one named read-only operation.

## Boundary

- Mode: `observe`.
- Permitted observer actions: connect, inventory, start, checkpoint, stop, disconnect.
- The observer must not click, fill, navigate, evaluate JavaScript, submit, retry a mutation, or close the borrowed browser.
- Use a visible, user-established browser session.
- Keep DevTools evidence local. Record only redacted shape confirmations and counts.

## Procedure

1. Name the exact read-only operation and record explicit authorization.
2. Confirm the supplied CDP endpoint belongs to the visible borrowed browser.
3. Start observation before the user performs the operation.
4. The user performs exactly one read-only operation.
5. Stop the trace and inspect only its sanitized summary and candidates.
6. Independently compare target, redirect, and request shape in local DevTools. Do not copy raw values into chat or the repository.
7. Disconnect and confirm the browser, profile, session, and pre-existing tabs remain usable.
8. Create a local witness JSON using only fields accepted by `ShadowWitness`; use one safe operation code (`current_timetable_query`, `selection_list_query`, `grade_progress_query`, or `other_read_only_query`) and enum evidence codes rather than free text. Both witness and report must remain below `.private/`:

```powershell
uv run python -m course_selection.shadow_acceptance `.private\observer-shadow-witness.json `.private\observer-shadow-report.json
```

Any missing evidence, semantic contradiction, redaction uncertainty, failed browser-survival check, or absent authorization must produce `partial` or `failed`, never a pass.
