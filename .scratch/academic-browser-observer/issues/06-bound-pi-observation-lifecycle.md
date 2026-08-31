# 06: Bound the Pi observation lifecycle across reload and cancellation

**What to build:** Make observation dependable during real Pi development. Keep one borrowed connection and one active Trace per Pi session, propagate cancellation, enforce resource and output budgets, and detach safely on reload, shutdown, failure, or repeated calls while persisting only a concise sanitized summary.

**Blocked by:** 03/Keep sensitive browser evidence out of the model and disk

**Status:** automated-test verified (core lifecycle and Pi load; reload/budget matrix pending)

- [x] One Pi session cannot create multiple competing CDP connections or active traces through repeated tool calls.
- [x] Pi cancellation stops local observation work and returns a typed outcome without closing the borrowed browser.
- [x] Reload and shutdown finish or invalidate the active Trace, remove listeners, and detach from CDP idempotently.
- [x] Reload does not restore a stale live connection or duplicate listeners; later use requires a valid explicit connection state.
- [x] Runtime, target-count, candidate-count, output-byte, output-line, and recovery budgets end with inspectable typed outcomes.
- [x] Pi session persistence contains only sanitized Trace identity, timestamps, completeness, counts, warnings, and compact checkpoints.
- [x] Tool results use one shared structured status envelope with safe next actions for blocked, partial, invalid, cancelled, and failed states.
- [ ] Lifecycle tests verify browser and tab survival after cancellation, reload, shutdown, budget exhaustion, and duplicate calls. (Cancellation, shutdown, duplicate calls, and single-thread ownership are covered; live Pi reload and every budget class remain pending.)
