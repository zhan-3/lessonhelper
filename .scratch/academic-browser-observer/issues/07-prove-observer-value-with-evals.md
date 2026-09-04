# 07: Prove observer value with historical and hidden evals

**What to build:** Demonstrate that the observer solves the developer's actual problem rather than merely passing implementation-derived tests. Convert sanitized historical failure classes into deterministic public and hidden fixtures, then compare equivalent Agent workflows with and without the observer using local-only value metrics.

**Blocked by:** 04/Generate Request Contract Candidates from an Observation Delta; 05/Preserve observation across complex browser topology; 06/Bound the Pi observation lifecycle across reload and cancellation

**Status:** fidelity-valid hidden Agent A/B completed; safety and efficiency passed, semantic value not demonstrated

- [x] Public scenarios cover wrong page, missing or nested frame, top-level-versus-frame confusion, blank popup, late capture, authentication redirect, false empty/complete result, duplicate ownership, and destructive detach.
- [x] Hidden variants alter frame depth, origin, names, timing, target order, request method, and field names without changing the failure class. (An independent black-box audit verified 14/14 executable fixtures with two fresh runs per task.)
- [x] Expected semantic outcomes were declared independently of plugin output; later scorer calibration normalized safe representation differences without changing the failure classes or required semantics.
- [x] A repeatable Agent A/B harness measures model Token use, elapsed time, tool calls, developer interventions, target/request/redirect recall, candidate accuracy, completeness errors, sensitive leakage, and browser disruption.
- [x] Sensitive诱饵 leakage, borrowed-browser closure, and scorer-classified false-complete outcomes are zero across the fidelity-valid hidden suite.
- [ ] Expected target, request, and redirect recall is complete for accepted observations. (The hidden run found treatment recall of 0% targets, 57.14% requests, and 85.71% redirects despite all treatment rows reporting accepted.)
- [x] Median manual interventions and model Token use improve by at least 50 percent relative to the baseline. (Tokens fell from 11,157 to 5,513.5, a 50.58% reduction; interventions fell from 1 to 0.)
- [x] Results remain local and contain no real browser history, raw Session chat, university response, credential, or personal data.
- [x] A failing value gate recommends using existing Chrome DevTools or Playwright tooling instead of presenting the observer as a distinct product.

## Fidelity-valid hidden A/B result

After invalidating the earlier calibration run, the synthetic fixture runner was corrected and independently audited through its public start/trigger/status/stop interface. The final audit passed 14/14 tasks with two fresh runs per task, actual second-trigger idempotency checks, browser/tab survival, and cleanup verification. The final v9 A/B then ran all 14 baseline/treatment pairs from scratch.

- Safety: 14/14 treatment rows passed with zero leakage, browser disruption, and scorer-classified false-complete outcomes.
- Efficiency: median Tokens fell from 11,157 to 5,513.5 (50.58%); elapsed time fell from 92.877 to 24.6335 seconds (73.47%); median interventions fell from 1 to 0; tool calls rose from 1 to 2.
- Treatment completion: 14/14 Agent rows reported accepted, versus 1/14 baseline.
- Treatment semantics: target recall 0%, request recall 92.86%, redirect recall 100%, and candidate accuracy 0%.
- Decision: `not_demonstrated`. Safety and efficiency are now supported, but target semantics and candidate ranking fail the required correctness gate. Continue using existing Chrome DevTools or Playwright until those two defects are corrected and the hidden suite is rerun.

All reports remain local and Git-ignored under `.private/observer-eval-results/`; the independent fixture audit is `.private/observer-hidden-evals/fidelity-review-v4.json`.
