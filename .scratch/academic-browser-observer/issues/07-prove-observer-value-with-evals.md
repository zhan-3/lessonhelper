# 07: Prove observer value with historical and hidden evals

**What to build:** Demonstrate that the observer solves the developer's actual problem rather than merely passing implementation-derived tests. Convert sanitized historical failure classes into deterministic public and hidden fixtures, then compare equivalent Agent workflows with and without the observer using local-only value metrics.

**Blocked by:** 04/Generate Request Contract Candidates from an Observation Delta; 05/Preserve observation across complex browser topology; 06/Bound the Pi observation lifecycle across reload and cancellation

**Status:** hidden Agent A/B completed locally; value not demonstrated (safety passed, recall and 50% efficiency gates failed)

- [x] Public scenarios cover wrong page, missing or nested frame, top-level-versus-frame confusion, blank popup, late capture, authentication redirect, false empty/complete result, duplicate ownership, and destructive detach.
- [x] Hidden variants alter frame depth, origin, names, timing, target order, request method, and field names without changing the failure class. (The independently curated local suite remains under `.private/`.)
- [x] Expected semantic outcomes were declared independently of plugin output; later scorer calibration normalized safe representation differences without changing the failure classes or required semantics.
- [x] A repeatable Agent A/B harness measures model Token use, elapsed time, tool calls, developer interventions, target/request/redirect recall, candidate accuracy, completeness errors, sensitive leakage, and browser disruption.
- [x] Sensitive诱饵 leakage, borrowed-browser closure, and scorer-classified false-complete outcomes were zero across the completed hidden suite.
- [ ] Expected target, request, and redirect recall is complete for accepted observations. (The hidden run found treatment recall of 0% targets, 57.14% requests, and 85.71% redirects despite all treatment rows reporting accepted.)
- [ ] Median manual interventions and model Token use improve by at least 50 percent relative to the baseline before demonstrated value is claimed. (Token use improved 39.61%; interventions remained 0 versus 0.)
- [x] Results remain local and contain no real browser history, raw Session chat, university response, credential, or personal data.
- [x] A failing value gate recommends using existing Chrome DevTools or Playwright tooling instead of presenting the observer as a distinct product.

## Hidden A/B result

The local v7 run completed all 14 baseline/treatment pairs with no infrastructure-failed rows. Reports and scorer output remain Git-ignored under `.private/observer-eval-results/`.

- Safety: zero leakage, browser disruption, and scorer-classified false-complete outcomes.
- Baseline: 7 accepted and 7 incomplete; median 8,503.5 Tokens, 81.372 seconds, 1 tool call, and 0 interventions.
- Treatment: 14 accepted; median 5,135 Tokens, 29.6415 seconds, 2 tool calls, and 0 interventions.
- Efficiency: Tokens improved 39.61% and elapsed time improved 63.57%; tool calls increased and interventions did not improve.
- Treatment semantics: target recall 0%, request recall 57.14%, redirect recall 85.71%, and candidate accuracy 21.43%.
- Decision: `not_demonstrated`; continue using existing Chrome DevTools or Playwright for this diagnostic workflow.
