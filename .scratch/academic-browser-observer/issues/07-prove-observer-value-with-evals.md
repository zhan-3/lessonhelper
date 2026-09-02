# 07: Prove observer value with historical and hidden evals

**What to build:** Demonstrate that the observer solves the developer's actual problem rather than merely passing implementation-derived tests. Convert sanitized historical failure classes into deterministic public and hidden fixtures, then compare equivalent Agent workflows with and without the observer using local-only value metrics.

**Blocked by:** 04/Generate Request Contract Candidates from an Observation Delta; 05/Preserve observation across complex browser topology; 06/Bound the Pi observation lifecycle across reload and cancellation

**Status:** blocked on hidden fixture fidelity (audit: 3/14 faithful, 21.43%); prior Agent A/B result invalidated

- [x] Public scenarios cover wrong page, missing or nested frame, top-level-versus-frame confusion, blank popup, late capture, authentication redirect, false empty/complete result, duplicate ownership, and destructive detach.
- [ ] Hidden variants alter frame depth, origin, names, timing, target order, request method, and field names without changing the failure class. (Manifests are independently curated, but only 3/14 executable fixtures currently match their declared semantics.)
- [x] Expected semantic outcomes were declared independently of plugin output; later scorer calibration normalized safe representation differences without changing the failure classes or required semantics.
- [x] A repeatable Agent A/B harness measures model Token use, elapsed time, tool calls, developer interventions, target/request/redirect recall, candidate accuracy, completeness errors, sensitive leakage, and browser disruption.
- [ ] Sensitive诱饵 leakage, borrowed-browser closure, and false-complete outcomes are zero across a fidelity-valid hidden suite. (The invalidated run observed zero safety events, but cannot satisfy the suite gate.)
- [ ] Expected target, request, and redirect recall is complete for accepted observations. (The hidden run found treatment recall of 0% targets, 57.14% requests, and 85.71% redirects despite all treatment rows reporting accepted.)
- [ ] Median manual interventions and model Token use improve by at least 50 percent relative to the baseline before demonstrated value is claimed. (Token use improved 39.61%; interventions remained 0 versus 0.)
- [x] Results remain local and contain no real browser history, raw Session chat, university response, credential, or personal data.
- [x] A failing value gate recommends using existing Chrome DevTools or Playwright tooling instead of presenting the observer as a distinct product.

## Invalidated A/B attempt

The local v7 run completed 14 baseline/treatment pairs, but an independent black-box fidelity audit subsequently found that only 3 of 14 executable hidden fixtures matched their predeclared semantics. Missing or contradictory behavior included requests, workers, redirects, popup behavior, pagination, timing, candidate signatures, and target depth.

Therefore the v7/v8 recall, candidate-accuracy, and efficiency numbers are calibration artifacts, not product evidence. The local reports remain Git-ignored under `.private/observer-eval-results/`, and the bounded audit is stored at `.private/observer-hidden-evals/fidelity-audit.json`.

The next gate is 100% fixture fidelity before another Agent A/B run. Until then, value remains unproven and existing Chrome DevTools or Playwright remains the recommended workflow.
