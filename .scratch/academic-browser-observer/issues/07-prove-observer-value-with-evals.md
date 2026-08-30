# 07: Prove observer value with historical and hidden evals

**What to build:** Demonstrate that the observer solves the developer's actual problem rather than merely passing implementation-derived tests. Convert sanitized historical failure classes into deterministic public and hidden fixtures, then compare equivalent Agent workflows with and without the observer using local-only value metrics.

**Blocked by:** 04/Generate Request Contract Candidates from an Observation Delta; 05/Preserve observation across complex browser topology; 06/Bound the Pi observation lifecycle across reload and cancellation

**Status:** ready-for-agent

- [ ] Public scenarios cover wrong page, missing or nested frame, top-level-versus-frame confusion, blank popup, late capture, authentication redirect, false empty/complete result, duplicate ownership, and destructive detach.
- [ ] Hidden variants alter frame depth, origin, names, timing, target order, request method, and field names without changing the failure class.
- [ ] Expected outcomes come from fixture manifests created independently of plugin output and are not generated from implementation behavior.
- [ ] A repeatable Agent A/B harness measures model Token use, elapsed time, tool calls, developer interventions, target/request/redirect recall, candidate accuracy, completeness errors, sensitive leakage, and browser disruption.
- [ ] Sensitive诱饵 leakage, borrowed-browser closure, and false-complete outcomes are zero across the evaluation suite.
- [ ] Expected target, request, and redirect recall is complete for accepted observations; missing evidence produces partial rather than false success.
- [ ] Median manual interventions and model Token use improve by at least 50 percent relative to the baseline before demonstrated value is claimed.
- [ ] Results remain local and contain no real browser history, raw Session chat, university response, credential, or personal data.
- [ ] A failing value gate recommends using existing Chrome DevTools or Playwright tooling instead of presenting the observer as a distinct product.
