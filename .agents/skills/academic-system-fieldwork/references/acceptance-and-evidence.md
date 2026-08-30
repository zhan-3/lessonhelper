# Acceptance and evidence

## Evidence tiers

Track each capability independently:

| Tier | Required evidence | Does not prove |
| --- | --- | --- |
| `implemented` | current code path exists and is reviewed | tests or live compatibility |
| `automated-test verified` | named automated tests exercise success and failure paths | current university-system compatibility |
| `real-environment verified` | dated, authorized live witness satisfies defined invariants | future compatibility after system drift |

Never collapse these tiers into “done.”

## Live read invariants

A read-only acceptance should record, without secrets:

- browser provenance and single profile owner;
- reused-profile, expired-profile, and human-login behavior;
- page/frame target contract;
- protected authentication probe;
- request-contract version;
- nonempty result when a known witness is populated, or evidence that empty is plausible;
- term/category/page completeness;
- redaction and persistence boundaries;
- cleanup and borrowed-CDP detach behavior.

## Live write invariants

A write acceptance additionally requires:

- explicit authorization for one low-risk target;
- exact server-issued identity;
- fresh preflight and dynamic fields;
- pre-armed evidence capture;
- one outbound mutation maximum;
- recognized outcome or unknown quarantine;
- authoritative/manual follow-up verification.

## Fieldwork record

For every hard-won lesson capture:

```yaml
problem: observable failure
failed_assumption: what the agent treated as true
intervention: human action or decisive runtime observation
resolution: what changed
reusable_rule: behavior another agent can follow
tool_gap: executable primitive that would remove manual work
evidence:
  tier: implemented | automated-test verified | real-environment verified
  pointer: sanitized session/commit/test reference
residual_risk: what remains unknown
```

A lesson without a failed assumption or changed behavior is usually generic advice, not mined experience.

## Release report

Report:

1. capability and evidence tier;
2. dated live witnesses, if any;
3. manual intervention count and reason;
4. known contract drift risks;
5. privacy review status;
6. unresolved `possibly_applied` actions;
7. next safe verification step.

Green tests must not upgrade a live-system claim.

## Tool-gap backlog derived from fieldwork

Prioritize tools that remove repeated human rescue:

1. persistent redacted trace across navigation;
2. full page/frame/popup/worker topology inventory;
3. capability-based target resolver;
4. layered protected-session verifier;
5. controlled-interaction request correlator;
6. request-shape differ and contract generator;
7. completeness barrier and atomic publisher;
8. owned/borrowed CDP lifecycle guard;
9. one-shot mutation ledger and ambiguity lockout;
10. verification-matrix reporter.
