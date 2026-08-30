# Safe execution

State-changing operations include enrollment, booking, withdrawal, save, submit, confirmation, and any request whose effect is unknown.

## Capability boundary

Discovery and execution use different commands, tools, and policies:

- observation allows verified reads and blocks unknown/write requests;
- execution exposes one narrowly scoped operation after explicit authorization.

A generic browser explorer should not possess an unrestricted mutation capability.

## Preflight

Immediately before sending:

1. identify the exact server-issued target identity;
2. re-read the target's authoritative source category/page;
3. require exactly one matching executable target;
4. verify term/window/profile/context and applicable conflict rules;
5. acquire fresh dynamic fields and tokens from the live page;
6. display a redacted action summary naming the target;
7. bind user confirmation to a fingerprint/nonce for that exact action.

A stale snapshot, display-only fallback identity, multiple matches, missing dynamic field, or changed context stops execution.

## Point of no return

The operation has two phases:

```text
pre-send: cancellable, fail closed
post-send: not safely cancellable, never infer non-application
```

Send at most one request. Do not retry transport timeout, connection loss, browser closure, unknown alert text, or unclassified response.

## Outcome classification

- `confirmed_success`: recognized response plus authoritative follow-up evidence where available;
- `confirmed_failure`: recognized rejection that proves no application;
- `possibly_applied`: request may have reached the server but outcome is not authoritative.

Persist a redacted durable ledger containing target fingerprint, send time, outcome class, and verification status. `possibly_applied` blocks the same target until authoritative read or human verification resolves it.

## Live acceptance

A mutation is not real-environment verified by mocks, selectors, or a captured historical request. Live acceptance requires separate authorization, a low-risk target, pre-armed redacted evidence capture, exactly one action, and manual/result verification.

## Suggested evals

- fallback/display identity rejected;
- stale snapshot rejected;
- multiple target identities rejected;
- fresh re-read no longer contains target;
- confirmation fingerprint mismatch;
- exactly one outbound mutation;
- timeout after send enters `possibly_applied`;
- process restart preserves ambiguity lockout;
- unknown server message does not retry;
- resolved authoritative follow-up clears lockout.
