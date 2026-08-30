# Authentication state machine

## Separate layers

Track at least these states independently:

```text
browser: unavailable | ready | closed
network gateway: unknown | valid | invalid | waiting-for-human
academic destination: unknown | valid | invalid | stabilizing
```

A persistent profile or storage-state file is an input, not a state verdict.

## Protected verification

Verify the destination with the smallest authorized protected request that proves the intended identity/context. Reproduce request shape—required header names, same-origin context, and redirect handling—without copying captured secrets.

Authentication is valid only when:

1. the requested academic target, not an unrelated loopback/portal tab, is being checked;
2. the protected probe returns the expected authenticated schema/marker;
3. the result remains stable for the configured checks.

Treat HTTP login redirects, login HTML returned to XHR, ticket intermediates, and gateway-specific kick pages as typed states rather than generic failures.

## Human handoff

When CAPTCHA, consent, changed markup, or uncertain credentials appear:

1. activate the exact account-login page;
2. state the bounded human action without requesting secret values in chat;
3. set task state to `waiting_for_authentication` before blocking;
4. leave trace capture active but redact sensitive values;
5. resume only after protected verification succeeds;
6. map timeout, cancellation, and browser closure to explicit outcomes.

## Bounded recovery

Recovery attempts must be bounded and evidence-driven:

- retry ignored credential submission only when the page remains a recognized credential page;
- re-navigate a stalled ticket intermediate only within a small limit;
- tolerate a known one-shot post-login binding transition only within a short stabilization window;
- surface persistent gateway expiry as a typed terminal result.

Unknown login pages remain human-required. Do not expand credential autofill to an unverified origin or path.

## Persistence and reset

Persist authentication state only after authentication actually changed. Ordinary protected reads must not rewrite state or claim a new login.

Reset is destructive:

1. stop the owner browser;
2. remove authentication state/profile artifacts in the approved scope;
3. verify deletion;
4. block subsequent work if reset is incomplete.

Continuing with uncertain credentials is less safe than stopping.

## Suggested evals

- valid reused profile;
- unreadable state starts clean;
- ignored first credential submission;
- human CAPTCHA handoff;
- stalled ticket intermediate;
- transient post-login kick;
- protected probe returns login HTML;
- reset fails because a file is locked;
- requested academic page competes with a loopback workbench page.
