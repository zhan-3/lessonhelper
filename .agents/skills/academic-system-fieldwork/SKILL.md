---
name: academic-system-fieldwork
description: Develop and diagnose authenticated academic or campus-system browser automation. Use when exploring CAS, WebVPN, or SSO flows; finding requests behind dynamic pages, iframes, or popups; converting observations into fixed read contracts; handling human authentication; or implementing guarded course-selection and booking writes.
compatibility: Requires an authorized local account and browser automation or DevTools access. State-changing work requires explicit user confirmation.
---

# Academic System Fieldwork

Treat the live system as fieldwork: observations are evidence, not facts, until a repeatable contract passes its acceptance gates.

## 1. Declare the mode

Choose exactly one mode before opening the remote system:

- `observe`: navigation and read-only experiments; block suspected writes.
- `execute-once`: one explicitly identified mutation authorized by the user.

Keep credentials, cookies, tokens, raw HTML/HAR, student records, and browser profiles local. Persist only redacted request shape, schema, provenance, and typed outcomes. Stop if authorization or mode is unclear.

**Complete when:** the mode, allowed methods/actions, sensitive-data boundary, and human checkpoints are explicit.

## 2. Establish browser ownership and evidence capture

Identify the browser binary, process/profile owner, CDP attachment mode (`owned` or `borrowed`), pages, frames, popups, and workers. A borrowed CDP connection may disconnect but must not close the browser.

Arm network and console capture before the first navigation or controlled gesture, preserving redirects and navigation boundaries. If tooling cannot retain evidence, ask the user to enable DevTools Preserve log before the gesture and export locally for redaction.

Read [browser-topology-and-tracing.md](references/browser-topology-and-tracing.md) when pages, iframes, popups, dynamic menus, or missing requests are involved.

**Complete when:** one owner is established, target topology is inventoried, and capture is demonstrably active before interaction.

## 3. Establish layered authentication

Model browser availability, gateway/VPN authentication, and destination-system authentication as separate states. A URL change, portal shell, persisted profile, or storage-state file is not proof.

Use a visible browser for human CAPTCHA, consent, or changed login flows. Verify success with a protected destination request using the observed request shape and a bounded stability check.

Read [authentication-state-machine.md](references/authentication-state-machine.md) for recovery and human handoff rules.

**Complete when:** the intended protected target passes its business/API probe repeatedly, or the task ends in a typed human-required/invalid state.

## 4. Run one controlled interaction experiment

Capture the before-state, perform one read-only UI action, then correlate new requests by time, initiator/frame, method, redacted path pattern, status, content type, and response-schema fingerprint. UI clicks are experiments; screenshots and labels are supporting evidence only.

In `observe` mode, abort or block requests whose mutation risk is not proven safe. Never use a live write merely to discover an interface.

**Complete when:** every candidate request is tied to an observed action and classified as authentication, navigation, read, possible write, confirmed write, or unknown.

## 5. Promote a fixed read contract

A candidate becomes a read contract only when it has:

- stable method and redacted path pattern;
- required header names and parameter names without secret values;
- explicit term/category/page provenance;
- a stable business identity field;
- redirect/login/error classification;
- schema and semantic acceptance checks;
- repeatable results across at least two controlled reads.

Read [contracts-and-publication.md](references/contracts-and-publication.md) before implementing collection or publishing a snapshot.

**Complete when:** transport, parse, and domain acceptance all pass; otherwise retain it as diagnostic evidence only.

## 6. Publish only complete facts

Build an expected-segment manifest for every required term, category, and page. Publish atomically only after all required segments pass. A missing segment is unknown, not zero. Preserve the last complete snapshot and record failed attempts separately.

**Complete when:** expected and completed segments match exactly and the result passes semantic plausibility checks.

## 7. Cross the write boundary separately

Load [safe-execution.md](references/safe-execution.md) before any enrollment, booking, withdrawal, save, or submission action.

Require an exact server-issued target identity, a fresh source-page re-read, fresh dynamic fields, and explicit user confirmation naming the target. Send at most once. After transmission, an ambiguous result becomes `possibly_applied` and blocks retry until a human or authoritative read resolves it.

**Complete when:** the action is explicitly confirmed and ends as confirmed success, confirmed failure, or quarantined unknown—with no automatic retry.

## 8. Report evidence, not confidence

Classify each feature independently as:

- `implemented`
- `automated-test verified`
- `real-environment verified`

Read [acceptance-and-evidence.md](references/acceptance-and-evidence.md) before claiming live compatibility. Record dated evidence and unresolved live risks.

**Complete when:** every claim names its evidence tier, manual interventions, residual risks, and next safe verification step.

## Maintaining this skill

Read [origin-case-study.md](references/origin-case-study.md) before changing these rules. Preserve only lessons that changed behavior and remain supported by runtime, code, test, or dated live evidence.
