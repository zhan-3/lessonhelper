# Browser topology and tracing

## Topology before selectors

Inventory the whole browser surface:

```text
Browser/profile owner
└── context
    ├── page
    │   ├── frame
    │   │   └── child frame / OOPIF
    │   ├── worker
    │   └── popup
    └── service worker
```

For each page and frame retain only safe descriptors: origin class, redacted path pattern, title class, frame name, parent relation, load state, and stable DOM/API markers.

Search every viable page and frame before declaring an element absent. `page.frames` provides a flattened inventory; parent/child relationships retain the path. Revalidate the selected target immediately before use because frames detach and navigate.

A target is accepted by capability, not position:

```text
protected origin/path class
+ stable marker
+ expected read capability
+ semantically valid protected response
```

`pages[-1]`, “first matching frame,” and one hard-coded iframe name are not identities.

## Persistent evidence

Start capture before navigation or the gesture under investigation. Retain:

- request ID, loader ID, frame/target ID;
- method and redacted path pattern;
- redirect chain;
- initiator class;
- header names, never authorization/cookie values;
- parameter/form field names, with sensitive values removed;
- status, content type, and response-schema fingerprint;
- console exceptions and target lifecycle events.

Append events to an agent-owned trace across navigation rather than relying on DevTools UI state. DevTools Preserve log is the manual fallback, not the preferred automation contract.

## Popup rule

A popup event proves target creation only. Accept application launch after:

1. a nonblank navigation commits within a deadline;
2. the destination passes its protected-page marker/probe;
3. abandoned blank targets are closed only when owned by this task.

When a portal's `window.open()` creates a permanently blank target, capture the intended destination from the authorized trigger and navigate a controlled owned page directly. Do not guess the destination.

## Ownership rule

Every attachment declares:

- `owned`: this task created the browser/context and may close it;
- `borrowed`: another process owns it; this task may detach only.

Before diagnosing navigation storms or contradictory pages, inventory process, profile, lock, port, and supervisor ownership. Establish exactly one browser/profile owner.

## Missing tool primitives

A reusable browser extension should eventually expose:

```text
inspect_browser_targets
start_persistent_trace
checkpoint_trace
record_controlled_interaction
find_target_by_capability
attach_cdp_owned
attach_cdp_borrowed
assert_single_browser_owner
```

Until those primitives exist, report the manual step and its evidence instead of pretending the observation was automated.
