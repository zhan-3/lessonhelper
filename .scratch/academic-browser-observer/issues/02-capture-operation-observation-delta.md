# 02: Capture one external operation as a cross-navigation Observation Delta

**What to build:** Let a developer start an observation, perform one external browser operation, checkpoint or stop the observation, and receive the requests, responses, target changes, and redirect chain introduced during that window. The actor remains outside the observer, so the same flow works for a human or another browser automation tool.

**Blocked by:** 01/Safely connect to and inspect a borrowed browser

**Status:** ready-for-agent

- [ ] Observation starts before the actor operates and captures navigation-bound requests that would be lost by late inspection.
- [ ] Requests, responses, redirects, loader identity, frame identity, target identity, and initiator class remain correlated across navigation.
- [ ] Checkpoint returns only the delta since observation start or the preceding checkpoint rather than replaying unchanged history.
- [ ] Stop returns a final typed Observation Delta and releases active Trace listeners without detaching the borrowed browser connection.
- [ ] Duplicate start, checkpoint, and stop calls are idempotent and do not multiply listeners or events.
- [ ] A local fixture demonstrates a multi-hop redirect and requests emitted immediately before and after navigation.
- [ ] Output is bounded and reports complete, partial, or failed status instead of treating collected-so-far events as automatically complete.
