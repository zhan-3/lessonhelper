# 05: Preserve observation across complex browser topology

**What to build:** Make the same observation workflow reliable when the relevant operation occurs in nested same-origin or cross-origin frames, an OOPIF, popup, worker, or service worker. Report the complete safe topology and preserve target provenance without assuming that the newest page or one frame name is the business target.

**Blocked by:** 03/Keep sensitive browser evidence out of the model and disk

**Status:** ready-for-agent

- [ ] Target inventory retains parent-child relationships for pages, frames, OOPIFs, popups, workers, and service workers using sanitized descriptors.
- [ ] Target candidates are identified by safe capabilities and stable markers rather than list position, creation order, or one hard-coded frame name.
- [ ] Requests emitted by nested and cross-origin frames retain their originating target and frame provenance in the Observation Delta.
- [ ] Popup creation is distinct from committed nonblank navigation, and delayed navigation remains observable.
- [ ] A permanently blank popup is reported as an unresolved or failed navigation rather than a successful application target.
- [ ] Target attach, detach, navigation, and closure during an observation produce deterministic topology updates without crashing the Trace.
- [ ] Failure to attach or retain a required target marks the observation partial and lists the missing evidence.
- [ ] Synthetic tests vary origins, frame depth, names, timing, and target order while preserving the expected external result.
