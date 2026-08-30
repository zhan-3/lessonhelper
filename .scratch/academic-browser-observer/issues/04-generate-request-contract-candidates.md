# 04: Generate Request Contract Candidates from an Observation Delta

**What to build:** Turn one privacy-safe Observation Delta into a small ranked set of request contracts that an Agent can evaluate without reading a HAR. Remove obvious background noise, aggregate repeated request shapes, preserve uncertain candidates, and explain why each candidate may represent the actor's business operation.

**Blocked by:** 03/Keep sensitive browser evidence out of the model and disk

**Status:** ready-for-agent

- [ ] Candidate generation correlates requests with the observation window using time, target, frame, initiator, redirects, schema, repetition, and provenance evidence.
- [ ] Repeated polling and equivalent request shapes are summarized without hiding distinct business requests.
- [ ] Each candidate includes sanitized method and path shape, header and parameter names, dynamic-field indicators, response-schema fingerprint, identity indicators, pagination indicators, and completeness warnings.
- [ ] Candidate ranking includes machine-readable reasons and never claims that confidence alone verifies a production contract.
- [ ] POST reads remain eligible and GET requests are not assumed safe solely from their method.
- [ ] Default model-facing output is a compact ranked summary with an evidence identity; detailed sanitized evidence is retrieved only in bounded slices or diffs.
- [ ] Decoy requests with similar names or schemas do not outrank the request causally tied to the fixture operation.
- [ ] Candidate output remains diagnostic-only and cannot publish an academic snapshot or enable a state-changing request.
