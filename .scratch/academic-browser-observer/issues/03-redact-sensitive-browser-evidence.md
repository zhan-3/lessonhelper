# 03: Keep sensitive browser evidence out of the model and disk

**What to build:** Make the complete observation flow safe for a sensitive authenticated application. Before any evidence reaches a Pi tool result, Pi session entry, diagnostic summary, or persistent store, apply one redaction policy across every supported channel and retain only safe structure plus optional Trace-local correlation fingerprints.

**Blocked by:** 02/Capture one external operation as a cross-navigation Observation Delta

**Status:** ready-for-agent

- [ ] URL query values, headers, cookies, request bodies, forms, JSON keys and values, response summaries, and console output pass through the same redaction module.
- [ ] Common credential, authentication, session, CSRF, ticket, account, identity, and student fields are recognized case-insensitively, with trusted project configuration for additional names.
- [ ] Equal sensitive values receive an equal HMAC fingerprint within one Trace without revealing the original value.
- [ ] The same value receives an unlinkable fingerprint in a different Trace, and the per-Trace key is never serialized or retained after completion.
- [ ] The MVP writes no unredacted HAR, body, HTML, screenshot, cookie, token, personal record, or console data to disk.
- [ ] Exact sensitive诱饵 placed in every supported fixture channel are absent from all model-facing and persisted output.
- [ ] Sanitized evidence remains useful for request-shape comparison by retaining safe names, types, lengths where appropriate, schema fingerprints, and provenance.
- [ ] Redaction failure blocks evidence publication instead of returning a partially sanitized success.
