# 01: Safely connect to and inspect a borrowed browser

**What to build:** Provide the first end-to-end observer slice: a Pi user explicitly connects to an already running CDP browser, receives a compact sanitized inventory of its basic targets, and disconnects without closing or changing the browser or its existing tabs. Establish the harness-neutral observation seam, the thin Pi Adapter, and the minimal synthetic browser system needed to verify this behavior.

**Blocked by:** None (can start immediately)

**Status:** ready-for-agent

- [ ] An explicit HTTP or WebSocket CDP endpoint connects successfully; the Extension never scans local ports or silently selects another browser.
- [ ] Connection state is classified as borrowed and exposes detach semantics without any remote-browser close option.
- [ ] The model-callable interface can connect, inspect basic page targets, and disconnect through the same observation seam used by tests.
- [ ] The returned inventory is bounded and omits full URLs, page content, credentials, cookies, and personal values.
- [ ] Disconnect leaves the fixture browser process and its pre-existing tabs alive and usable.
- [ ] Repeated connect, inspect, and disconnect calls return typed, deterministic outcomes rather than creating duplicate ownership.
- [ ] Automated tests cross the public observation seam and verify browser-visible outcomes rather than private CDP helpers.
