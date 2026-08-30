# Academic Browser Observer

**Status:** ready-for-agent

## Problem Statement

使用 AI 开发需要 CAS、WebVPN、SSO、动态菜单、多标签页、弹窗或多层 iframe 的教务及其他敏感后台系统时，开发者仍需反复打开 DevTools、提前开启 Preserve log、寻找正确页面和 frame、从大量网络请求中找出与一次操作相关的业务请求，再手工删除 Cookie、令牌、账号和个人数据后交给 Agent。这个过程耗时、消耗大量模型 Token，并且容易因日志开启过晚、目标页面判断错误、弹窗未真正导航、请求链不完整或脱敏遗漏而得到错误结论。

现有浏览器自动化工具擅长点击、填写、截图和通用调试，但没有为敏感认证系统提供一个紧凑的观察与理解 seam：在原始浏览器数据进入模型上下文之前完成拓扑识别、跨导航请求关联、不可逆脱敏、增量压缩和请求契约候选生成。结果是 Agent 对真实浏览器状态缺少可靠证据，开发者需要频繁人工救场。

项目需要一块可组合的完全自动化基础能力，而不是另一个通用浏览器控制器。它应观察一次由用户或其他自动化工具执行的操作，将操作前后的变化转换为脱敏、可验证、低 Token 的证据，并保持对 borrowed 浏览器的非破坏性连接语义。

## Solution

提供一个项目内孵化的 Pi Extension，由 harness-neutral 观察核心和薄 Pi Adapter 组成。第一版只连接用户显式指定的、已经运行的 CDP 浏览器，不启动、关闭、点击、填写、截图或执行任意页面 JavaScript。它在操作发生前开始观察，跨导航保留 Target、frame、popup、worker、请求、响应和重定向事件，在操作结束后生成 Observation Delta、脱敏证据摘要和 Request Contract Candidates。

原始敏感值只在进程内短暂存在，由核心模块在任何结果进入 Agent 上下文或持久化之前统一删除或替换为 Trace 内临时 HMAC 指纹。详细脱敏证据保存在本地 Evidence Store，通过证据 ID 和有界查询按需获取；默认工具结果只包含紧凑摘要。第一版不保存未脱敏原始证据。

MVP 以一个用户可理解的工作流验证价值：开始观察，用户执行一次只读操作，结束观察，插件自动指出与该操作相关的请求契约候选。模拟站点测试、历史事故回放、Agent A/B 和一次明确触发的真实系统只读 Shadow 验收共同证明它是否减少人工 DevTools 操作、Token 和时间。达不到价值门槛时停止独立插件开发，改用现有 Chrome DevTools 或 Playwright 工具。

## User Stories

1. As an AI-assisted academic-system developer, I want to begin observing before an interaction, so that navigation-time requests are not lost.
2. As an AI-assisted academic-system developer, I want to perform one normal browser operation without opening DevTools, so that interface discovery is less manual.
3. As an AI-assisted academic-system developer, I want the observer to finish after that operation, so that unrelated background activity does not dominate the evidence.
4. As an AI-assisted academic-system developer, I want the result to identify requests introduced by the observation window, so that I can focus on likely business contracts.
5. As an AI-assisted academic-system developer, I want request candidates ranked by temporal, target, initiator, schema, and provenance evidence, so that the most likely contract is visible first.
6. As an AI-assisted academic-system developer, I want redirects retained across navigation, so that authentication and application transitions remain explainable.
7. As an AI-assisted academic-system developer, I want page, frame, popup, OOPIF, worker, and service-worker topology inventoried, so that visible data is not missed because the wrong document was inspected.
8. As an AI-assisted academic-system developer, I want target identity based on capabilities and stable markers rather than list position, so that tab order changes do not corrupt observations.
9. As an AI-assisted academic-system developer, I want popup creation distinguished from committed navigation, so that a blank target is not reported as a successful application launch.
10. As an AI-assisted academic-system developer, I want partial target attachment reported explicitly, so that incomplete evidence is not presented as complete.
11. As a privacy-conscious developer, I want Cookie, Authorization, tokens, tickets, passwords, account identifiers, student identifiers, CSRF values, and session values removed before model exposure, so that private data stays local.
12. As a privacy-conscious developer, I want URL queries, headers, request bodies, JSON bodies, form fields, response summaries, and console output to share one redaction policy, so that sensitive values cannot escape through a weaker channel.
13. As a privacy-conscious developer, I want repeated sensitive values represented by a Trace-local fingerprint, so that the Agent can correlate requests without learning the original value.
14. As a privacy-conscious developer, I want fingerprints to be unlinkable across traces, so that observations cannot build a durable identity profile.
15. As a privacy-conscious developer, I want the MVP to avoid saving any unredacted evidence, so that temporary development artifacts do not become a new privacy risk.
16. As an Agent, I want a compact default summary, so that routine observations consume far fewer tokens than HAR, DOM, or accessibility-tree dumps.
17. As an Agent, I want an evidence ID for every observation, so that I can retrieve only the slice needed for the next decision.
18. As an Agent, I want checkpoint queries to return deltas rather than full history, so that repeated inspection does not resend unchanged evidence.
19. As an Agent, I want request shapes deduplicated locally, so that polling and repeated assets do not consume context.
20. As an Agent, I want schema fingerprints instead of full sensitive responses, so that I can compare candidate contracts safely.
21. As an Agent, I want transport, parse, domain, and completeness states kept distinct, so that HTTP success is not mistaken for a valid academic fact.
22. As an Agent, I want missing targets or events listed with a partial status, so that I know when to request another observation.
23. As an Agent, I want safe next actions returned with every non-success state, so that recovery remains bounded.
24. As a developer using an existing browser, I want to provide the CDP endpoint explicitly, so that the plugin never scans for or silently connects to another browser.
25. As a developer using an existing browser, I want the connection classified as borrowed, so that the observer cannot claim ownership of my process or profile.
26. As a developer using an existing browser, I want disconnect to leave the browser and all existing tabs alive, so that diagnostics cannot destroy my workbench session.
27. As a developer using an existing browser, I want only one active connection and one active Trace per Pi session, so that duplicate listeners cannot create contradictory evidence.
28. As a developer using an existing browser, I want repeated start, checkpoint, stop, and disconnect operations to be idempotent, so that retries do not multiply side effects.
29. As a developer using Pi, I want the Extension to map Pi cancellation into the observation lifecycle, so that Escape can stop local work without closing the browser.
30. As a developer using Pi, I want reload and shutdown to finish the active Trace and detach safely, so that hot iteration does not leave hidden listeners.
31. As a developer using Pi, I want only a脱敏 summary persisted into the Pi session, so that session history remains reviewable without containing raw browser evidence.
32. As a developer using Pi, I want the fieldwork Skill to explain when to invoke the observer, so that tool mechanics and decision policy remain separate.
33. As a developer using another Agent Harness in the future, I want the observation core independent from Pi, so that a DSH or MCP Adapter can reuse it without rewriting CDP, redaction, and evidence logic.
34. As a maintainer, I want the first Adapter to be Pi-only, so that current value can be proven before supporting additional Harnesses.
35. As a maintainer, I want the Extension to complement existing browser automation tools, so that it does not duplicate clicking, filling, navigation, screenshot, or arbitrary-evaluation capabilities.
36. As a maintainer, I want the external actor to be opaque to the observer, so that the same observation window works for a human, Playwright, Chrome DevTools MCP, or another Agent.
37. As a maintainer, I want a local synthetic site with nested same-origin and cross-origin frames, so that topology behavior is deterministic.
38. As a maintainer, I want fixture popups with delayed and absent navigation, so that target creation and navigation success are tested separately.
39. As a maintainer, I want fixture login redirects and navigation-bound requests, so that Preserve-log behavior is tested without a university system.
40. As a maintainer, I want fixture GET and POST read requests, so that observation is not incorrectly classified by method alone.
41. As a maintainer, I want fixture requests containing sensitive诱饵 in every supported channel, so that redaction failures go red reliably.
42. As a maintainer, I want borrowed disconnect tested against a live fixture browser, so that process survival is an external invariant.
43. As a maintainer, I want historical project failures represented as sanitized eval scenarios, so that the plugin is measured against real developer pain rather than implementation-derived tests.
44. As a maintainer, I want hidden variants of historical scenarios, so that the Skill cannot solve evals by memorizing exact fixture structure.
45. As a maintainer, I want an Agent A/B benchmark with and without the observer, so that reduced manual intervention, time, and tokens can be measured.
46. As a maintainer, I want local-only benchmark statistics, so that value measurement does not create telemetry or upload browsing data.
47. As a maintainer, I want one explicit real-system read-only Shadow acceptance, so that simulated correctness is not misrepresented as live compatibility.
48. As a maintainer, I want real-system acceptance to use only connect, observe, checkpoint, stop, and disconnect, so that MVP validation cannot change academic state.
49. As a product owner, I want sensitive-data leakage, browser closure, and false completeness to remain zero, so that speed improvements never weaken safety.
50. As a product owner, I want median manual interventions and token use reduced by at least half, so that the Extension demonstrates value beyond existing tools.
51. As a product owner, I want median completion time reduced materially, so that the observer improves the development loop rather than adding ceremony.
52. As a product owner, I want development stopped if the value thresholds are missed, so that the project uses existing browser tools instead of maintaining an indistinguishable plugin.
53. As a future automation developer, I want the observer to support an automated actor without core changes, so that it can become a composable fragment of bounded full automation.
54. As a future automation developer, I want CAPTCHA, unknown authentication pages, authorization prompts, ambiguous targets, incomplete traces, and exhausted budgets surfaced as human-required states, so that full automation knows when to stop.
55. As a future automation developer, I want state-changing requests excluded from this observer, so that a later execution capability can preserve explicit confirmation and single-submit safeguards.

## Implementation Decisions

- Build one deep observation module behind a single observation lifecycle seam. Callers and tests exercise connection, observation start, checkpoint, finish, report, and safe detach through that seam rather than testing internal CDP helpers directly.
- Keep the observation core independent from Pi and other Agent Harnesses. The core receives ordinary configuration, run identity, cancellation, evidence storage, and redaction policy through its interface.
- Implement a thin Pi Adapter that registers model-callable tools, maps the Pi session identity and cancellation signal, presents concise status, and persists only sanitized observation metadata.
- Incubate the first version as a project-local Pi Extension. Keep its internal module organization suitable for later migration into a package, but do not claim a stable public interface or open-source distribution yet.
- Use a CDP client designed for explicit attachment to an existing browser. Do not use the Extension to launch Chromium or manage a browser profile.
- Treat every connection as borrowed. The core exposes detach semantics only; no interface, configuration, or tool parameter may request remote browser shutdown.
- Require an explicit HTTP or WebSocket CDP endpoint. Provide a conventional local endpoint as documentation or configuration default, but do not scan ports or connect to the first discovered browser.
- Allow one borrowed connection and one active Trace per Pi session. Lifecycle operations are idempotent, and duplicate starts return or reference the existing state rather than adding listeners.
- Keep the external actor outside the observation interface. The actor may be a user or another browser automation tool; the observer measures the window before and after the action without taking control of that action.
- Make the MVP user flow begin observation, perform one external read-only operation, end observation, and receive contract candidates. Internal topology and tracing tools support this flow but are not the user-facing value proposition.
- Register explicit tools for connect, target inventory, Trace start, Trace checkpoint, Trace stop, and disconnect. Evidence query, evidence diff, layered session verification, and contract-candidate refinement may be introduced later or dynamically activated after the core workflow is proven.
- Do not expose click, fill, navigate, screenshot, arbitrary Runtime evaluation, browser-dialog handling, or page mutation tools in the MVP.
- Attach to relevant CDP Target types and retain parent-child topology for pages, frames, OOPIFs, popups, workers, and service workers. Identify candidates by safe capability markers rather than creation order or one hard-coded name.
- Start request capture before the external actor operates. Preserve request, response, redirect, initiator, loader, frame, and target relations across navigation until the observation ends.
- Produce an Observation Delta comparing the observation start with its checkpoint or end. Filter background noise using time, target, initiator, request shape, schema fingerprint, and repetition evidence without silently deleting uncertain candidates.
- Produce Request Contract Candidates containing redacted method/path shape, header and parameter names, dynamic-field indicators, response-schema fingerprint, business-identity indicators, pagination indicators, provenance, completeness, confidence evidence, and warnings.
- Treat request candidates as diagnostic evidence only. The observer does not promote a candidate into an application read contract, publish an academic snapshot, or modify LessonHelper persistence.
- Distinguish complete, partial, and failed observations. Missing required target sessions or event gaps may still yield diagnostics but cannot return a successful completeness state.
- Use one redaction implementation for URLs, headers, request and form bodies, JSON structures, response summaries, and console data. The implementation operates before data is returned to an Adapter or persisted.
- Redact common authentication, identity, student, CSRF, session, credential, ticket, code, and token fields case-insensitively, while allowing additional field and header names through trusted project configuration.
- Represent sensitive values with type, approximate length where safe, and a Trace-local HMAC fingerprint when correlation is useful. Generate a random key per Trace, never persist it, and destroy it when the Trace ends.
- Do not persist raw HAR, bodies, HTML, screenshots, Cookie values, authorization values, personal records, or unredacted console output in the MVP.
- Store detailed sanitized evidence outside model context and address it by Trace and checkpoint identity. Default tool content is a compact bounded summary; later queries return bounded slices or diffs.
- Persist only sanitized Trace metadata into the Pi session: identifiers, timestamps, completeness, event counts, warnings, and concise checkpoints. Reload and shutdown finish or invalidate active observation state and detach safely.
- Return a shared structured result envelope with typed status, concise summary, optional data, evidence reference, warnings, and safe next actions. Adapter-specific rendering does not alter core outcome semantics.
- Use event-driven waits and bounded recovery rather than fixed sleeps. Enforce explicit limits for runtime, navigation count, interaction window, authentication recovery, target count, candidate count, output bytes, and lines.
- Preserve the optimization priority order: safety, correctness, completeness, elapsed time, then model Token consumption.
- Keep the existing academic-system fieldwork Skill as the policy and decision layer. The Extension supplies observation evidence but does not replace the Skill's authentication, publication, execution, or acceptance rules.
- Treat existing Chrome DevTools MCP, Playwright MCP, and Pi browser-CDP tools as complementary actors or fallback solutions. Stop independent development if the observer cannot demonstrate its privacy, evidence, and efficiency distinction.
- Defer DSH and MCP Adapters until the core interface is stable, Pi validation passes, and a real second-Harness need exists.
- Defer license and public package decisions until privacy review, security tests, value benchmarks, and external-use evidence exist.

## Testing Decisions

- The primary automated test seam is the observation lifecycle interface exercised against a live local synthetic browser system. Tests cross the same seam as the Pi Adapter and assert final reports and browser-visible outcomes rather than internal CDP event-handler calls.
- Tests verify external behavior: complete topology, retained redirects, correct Observation Delta, ranked contract candidates, zero sensitive leakage, explicit partial status, idempotent lifecycle, and borrowed browser survival. They do not assert private helper structure, listener implementation, or exact internal event ordering unless observable correctness depends on it.
- Add only a small set of Pi Adapter contract tests. These verify tool registration, parameter mapping, cancellation propagation, bounded summaries, sanitized session metadata, and safe reload/shutdown. They do not duplicate core CDP tests.
- Build a synthetic site spanning multiple local origins. It includes a top-level portal, nested same-origin frame, cross-origin nested frame, OOPIF behavior, popup creation, delayed popup navigation, permanently blank popup, workers, and target detach.
- The synthetic site includes multiple authentication redirects, requests emitted around navigation, background polling noise, GET reads, POST reads, schema-similar decoys, pagination, and deliberately incomplete target/event conditions.
- Seed sensitive诱饵 into URL queries, fragments where observable, headers, cookies, form bodies, JSON keys and values, response structures, and console output. Tests search every serialized model-facing and persisted result for those exact values and must find none.
- Verify Trace-local correlation by asserting equal sensitive values receive equal fingerprints inside one Trace and unlinkable fingerprints across traces. Verify the per-Trace key is absent after completion and never serialized.
- Verify disconnect by querying the browser endpoint and existing page state after the Extension detaches. Browser process and tabs must remain alive; any remote close fails the test.
- Verify reload and shutdown with an active Trace. The result must be a typed incomplete/finished state, listeners must not duplicate on reload, and the browser must remain alive.
- Verify output bounds and incremental behavior. Checkpoints return only relevant deltas, repeated request shapes are summarized, and oversized evidence requires bounded follow-up queries rather than flooding model context.
- Derive public eval scenarios from sanitized historical failures: wrong tab, missing/nested frame, top-level content mistaken for iframe content, blank popup, authentication redirect mistaken for success, zero-row false completeness, late capture, partial pagination, duplicate observer ownership, and destructive borrowed disconnect.
- Maintain hidden variants that alter frame depth, origins, timing, names, target order, field names, and request methods while preserving the failure class. Expected outcomes are declared by fixture manifests before plugin execution, not generated from implementation behavior.
- Run Agent A/B evaluations on the same hidden tasks. The baseline uses the fieldwork Skill and ordinary browser/shell capabilities; the treatment adds the observer. Capture model Token use, elapsed time, tool calls, developer interventions, candidate accuracy, completeness errors, leaks, and browser disruption.
- Require zero sensitive诱饵 leakage, zero borrowed-browser closure, zero false-complete results, full expected target/request/redirect recall, at least a 50 percent reduction in median manual interventions, at least a 50 percent reduction in median model Token use, and a material completion-time reduction before claiming demonstrated value.
- Run one explicit real-system read-only Shadow acceptance only after synthetic tests pass. The observer connects to the user-established browser, begins observation, watches one user-performed read-only query, emits sanitized evidence, stops, and disconnects. It does not click, fill, navigate, or submit.
- Independently compare the real-system sanitized result with local human DevTools observation without copying raw evidence into the repository or model context. Record only dated invariants and residual risks.
- Classify implementation, automated-test verification, and real-environment verification separately. Synthetic tests and A/B benchmarks cannot promote current university compatibility.
- Existing project tests for academic-session targeting, diagnostic-only discovery, sanitization, complete-snapshot publication, exact execution identity, single-submit behavior, and unknown-result lockout are prior art for failure classification and safety semantics; the new Extension tests remain TypeScript-side and cross the new observation seam.

## Out of Scope

- General-purpose browser automation, including click, fill, navigation, screenshot-driven action, arbitrary JavaScript evaluation, and dialog control.
- Launching, installing, closing, or owning Chromium; managing user profiles; copying storage state; or recovering credentials.
- Automatic CAPTCHA, MFA, consent, or unknown authentication-page handling.
- Unattended course selection, lab booking, withdrawal, saving, submission, or any other state-changing academic operation.
- Discovering a write contract through a live mutation.
- Promoting diagnostic request candidates directly into fixed academic read contracts or current academic snapshots.
- Persisting raw HAR, response bodies, HTML, screenshots, cookies, tokens, credentials, student records, or unredacted logs.
- Replacing Chrome DevTools MCP, Playwright MCP, or other tools that already provide browser actions and general debugging.
- Implementing MCP, DSH, Claude Code, Codex, or other Harness Adapters in the MVP.
- Supporting multiple simultaneous CDP endpoints or multiple active traces per Pi session.
- Automatic local-port scanning, browser selection by first match, or remote CDP discovery.
- Publishing a stable public package, marketplace entry, license commitment, or compatibility promise.
- Real-environment state-changing acceptance.
- Guaranteeing future compatibility with changing Chrome, CDP, CAS, WebVPN, or university interfaces.

## Further Notes

- The observer is a fragment of bounded full automation: it supplies perception, evidence compression, and contract candidates. Existing browser actors perform actions; the fieldwork Skill supplies policy; project adapters and guarded execution capabilities remain separate.
- The core value proposition is not Target inventory or Network listing by themselves. It is converting one external operation into a privacy-preserving, low-Token, evidence-backed contract candidate without requiring the developer to operate DevTools.
- The first user is the maintainer of this repository. Potential later users are developers using Agents against authorized sensitive internal systems such as academic portals, enterprise administration, ERP, healthcare, government, finance, and legacy SSO applications. Ordinary browser users and simple automation scripts are not the target audience.
- Historical Session mining established the motivating failure classes, but raw Pi and Codex conversations remain local and must never become fixtures. Only sanitized behavioral scenarios and repository-verifiable decisions may enter tests or documentation.
- The Extension must know when to stop. Unknown authentication, CAPTCHA, authorization prompts, ambiguous targets, candidate writes, unstable identity, semantic contradiction, incomplete trace, and exhausted budgets are human-required or blocked outcomes rather than retry opportunities.
- The target long-term optimization is at least a 50 percent reduction in model Token use and manual interventions for equivalent interface-discovery work, without trading away privacy, correctness, completeness, or browser safety.
- If the MVP misses its privacy and value thresholds, maintainers should stop developing an independent observer and use existing Chrome DevTools MCP or Playwright tooling with the fieldwork Skill instead.
