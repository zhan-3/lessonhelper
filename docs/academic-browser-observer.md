# Academic Browser Observer（实验性）

Academic Browser Observer 是项目内孵化的只读诊断能力。它连接用户明确指定的本机 CDP 浏览器，在外部操作者执行一次操作期间捕获脱敏的拓扑和网络增量，并生成诊断性的请求契约候选。

## 证据等级

- 核心连接、跨导航增量、脱敏、候选排序、复杂拓扑和 borrowed detach：**automated-test verified**。
- Pi Extension 加载：**automated-test verified**。
- Agent A/B 的 50% Token/人工干预价值门槛：**尚无有效结论**。一次本地 hidden eval 已因 fixture fidelity 仅为 3/14（21.43%）而作废；必须先达到 100% fixture fidelity 才能重跑并评价价值。
- 当前大学教务系统兼容性：**尚未 real-environment verified**。

测试结果不能证明真实 CAS、WebVPN 或教务页面当前兼容。

## 安全边界

Observer：

- 只接受明确提供的 loopback HTTP/WS CDP endpoint；
- 不扫描端口、不启动浏览器、不管理 profile；
- 不提供 click、fill、navigate、截图或任意 JavaScript 执行；
- borrowed 连接只能 detach，不能关闭远端浏览器；
- 只向模型返回有界、脱敏的目标摘要、增量计数和候选摘要；
- 不保存原始 HAR、HTML、请求/响应 body、Cookie、Token、学生记录或截图。

## 本地使用

1. 用户自行启动一个可见、启用 CDP 的浏览器，并确认它属于当前用户和 profile。
2. 启动本地工作台，例如 `start-workbench.cmd`。
3. 在 Pi 中确认项目级 Extension `.pi/extensions/academic-browser-observer.ts` 已加载。
4. 推荐使用两步组合接口，减少模型工具调用和 Token：
   - `academic_browser_begin`，传入明确 endpoint，例如 `http://127.0.0.1:9222/json/version`；
   - 用户或另一浏览器工具执行一次只读操作；
   - `academic_browser_finish`。

需要中途检查时，也可使用完整生命周期：`connect` → `inspect` → `start` → `checkpoint`/`stop` → `disconnect`。组合接口只合并这些安全步骤，不增加 click、navigate、JavaScript 或浏览器关闭能力。

`checkpoint`、`stop` 和 `finish` 返回紧凑的诊断候选，但候选不是已验证的生产读取契约，不能直接发布教务快照或启用写操作。

## 本地 eval

公开场景清单位于：

`.scratch/academic-browser-observer/evals/public-scenarios.json`

评估命令：

```powershell
uv run python -m course_selection.observer_evals `
  .scratch\academic-browser-observer\evals\public-scenarios.json `
  .private\observer-eval-runs.json
```

评估结果只应保存在 `.private/` 或已忽略的 `.scratch/academic-browser-observer/eval-results/`。未达到安全、完整 recall、50% Token/人工干预降低及时间改善门槛时，应继续使用现有 Chrome DevTools 或 Playwright 工具，而不是宣称 Observer 已证明独立价值。

## 真实只读 Shadow 验收

真实验收必须由用户授权一个具体只读操作，并遵循：

`docs/academic-browser-observer-shadow-acceptance.md`

没有该 dated witness 时，任何功能都不得标记为 `real-environment verified`。
