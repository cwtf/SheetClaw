# SheetClaw Planning Suite Index

SheetClaw is an Excel task pane add-in (Office.js + React) providing an agentic chat interface
for reading and editing Excel workbooks across multiple LLM backends. It currently runs via local
shared-folder sideload, and **targets eventual publication to Microsoft Partner Center (AppSource)** —
see [Doc 15](15-appsource-readiness.md) for the readiness checklist. The architecture is BYOK
(bring-your-own-key) and backend-less, which is compatible with AppSource distribution.

## Documents

| # | Document | Purpose |
|---|----------|---------|
| 1 | [Architecture Overview](01-architecture-overview.md) | System diagram, tech choices, data flow, key decisions, browser-vs-sidecar split |
| 2 | [Component Specification](02-component-spec.md) | Responsibility, interface, state, deps, error states per component |
| 3 | [Data Models & Storage Schema](03-data-models-storage.md) | Every data structure + localStorage key schema |
| 4 | [Tool Schema Reference](04-tool-schema-reference.md) | Every agent tool, params, returns, Office.js mapping, serialization |
| 5 | [Auth & OAuth Flow Specification](05-auth-oauth-spec.md) | PKCE flow, Edge WebView constraints, API-key fallback, auth state machine |
| 6 | [Agentic Loop Specification](06-agentic-loop-spec.md) | Run state machine, context assembly, tool parsing, confirmation gate, errors, budget |
| 7 | [Usage Tracking Specification](07-usage-tracking-spec.md) | Usage record schema, token extraction, cost formula, aggregation, CSV export |
| 8 | [UI/UX Specification](08-ui-ux-spec.md) | Layout, states, interactions per surface |
| 9 | [Risk Register](09-risk-register.md) | Risk, likelihood, impact, mitigation, validation phase |
| 10 | [Build Sequencing & Dependency Map](10-build-sequencing.md) | Ordered build sequence, gates, references, complexity |
| 11 | [Web Access & Scope Clarification Spec](11-web-access-spec.md) | web_search/fetch_url tools, request_user_choice menu flow, provider config, caps, genericity rules (reattempt of `5a267ba`) |
| 12 | [OpenClaw Bridge Spec](12-openclaw-bridge-spec.md) | **Sunset pending** (superseded by Doc 13): delegate_web_task tool for a locally-running OpenClaw agent; retained as historical record |
| 13 | [Native Provider Search Spec](13-native-search-spec.md) | Two-tier search: native mechanism for capable LLM providers (OpenRouter, Anthropic, Kimi, Qwen, GLM), Doc 11 BYOK stack for the rest; tier-aware toggle gating; OpenClaw sunset plan |
| 14 | [Agent Speed Optimization Findings](14-agent-speed-optimization.md) | Findings + actionable backlog from a real slow session (`chat log.md`): broken web_search base-URL fallback, unbounded tool-payload resends, missing runtime/CORS guidance; prioritized fixes with file:line anchors |
| 15 | [AppSource / Partner Center Publish Readiness](15-appsource-readiness.md) | Certification checklist for publishing to Microsoft Partner Center (AppSource), mapped to current code: what's already ready, gaps (Office-on-the-web compat, listing assets, manifest validation, licensing), and per-item acceptance criteria |

## Decision log (items marked [DECISION REQUIRED] across the suite)

These require a human call before or during implementation. Each is flagged inline in the
relevant document; collected here for visibility.

| ID | Document | Decision |
|----|----------|----------|
| D1 | 1, 5 | Whether to run a Node.js sidecar at all (OAuth token exchange / CORS proxy) vs pure browser |
| D2 | 1 | State management library: Zustand vs Redux Toolkit vs React Context+reducer |
| D3 | 3, 7 | Where to store API keys: localStorage plaintext vs OS credential vault via sidecar — **Resolved (implemented): AES-GCM-256 encryption at rest via `src/auth/secureStore.ts`; OS-vault option deferred (needs a sidecar). Residual XSS-can-decrypt risk tracked in Doc 15.** |
| D4 | 5 | OpenAI consumer OAuth: confirm whether a public PKCE client is actually available; fallback to API key only |
| D5 | 4 | Pivot table scope: full create/modify vs read + limited create (Office.js maturity) |
| D6 | 6 | Confirmation granularity: per-tool-call vs batched per-turn write confirmation |
| D7 | 7 | Pricing table source: bundled static JSON vs fetched-and-cached remote |
| D8 | 3 | Multi-workbook handle persistence across Office add-in restarts (handles are not stable) |

## Conventions used throughout

- "Browser context" = the Edge WebView2 runtime hosting the task pane iframe.
- "Sidecar" = an optional local Node.js process (loopback HTTPS) for tasks the browser cannot do.
- All IDs are ULIDs unless stated otherwise (sortable, collision-resistant, no server needed).
- All timestamps are ISO-8601 UTC strings.
- "Provider" = an LLM backend; "adapter" = the code implementing the `LLMClient` interface for one provider family.
