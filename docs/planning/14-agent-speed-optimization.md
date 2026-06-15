# 14 — Agent Speed Optimization Findings

> **Status:** Findings + actionable backlog. Nothing here is implemented yet.
> **Audience:** Any LLM/engineer picking this up cold. Each task below is self-contained with file:line anchors, concrete steps, and acceptance criteria.
> **Source of evidence:** A real session captured in `chat log.md` at the repo root (deepseek-v4-pro, populating GDP-per-capita-PPP data for ~19 countries).

---

## TL;DR

A task that needed **~8–10 productive tool calls** instead consumed **38 iterations and ~2.98M tokens**, hit the 25-iteration cap, and had to be continued. The slowness is **not** the model's raw throughput. It is two compounding harness problems:

1. **Discovery was broken** — every `web_search` returned the local dev server's HTML instead of results, so the model fell back to ~30 blind `fetch_url` guesses.
2. **Context grew unbounded** — every fetched payload was re-sent to the model on every later iteration, so input size (and per-call latency) grew turn over turn.

Fix the three "Highest impact" tasks below and the bulk of the blowup disappears.

---

## Architecture constraints (read before changing anything)

- SheetClaw is a **browser-only Office Add-in** (taskpane webview). There is **no backend/proxy** owned by the project; it ships as a static deployment. See `docs/planning/11-web-access-spec.md`.
- Therefore **CORS is fundamental**: a direct `fetch()` from the taskpane only succeeds for hosts that send permissive CORS headers. Many open-data APIs do; most websites don't.
- `fetch_url` mitigates this with an **automatic reader-proxy fallback** (`r.jina.ai`) on network error — see [src/web/fetch.ts:64](../../src/web/fetch.ts). The model **cannot** control or bypass this fallback.
- CORS-blocked hosts are **cached at runtime** so repeat calls fail fast — see [src/web/net.ts:73](../../src/web/net.ts).
- Web search/fetch are read-only tools with `runtime: 'none'`, so the loop is allowed to run consecutive ones **concurrently** — see [src/agent/loop.ts:277](../../src/agent/loop.ts).

---

## Evidence from `chat log.md`

| Symptom | Log lines | What it proves |
|---|---|---|
| `web_search` returned `<!DOCTYPE html> … /@react-refresh` | 9, 18, 103 | The Tavily request resolved to the **Vite dev server**, not `api.tavily.com`. HTTP 200 + HTML rules out a CORS block (that throws `TypeError`). |
| Model "trying without the Jina proxy" repeatedly | 31, 40, 46, 55 | Wasted ~5 iterations on something it cannot control (`fetch_url` always falls back to the reader). |
| Repeated truncated fetches of the IMF "all countries" endpoint | 14, 29, 41, 65, 68 | Model needed ~19 of ~190 countries but had no way to request a slice of a large JSON. |
| `deepseek-v4-pro | iter 38/50 | 2984526 tok` | header | ~3M tokens for an ~8-call task; transcript re-billed every turn. |
| "Stopped after 25 iterations… Continuing for 25 more." | 82–83 | Hit `MAX_ITERATIONS` and had to be manually extended. |

---

## Findings & backlog

Tasks are ordered by impact. Each can be done independently.

### ★ Task 1 — Fix `web_search` resolving to the dev server (HIGHEST)

**Problem.** Search results came back as the dev server's `index.html`. The base URL flows
`appConfig.webAccess.baseUrl` → [src/taskpane/workbookLayer.ts:21](../../src/taskpane/workbookLayer.ts) → provider `opts.baseUrl`.
In [src/web/providers/tavily.ts:25](../../src/web/providers/tavily.ts) the code does `opts.baseUrl ?? this.endpoint`. The `??` operator only falls back on `null`/`undefined`, so an **empty string or relative/same-origin value passes through**, and `fetch('')`/a relative path hits the local origin and returns the SPA fallback page.

**Fix.**
1. Add a shared guard that returns `this.endpoint` unless `baseUrl` is a **valid absolute http(s) URL**. Apply it in every provider that reads `opts.baseUrl`:
   - [src/web/providers/tavily.ts:25](../../src/web/providers/tavily.ts)
   - [src/web/providers/searxng.ts:22](../../src/web/providers/searxng.ts)
   - [src/web/providers/google-cse.ts:27](../../src/web/providers/google-cse.ts)
   - [src/web/providers/jina.ts:22](../../src/web/providers/jina.ts)
   - [src/web/providers/wikipedia.ts:31](../../src/web/providers/wikipedia.ts)
2. When a search response has a non-JSON `content-type` (e.g. `text/html`), throw a clear provider error ("web-access base URL looks misconfigured — it returned HTML, not JSON") instead of surfacing a 180-char HTML preview the model can't act on. Current preview logic: [src/web/providers/tavily.ts:55](../../src/web/providers/tavily.ts).

**Acceptance criteria.**
- With `webAccess.baseUrl` set to `''`, a relative path, or a same-origin URL, the provider uses the real endpoint instead.
- An HTML response yields an actionable config error, not a JSON-parse error with an HTML snippet.
- Add/extend tests in `src/web/__tests__/search.test.ts`.

---

### ★ Task 2 — Compact stale tool payloads proactively (HIGHEST)

**Problem.** `compact()` only degrades old results **when the transcript exceeds the context budget** — see [src/agent/context-builder.ts:72](../../src/agent/context-builder.ts). The session ran on a large-context model, so `fits(history)` stayed true and **all 38 fetch payloads (each up to ~24k chars ≈ 6k tokens) were resent in full every turn**. That is the ~3M-token figure: the same growing transcript re-billed ~38 times.

**Fix.** Make compaction **age-based and unconditional**, not budget-gated:
- Once a `tool` result is older than N turns (start with N = 2–3), collapse it to a short summary/handle, e.g. `fetch_url https://…/PPPPC → JSON, ~190 countries; kept: CHN,JPN,KOR,…`.
- Keep the most recent 1–2 tool results at full size (the `recent`/`squash` machinery at [src/agent/context-builder.ts:87](../../src/agent/context-builder.ts) already distinguishes recent vs old — reuse it, just run it always).
- `MAX_TOOL_RESULT_CHARS` is defined at [src/agent/context-builder.ts:20](../../src/agent/context-builder.ts).

**Acceptance criteria.**
- For a synthetic 30-iteration transcript with large tool results, estimated input tokens stay roughly flat across iterations instead of growing linearly.
- The two most recent tool results are preserved verbatim (don't force a refetch of data the model is acting on now).
- Add/extend tests in `src/agent/__tests__/context-builder.test.ts`.

---

### ★ Task 3 — Tell the model the truth about its runtime (HIGHEST, cheap)

**Problem.** The [system prompt](../../src/agent/system-prompt.ts) never says the agent runs in a browser, that direct fetches are CORS-limited, or that the reader fallback is automatic. So the model wasted iterations "trying without the Jina proxy" (log 31/40/46/55), which it cannot do.

**Fix.** Add ~3 lines to [src/agent/system-prompt.ts](../../src/agent/system-prompt.ts):
- "You run in a browser; many hosts block direct fetches (CORS). `fetch_url` automatically retries through a reader proxy — you do **not** control or bypass this. Never retry a URL just to 'avoid the proxy.'"
- "If a host is reported CORS-blocked, or a preview is truncated, switch sources or call `request_user_choice` — do not retry the same host." (Runtime already caches CORS-blocked hosts: [src/web/net.ts:73](../../src/web/net.ts).)

**Acceptance criteria.**
- Prompt explicitly describes the browser/CORS/reader-fallback reality.
- Existing prompt tests (`src/agent/__tests__/`) still pass; update snapshot/assertions if any.

---

### Task 4 — Add field/JSON-path extraction to `fetch_url` (MEDIUM)

**Problem.** The model hit the IMF "all countries" endpoint repeatedly and got truncated (log 14/29/41/65/68). It needed ~19 of ~190 countries but had no way to slice a large JSON. `formatJsonPayload` ([src/web/fetch.ts:113](../../src/web/fetch.ts)) only truncates and advises "refetch with a narrower API query" — impossible for one fixed endpoint.

**Fix.** Add an optional `json_path` / `select` parameter to the `fetch_url` tool spec ([src/web/fetch.ts:16](../../src/web/fetch.ts)) so the model can extract e.g. `values.PPPPC.CHN` or a key-filtered subset before the payload is capped.

**Acceptance criteria.**
- A large JSON body filtered by `json_path` returns only the selected slice, under the char cap.
- Invalid paths return a clear validation error.
- Tests in `src/web/__tests__/fetch.test.ts`.

---

### Task 5 — Enable prompt caching on the Anthropic path (MEDIUM)

**Problem.** [src/adapters/anthropic.ts](../../src/adapters/anthropic.ts) sets no `cache_control` breakpoints, so every turn reprocesses the full system prompt + tool schemas + manifest + history at full latency/cost. The OpenAI/DeepSeek path already benefits from automatic server-side caching and reads it back ([src/adapters/openai.ts:257](../../src/adapters/openai.ts)).

**Fix.** Add `cache_control: { type: 'ephemeral' }` breakpoints after the stable prefixes: system block, tool definitions, and the manifest. The manifest is already pinned into the first user message ([src/agent/context-builder.ts:143](../../src/agent/context-builder.ts)), which is cache-friendly — keep it stable.

**Acceptance criteria.**
- Anthropic requests include cache breakpoints on stable prefixes.
- Usage reporting surfaces cache reads (mirror the OpenAI handling).

---

### Task 6 — Nudge the model to batch network reads (MEDIUM, cheap)

**Problem.** The loop already runs consecutive non-mutating network tools concurrently ([src/agent/loop.ts:277](../../src/agent/loop.ts), [src/agent/loop.ts:439](../../src/agent/loop.ts)), but the model issued one `fetch_url` per turn, so parallelism never engaged.

**Fix.** Add a system-prompt line: "When you need several independent pages/APIs, request them as multiple tool calls in one turn — they execute in parallel."

**Acceptance criteria.** Prompt encourages batched tool calls; no code change required beyond the prompt.

---

### Task 7 — Lower-impact tuning (LOW)

- **Faster model tier for the loop.** `deepseek-v4-pro` is heavy; route a faster model for tool-driving, keep a strong one for final synthesis.
- **Lower result caps.** `MAX_TOOL_RESULT_CHARS = 24_000` ([src/agent/context-builder.ts:20](../../src/agent/context-builder.ts)) and the 20k fetch cap ([src/web/fetch.ts:12](../../src/web/fetch.ts)) are generous; smaller caps compound with Task 2.
- **Trim tool schemas.** Full tool-spec JSON is in `fixedTokens` every call ([src/agent/context-builder.ts:136](../../src/agent/context-builder.ts)). `filterToolsForRun` prunes already; tightening descriptions shaves fixed overhead per iteration.
- **Earlier scope check.** Encourage an early `request_user_choice` (e.g. "IMF WEO vs World Bank?") when multiple sources appear, so the model stops exploring. (Largely fixed indirectly by Task 1.)

---

## Suggested order of execution

1. **Task 1** (restore discovery) — biggest single win, contained.
2. **Task 3** (prompt truth) — trivial, removes a whole class of wasted iterations.
3. **Task 2** (proactive compaction) — kills the token-resend blowup; helps every provider.
4. Then Tasks 4–7 as capacity allows.

## How to verify overall

Re-run the `chat log.md` scenario (GDP-PPP for East & SE Asia) after Tasks 1–3 and confirm:
- `web_search` returns real results (no HTML).
- The model completes in well under 25 iterations.
- Per-iteration input-token estimate stays roughly flat instead of climbing.

Relevant existing test suites: `src/web/__tests__/{search,fetch,net}.test.ts`, `src/agent/__tests__/{context-builder,loop}.test.ts`.
