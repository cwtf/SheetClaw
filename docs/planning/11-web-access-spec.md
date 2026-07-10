# Document 11 — Web Access & Scope Clarification Spec

> **Status update (2026-06-11):** search is now **two-tier** — see [Doc 13](13-native-search-spec.md). LLM providers with native search (OpenRouter, Anthropic, Kimi, Qwen, GLM) use their native mechanism and ignore this document's BYOK providers; all other LLM providers continue to use this document's `web_search` tool and BYOK search stack exactly as specified here. Doc 13 governs tier resolution and toggle gating. This document remains authoritative for the BYOK tier, `fetch_url`, `request_user_choice`, context budgets, and the genericity invariant.

Spec for the reattempt of the feature first tried in `5a267ba` (abandoned: its search pipeline could not work generically and was propped up by results hardcoded for a single demo scenario). Two capabilities:

1. **Web access** — the agent can search the web and fetch public URLs to bring external data into the workbook.
2. **Scope clarification** — when an external-data request is ambiguous or large, the agent pauses and asks the user to narrow scope via a structured choice menu instead of downloading speculatively.

## 11.1 Goals and non-goals

**Goals**
- G1. Generic web search + URL fetch that works for *arbitrary* queries and domains, or fails honestly.
- G2. A reusable `request_user_choice` interaction so the model can ask the user to pick from options (scope, dataset, sheet target — any decision), surfaced as a native menu in the taskpane.
- G3. Hard mechanical guardrails on how much fetched content enters model context, independent of model cooperation.
- G4. Privacy-respecting defaults compatible with the AppSource submission (web access off by default, disclosed in the privacy policy).

**Non-goals**
- No backend/proxy service owned by this project (stays a static GitHub Pages deployment).
- No browser automation, login-gated scraping, or JS rendering. Public HTTP(S) GET only.
- No file downloads other than text-representable bodies (text/HTML/JSON/CSV).
- No provider-native server-side search tools (e.g. Anthropic web search tool) in this iteration — tools must work uniformly across Ollama/OpenAI/Anthropic adapters.

## 11.2 Hard constraints (read before designing anything)

- **C1 — CORS.** The taskpane is a browser webview. A direct `fetch()` succeeds only if the target sends permissive CORS headers. Most websites don't; many open-data APIs do. There is no backend to proxy through.
- **C2 — Search needs a configured provider.** Keyless browser-callable *general web search* does not exist: search engines withhold CORS headers and block bot traffic by design (Bing's Search API was retired Aug 2025). "Configured provider" does **not** mean paid — free-tier keys (Tavily ~1 000/mo, Google CSE 100/day) and a keyless self-hosted SearXNG on localhost all qualify (§11.3). What is ruled out is zero-setup search: **if nothing is configured, the `web_search` tool is not exposed to the model at all.** No scraping fallbacks (DuckDuckGo HTML etc.) — they violate ToS, break constantly, and are how the last attempt died. Note `fetch_url` needs no provider: direct fetch + optional keyless reader fallback are free, so the agent can always *read* URLs it has; only *discovery* needs a provider.
- **C3 — Context budget.** Default provider is local Ollama with small context windows. Every tool result must be bounded *before* it reaches the transcript (§11.5).
- **C4 — Genericity invariant.** No domain-, dataset-, or geography-specific identifiers anywhere in `src/` (hostnames, dataset names, query-string patterns, canned menus). The only hostnames permitted in source are search/reader **provider endpoints** declared in the provider registry (§11.4). Enforced by test (§11.9, AC-10).
- **C5 — One code path per behavior.** Model emits a malformed tool call → it gets a `ValidationError` tool result and self-corrects (loop spec §6.7). No prose-inference fallbacks, no UI-side fabrication of options, no synthetic `user` messages — ever.

## 11.3 Phase 0 — provider verification spike (gate for everything else)

Before implementation, verify **from a sideloaded taskpane** (not Node, not curl — CORS behavior differs) for each candidate search provider:

| Check | Record |
|---|---|
| CORS preflight/headers from webview origin | pass/fail |
| Auth mechanism (header vs query param; query-param keys leak into logs) | mechanism |
| Response shape | JSON schema sample |
| Free-tier limits / pricing | numbers |

Candidates, by cost tier:

| Candidate | Cost | Setup burden | Expectation to verify |
|---|---|---|---|
| Tavily | free tier ~1 000 searches/mo | free account, no card | CORS from webview; header auth |
| Google Programmable Search (CSE) JSON API | free 100 queries/day | Google account + create a CSE | googleapis.com generally CORS-friendly; key rides in query param (log-leak concern, §11.8) |
| Jina `s.jina.ai` | free starter quota, then paid | free account | CORS; Bearer header |
| Brave Search API | free tier ~2 000/mo | account (card may be required) | expected to **fail** CORS — confirm before discarding |
| SearXNG self-hosted (localhost) | free, unlimited, keyless | run a Docker container; enable JSON format + CORS in instance config | `http://localhost` fetch from the HTTPS taskpane is allowed (same potentially-trustworthy-origin exemption the Ollama integration relies on) |
| Wikipedia/MediaWiki search API | free, keyless | none | CORS via `origin=*`; encyclopedic only — a supplementary provider, not general web search |
| `r.jina.ai` reader (keyless tier) | free, ~20 req/min | none | fetch-only fallback for `fetch_url`, not a search provider |

A local sidecar proxy (planning suite decision D1) would also unlock free search by moving fetches out of the browser entirely; it is out of scope here but noted as the escape hatch if every browser-side candidate fails.

Record results in Appendix A. **Gate:** at least one search provider must pass browser-side, else `web_search` is descoped and only `fetch_url` ships. Do not start Phase 1 with this unresolved — the last attempt discovered it mid-build and papered over it with hardcoded results instead of escalating it as a design decision.

## 11.4 Architecture

### Module layout

```
src/web/
  providers/        # search provider adapters (mirrors src/adapters/ pattern)
    index.ts        # registry: ProviderId -> SearchProviderAdapter
    tavily.ts       # (whichever pass Phase 0)
    ...
  search.ts         # web_search spec + handler
  fetch.ts          # fetch_url spec + handler (SSRF guards, caps, reader fallback)
  net.ts            # shared fetch w/ timeout, size ceiling, URL validation
src/agent/choice.ts # request_user_choice spec (loop-level tool, as in attempt)
```

Web tools do **not** live under `src/workbook/tools/` and must not run inside `Excel.run`.

### Executor change

`ToolSpec` gains `runtime?: 'excel' | 'none'` (default `'excel'`). For `runtime: 'none'`, `ToolExecutor.execute` invokes the handler directly without the Excel runner. Error mapping unchanged; keep `ToolNetworkError → 'NetworkError'` from the attempt.

### SearchProviderAdapter

```ts
interface SearchProviderAdapter {
  id: string;                       // 'tavily' | 'google-cse' | 'jina' | 'searxng' | keyless source ids
                                    // ('custom' is covered by the per-provider Base URL override)
  label: string;
  requiresKey: boolean;
  search(query: string, opts: { maxResults: number; apiKey: string; baseUrl?: string; signal: AbortSignal }):
    Promise<Array<{ title: string; url: string; snippet?: string; publishedAt?: string }>>;
}
```

Adapters parse **their provider's documented JSON response** — no text/markdown scraping parsers. An empty result set is returned as-is (`results: []`); nothing is ever merged in.

### Configuration & secrets

- `AppConfig` gains `webAccess: { provider: 'none' | ProviderId; baseUrl?: string; readerFallback: boolean }` (persisted; default `{ provider: 'none', readerFallback: false }`).
- API key stored via the existing `AuthState` pattern (`authStateRef`-style, same storage as LLM provider keys) under ref `search:<providerId>`.
- The composer **Search toggle is session state** (session slice, default off each new session), not persisted config — matching its "for this session" semantics and the privacy posture. The persisted part is *which provider/key* to use, not *whether the agent may browse right now*.

### Tool exposure gating

The Search toggle is the single gate for **all** web egress, and the toggle itself requires a configured provider:

- **No provider/key configured** → the toggle cannot be enabled. It renders in a disabled visual state but still receives clicks; clicking surfaces an error hint ("Configure a search key in Settings → Search") with an action that jumps straight to the Search settings tab. (Fluent's `disabled` prop swallows clicks, so render soft-disabled — `aria-disabled` + disabled styling — with an onClick that shows the hint.)
- **Provider configured + toggle ON** → both `web_search` and `fetch_url` are included in the request's tool specs.
- **Toggle OFF** → neither tool is exposed.
- Removing/clearing the key in Settings resets the session toggle to OFF.

## 11.5 Tool schemas

Conventions per Doc 4. None of these are mutating; none take `workbook_id`.

### `web_search` (runtime: none)

**Description (LLM-facing):** "Search the web for current information. Returns result links and snippets only — use fetch_url to read a result. If results suggest several distinct datasets or interpretations of the user's request, use request_user_choice before fetching large data."

| Param | T | R | Constraints |
|---|---|---|---|
| query | string | ✓ | non-empty after trim, ≤ 400 chars |
| max_results | number | ✗ | 1–10, default 5 |

**Returns:** `{ query, provider, results: [{ title, url, snippet?, publishedAt? }] }`
**Errors:** `NetworkError` (provider HTTP/transport failure — message includes provider id and status), `ValidationError` (bad args).
**Hard rule:** zero results → `results: []`. No injected, cached, or "known" results under any condition.

### `fetch_url` (runtime: none)

**Description (LLM-facing):** "Fetch a public HTTP(S) URL. Defaults to preview mode, which returns metadata plus a bounded sample (CSV: header + first rows; JSON: truncated structure; text/HTML: leading text). Call again with mode:'full' only after the scope is confirmed; output is still capped."

| Param | T | R | Constraints |
|---|---|---|---|
| url | string | ✓ | absolute http(s); public host (SSRF guard §11.8) |
| mode | string | ✗ | `"preview"` (default) \| `"full"` |
| format | string | ✗ | `"auto"` (default) \| `"text"` \| `"json"` \| `"csv"` |
| max_chars | number | ✗ | clamp 1 000–20 000; default 4 000 (preview) / 16 000 (full) |

**Returns (envelope, always):** `{ url, finalUrl, status, contentType, source: 'direct'|'reader', format, bytesFetched, returnedChars, truncated }` plus per-format payload:

- `text` → `{ text }` (HTML stripped to text when content-type is HTML)
- `json` → parse, then **re-serialize and cap at `max_chars`**: fits → `{ data }`; doesn't → `{ dataPreview, hint }` where `dataPreview` is the serialization truncated at a clean boundary and `hint` says how to narrow (e.g. "12 480 array items; refetch with a narrower API query or ask the user to choose a scope"). **The parsed-JSON bypass of the char cap was a bug in the attempt; the cap applies to every format.**
- `csv` → `{ headerRow, previewRows (≤ 20), approxRowCount, text? }` — `text` only in `full` mode, capped at `max_chars`; never both full text and preview redundantly.

**Network behavior:** direct fetch with 15 s timeout and 1 MB body ceiling; redirects followed with post-redirect URL re-validation. On transport/CORS failure only (not HTTP error statuses), if `webAccess.readerFallback` is enabled: **one** retry via `https://r.jina.ai/<original-url>` (single wrap, URL passed verbatim — the attempt's double-wrapped, scheme-mangled builder is the canonical example of what not to do), result labeled `source: 'reader'`. Reader disabled or also failing → `NetworkError` whose message states the **classified** cause and tells the model what to do next: a timeout abort → "timed out" (one retry reasonable); a `TypeError` from `fetch` triggers a short `mode: 'no-cors'` reachability probe — probe resolves → confirmed CORS block ("do not retry this host"), probe fails → network/DNS unreachable. Hosts with a *confirmed* CORS block are remembered in a session-scoped in-memory cache; later fetches to the same host fail fast without a network attempt (the reader fallback, if enabled, still applies). The cache is learned at runtime — never a hardcoded host list (C4). Transient failures (timeouts, unreachable) are not cached.

### `request_user_choice` (loop-level, runtime: n/a)

**Description (LLM-facing):** "Show the user a menu and wait for their selection. Use when the request is ambiguous or when fetching everything would be large — e.g. multiple matching datasets, several granularities, or unclear target range. Derive options from information you actually found; include an 'other' escape option."

| Param | T | R | Constraints |
|---|---|---|---|
| question | string | ✓ | non-empty, ≤ 200 chars |
| options | array | ✓ | 2–8 items; each `{ id, label, description? }` **or** a bare string (id slugified from label — the only permitted leniency, documented in the schema, not guessed across alias keys) |
| allow_multiple | boolean | ✗ | default false |

**Returns:** `{ selected_ids, selected_options }` as a normal tool result.
**Errors:** `ValidationError` (question empty, < 2 valid options, > 8 options) — fed back to the model verbatim so it self-corrects. **No option fabrication anywhere** (loop or UI). Dismiss → `ok: false`, `PermissionDenied`, "User dismissed the choice menu", loop continues so the model can wrap up or ask in prose.

## 11.6 Agent-loop integration

Port from the attempt (it was sound), minus the fallbacks:

- `request_user_choice` intercepted in `executeCall` before the executor; sets `status: 'awaiting_choice'` + `session.pendingChoice`; `waitForChoice` promise mirrors `waitForConfirmation`; `stop()` aborts it; resolution clears `pendingChoice` and appends the tool result.
- **Deleted concepts:** `inferChoiceFallback` (prose → menu hijack), `executeAssistantChoiceFallback`, synthetic `user` messages, `fallbackChoiceOptions` in the UI. A model that asks in prose gets a prose answer from the user — that's acceptable behavior, not something to intercept.
- Tool-spec assembly (toggle filtering + appending `REQUEST_USER_CHOICE`) is computed once per run, as in the attempt.

### System-prompt additions (generic — C4 applies)

- Workflow rule: "For external data: search first; read previews before full fetches; never paste large raw payloads into your reply — write data to the workbook with tools."
- Clarification rule: "Before fetching data in full, if the request could map to more than one distinct source, table, or granularity — or a preview shows the data is larger than what the task needs — call `request_user_choice` with options built from what you actually found. Do not enumerate options as plain text."
- No dataset names, no example menus, no geography. The *mechanical* guardrail is preview-mode-by-default + caps (§11.5); the prompt is guidance, not the enforcement layer.

## 11.7 UI spec

- **Composer:** Search toggle (pill button) + Auto-approve, as in the attempt. Three toggle states: *off* (secondary appearance), *on* (primary appearance), *unavailable* (no key configured — soft-disabled per §11.4; clicking shows a dismissible MessageBar "Web search needs a provider key. Configure it in Settings → Search." with an **Open Settings** action wired through `onOpenSettings`, extended to accept a target sub-tab: `onOpenSettings('search')`). Send/textarea disabled during `awaiting_choice`.
- **ChoiceBlock:** port the attempt's component (numbered options, descriptions, multi-select, Dismiss/Continue) with two changes: render exactly `pendingChoice.options` (no fallback path), and **no pre-selected default** (Continue stays disabled until the user picks — pre-selecting the first option biases the choice).
- **Status line:** `awaiting_choice` → "Awaiting selection".
- **Settings → Search tab:** new `SettingsTabKey: 'search'` in `SettingsPanel`'s existing `SETTINGS_TABS` (alongside `ollama`/`apiKeys`/`generic`). Contents: provider dropdown (None default), API key field (masked, stored via the `AuthState` pattern under ref `search:<providerId>`, same storage as LLM keys), optional base URL (SearXNG/custom), reader-fallback checkbox with a plain-language note that it routes fetched URLs through jina.ai, and a signup link for the selected provider (e.g. Tavily's free-key page) to soften BYOK friction. SHOULD include a "Test key" button that runs a 1-result query and reports pass/fail inline.
- Transcript rendering of web tool calls/results follows existing `Tool:`/`OK|ERR` meta rows (with the attempt's overflow CSS fixes).

## 11.8 Security & privacy

- **SSRF guards** (keep from attempt): http/https only; block localhost/`.local`/`.localhost`, private/loopback/link-local IPv4 and IPv6; re-validate after redirects. Add: block credentials-in-URL (`user:pass@host`) and non-default ports other than 80/443/8080.
- **Egress disclosure:** queries go to the configured search provider; fetched URLs go to the target site and (only if enabled) the reader proxy. Update `public/privacy.html` accordingly **in the same PR** that ships the tools (the attempt regressed AppSource posture by omitting this).
- Off by default at every layer: toggle off per session, provider `none` until configured, reader fallback off.
- API keys never placed in query strings if the provider supports header auth (Phase 0 records this per provider).

## 11.9 Testing & acceptance criteria

Spec-first: these tests are written against this document, not against the implementation.

| # | Criterion | Type |
|---|---|---|
| AC-1 | Toggle off → no web tool specs in the LLM request | unit |
| AC-2 | No provider/key configured → toggle renders unavailable (`aria-disabled`), cannot be enabled, and no web tool specs reach the LLM request; clicking it shows the configure-key hint whose action opens Settings → Search. With a key configured and toggle ON, both web tools are present | unit + component |
| AC-3 | Provider adapter parses its documented JSON for fixture queries from **≥ 3 unrelated domains** (e.g. finance, weather, sports — no shared vocabulary with any demo) | unit, mocked fetch |
| AC-4 | All network mocked dead → `web_search` returns `NetworkError`; **transcript contains zero result URLs** (kills fabricated results structurally) | unit |
| AC-5 | `fetch_url` JSON body > `max_chars` → `truncated: true`, `returnedChars ≤ max_chars`; same for text and csv | unit |
| AC-6 | `fetch_url` to private IP / localhost / post-redirect-private → `ValidationError`, no request issued | unit |
| AC-7 | `request_user_choice` with 0/1 options or empty question → `ValidationError` tool result; session never enters `awaiting_choice`; second model turn with fixed args succeeds | unit (loop) |
| AC-8 | Prose-only clarification from model → run ends as normal text turn; no menu appears | unit (loop) |
| AC-9 | Choice flow: select → tool result contains `selected_ids`; dismiss → `PermissionDenied`; no `role: 'user'` message is ever appended by the loop | unit (loop) |
| AC-10 | Genericity: a test greps `src/` for hostnames not in the provider registry allowlist and for any string literal also appearing in the demo scenario list (maintained in the test). Build fails on hit | static |
| AC-11 | Manual sideload: three end-to-end scenarios on unrelated domains (e.g. World Bank indicator CSV, a public JSON API, a Wikipedia table), each exercising search → choice menu (options visibly derived from live results) → preview → full fetch → confirmed `write_range` | manual checklist |

AC-11's three scenarios must be *named in the PR description with screenshots*; one scenario passing is explicitly insufficient to merge — the last attempt was validated against a single demo prompt, which is exactly how the hardcoding crept in.

## 11.10 Build sequencing

**Precondition:** build on master (rolled back to `6cea710` + this spec). The failed attempt is preserved on the local `attempt/web-search` branch (`5a267ba`); do not build on top of it — salvageable pieces are cherry-picked per the list below, not inherited.

| Phase | Scope | Gate |
|---|---|---|
| 0 | Provider verification spike (§11.3); fill appendix | ≥ 1 provider passes browser-side, else descope `web_search` |
| 1 | `runtime: 'none'` executor support; `src/web/net.ts` + `fetch_url` with caps/preview/SSRF; AC-5/6 | tests green |
| 2 | Provider registry + verified adapter(s); `web_search`; Settings UI + key storage; AC-1/2/3/4 | tests green |
| 3 | `request_user_choice` + ChoiceBlock + loop integration (ported, fallback-free); AC-7/8/9 | tests green |
| 4 | System-prompt rules; privacy.html update; AC-10 static check; AC-11 manual pass | PR with 3 documented scenarios |

Cherry-pick from `5a267ba` where it's clean: ChatPanel overflow CSS, `ToolNetworkError` plumbing, `PendingChoice` types/store, ChoiceBlock markup, SSRF validators.

## 11.11 Open questions (decide before Phase 2)

1. ~~Which search path to target first~~ — **Resolved: BYOK (bring-your-own-key), see §11.12.** Each user configures their own provider + free-tier key, mirroring how SheetClaw already handles LLM API keys. Works identically for personal use and published distribution.
2. ~~Is the jina.ai reader fallback acceptable privacy-wise for an AppSource-distributed add-in~~ — **Resolved for now: keep it, off by default, disclosed.** Reader fallback is opt-in and the egress (fetched URLs → reader service) is disclosed in `public/privacy.html` under "Optional web search and URL fetching". `fetch_url` direct-only remains the stricter alternative if a certification reviewer objects; routing blocked fetches through a local OpenClaw instance is recorded in Doc 12 §12.9 as a future possibility. Re-verify against Partner Center policy before submission — see [Doc 15](15-appsource-readiness.md).
3. ~~Should `fetch_url` be allowed without the Search toggle?~~ — **Resolved: no.** The Search toggle is the single gate for all web egress, and it requires a configured key (§11.4). A user who wants the agent to fetch a pasted URL configures a key and enables Search like everyone else.

## 11.12 Distribution scenario analysis (personal vs published)

How the search-provider decision plays out per audience. The governing fact: **SheetClaw is a static client-side bundle, so any API key shipped inside it is extractable by every user.** A shared developer key is therefore ruled out under every scenario — it would be stolen, the developer would pay for all usage, and one abuser exhausts the quota for everyone.

| Scenario | Search path | Cost to dev | Cost to user | Notes |
|---|---|---|---|---|
| Personal use (current) | Your own free-tier key, or local SearXNG | $0 | $0 | Any §11.3 candidate works |
| **Published, free, no backend (AppSource)** | **BYOK — each user supplies their own free-tier key in Settings** | **$0** | **$0 (free tier) unless heavy use** | Identical to how users already supply OpenAI/Anthropic keys or point at their own Ollama. Same settings pattern, same key storage, same privacy story ("your key, your traffic"). Setup friction is the price; mitigate with a Settings link to the provider signup page |
| Published, users on cloud LLMs | Provider-native web search (Anthropic/OpenAI server-side search tools) | $0 | metered on their existing LLM key | Zero extra signup, no CORS (search runs on the LLM provider's servers). Excluded from this iteration (§11.1 non-goal: Ollama parity), but the cleanest future upgrade for cloud-key users — design the toggle/gating so a native-search mode can slot in later |
| Published with backend/monetization | Developer-hosted proxy (e.g. Cloudflare Worker) holding the key, per-user rate limits | infra + search API usage at scale | $0 or subscription | Only path to "search just works" with zero user setup. Requires abandoning the no-backend constraint, absorbing usage costs, and taking on real privacy obligations (user queries transit your server → privacy policy, retention, abuse handling). Revisit only if the project moves to accounts/monetization |
| Power users / privacy-first | Self-hosted SearXNG (localhost or LAN), `searxng` provider with base URL | $0 | $0 + container upkeep | Keep as a supported provider; never the default for general users (Docker is a non-starter for typical Excel users) |

**Decision (confirmed 2026-06-10):** BYOK is the chosen path, with a dedicated Search tab in Settings and a soft-disabled Search toggle when no key is configured (§11.4, §11.7). The spec's architecture (user-configured provider + key via the existing `AuthState` pattern, `provider: 'none'` default) already implements BYOK, so the published scenario requires no design change — only that AppSource copy and the Settings UI clearly state that web search is optional, off by default, and powered by the user's own provider account. Provider-native search is noted as the planned follow-up for cloud-LLM users.

## Appendix A — Phase 0 results

| Provider | CORS/preflight | Auth mechanism | Status | Free-tier limits / pricing | Response shape sample | Notes |
|---|---|---|---|---|---|---|
| Tavily | pass / pass | bearer-header | 200 | Free-tier key; spec expectation: about 1,000 searches/month | `{ query: string, follow_up_questions: null, answer: null, images: [], results: [{ url: string, title: string, content: string, score: number, raw_content: null }], response_time: number, request_id: string }` | Verified 2026-06-10 from the sideloaded taskpane Phase 0 diagnostics surface. Endpoint: `https://api.tavily.com/search`; key provided by user and not committed. |
| Google Programmable Search | **not yet verified from taskpane** | `X-Goog-Api-Key` header + `cx` query param | — | free 100 queries/day | `{ items: [{ title, link, snippet }] }` | Adapter implemented 2026-06-10; needs an Engine ID (cx) configured in Settings. Verify with the Settings "Test key" button before relying on it. |
| Jina `s.jina.ai` | **not yet verified from taskpane** | bearer-header (`X-Respond-With: no-content`) | — | free starter quota | `{ data: [{ title, url, description, date? }] }` | Adapter implemented 2026-06-10. Verify with "Test key". |
| SearXNG (self-hosted) | **not yet verified from taskpane** | none (instance URL in Settings; default `http://localhost:8080/search`) | — | free, unlimited | `{ results: [{ title, url, content, publishedDate? }] }` | Adapter implemented 2026-06-10. Instance must enable `format=json` and CORS. Verify with "Test search". |
| Wikipedia/MediaWiki | **not yet verified from taskpane** | none (keyless, `origin=*`) | — | free | `{ query: { search: [{ title, snippet }] } }` | Adapter implemented 2026-06-10; encyclopedic only. Verify with "Test search". |

**Gate decision:** Tavily passed browser-side CORS/auth verification, so Phase 1 may proceed with `web_search` in scope for the BYOK Tavily path. The four later adapters are implemented against documented API shapes but keep "not yet verified" status until each passes the Settings test button from a sideloaded taskpane. Brave Search API remains unimplemented (expected CORS failure, §11.3); `custom` is covered by each adapter's Base URL override rather than a dedicated provider.

## Appendix B — AC-11 manual sideload checklist

Automated acceptance criteria AC-1 through AC-10 are covered by the Vitest suite and static genericity guard. AC-11 remains the manual release gate and must be filled with user-confirmed sideload runs before this feature is considered fully done.

| Scenario | Domain/source family | Required flow | Status | Evidence/notes |
|---|---|---|---|---|
| Public CSV import | Open-data CSV or API-exported CSV | search → choice menu → preview → full fetch → confirmed `write_range` | Pending confirmation | |
| Public JSON import | Public JSON API | search → choice menu → preview → full fetch → confirmed `write_range` | Pending confirmation | |
| Public table import | Public HTML/table source | search → choice menu → preview → full fetch → confirmed `write_range` | Pending confirmation | |

Manual notes so far: a long web workflow that exceeded the original 25-iteration cap was user-tested after adding same-session continuation and was reported as working ("looks good to me"). This validates the continuation fix, but does not by itself satisfy AC-11's three unrelated-domain scenario requirement.
