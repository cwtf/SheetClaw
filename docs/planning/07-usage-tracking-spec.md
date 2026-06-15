# Document 7 — Usage Tracking Specification

## 7.1 Usage record schema

See Doc 3 §3.1.5 for the canonical `UsageRecord`. Recap of the fields and their source:

| Field | Source | Notes |
|-------|--------|-------|
| id, sessionId, turnIndex, timestamp | loop | one record per LLM call |
| provider, model | active ProviderConfig | |
| inputTokens, outputTokens | provider usage (preferred) or local estimate | `estimated:true` when local |
| cacheReadTokens, cacheWriteTokens | provider usage | Anthropic / OpenAI caching; else 0/undefined |
| totalTokens | sum | |
| estimatedCostUsd | PricingTable × tokens | computed at record time, frozen |
| pricingVersion | PricingTable.version | for auditability |
| estimated | bool | token source flag |
| toolCallsCount | loop | tool calls emitted this turn |

## 7.2 Token-count extraction per provider

| Provider | Where usage lives | Extraction |
|----------|-------------------|-----------|
| **OpenAI / Generic** | final stream chunk `usage` (requires `stream_options.include_usage:true`) | `prompt_tokens`→input, `completion_tokens`→output, `prompt_tokens_details.cached_tokens`→cacheRead. Non-stream: `response.usage`. |
| **Anthropic** | `message_start.message.usage` (input, cache_creation/read) + `message_delta.usage.output_tokens` (final) | input = `input_tokens`; cacheWrite = `cache_creation_input_tokens`; cacheRead = `cache_read_input_tokens`; output from the last `message_delta`. |
| **Ollama** | native `/api/chat` final message: `prompt_eval_count` (input), `eval_count` (output) | OpenAI-compat endpoint may also return `usage`. |
| **Generic (no usage)** | some endpoints omit usage | **Fallback**: local tokenizer estimate of input + output; `estimated:true`. |

Adapters surface a normalized `usage` `LLMStreamEvent` `{inputTokens, outputTokens, cacheRead?, cacheWrite?, source:'provider'|'estimated'}`. UsageTracker consumes that — it never parses provider responses itself.

## 7.3 Cost estimation formula

```
cost = (inputTokens        - cacheReadTokens) /1e6 * inputPerMTok
     + (cacheReadTokens)  /1e6 * cacheReadPerMTok      (if defined, else inputPerMTok)
     + (cacheWriteTokens) /1e6 * cacheWritePerMTok     (if defined, else inputPerMTok)
     + (outputTokens)     /1e6 * outputPerMTok
```
- Pricing resolved by `findPricing(provider, model)`: exact match → prefix/glob match → `defaults`. On `defaults`, set `pricingVersion:"default"` and mark the row "estimated pricing" in the dashboard (Risk R4).
- Ollama (local) entries default to `0` cost but still track tokens.

## 7.4 Pricing table structure

See Doc 3 §3.1.7. **[DECISION REQUIRED D7]**: bundle a static `pricing.json` (simple, may go stale — Risk R4) vs. fetch-and-cache from a remote source (fresher, adds a network dependency and a trust decision). Recommendation (holds for published/AppSource distribution too): **bundled static with a visible `updatedAt` and an in-app "edit pricing" affordance**, so the user can correct any model's rate without a code change.

## 7.5 Pre-execution estimation

Before a run (and live as the user types), ChatPanel shows an estimated cost:
1. ContextBuilder assembles (or simulates) the request and computes **input tokens** via the tokenizer (OpenAI-family BPE; char/4 heuristic for others; Anthropic count-tokens endpoint optionally for accuracy).
2. **Output tokens are unknown pre-call** → estimate using a configurable assumed-output (e.g. 500 tokens) or a rolling per-session average of recent outputs.
3. `estimatedCost = price(inputTokens) + price(assumedOutput)`, shown as `~$0.00x` with a tooltip clarifying it's an estimate and assumptions used.
4. For agentic multi-step runs, the single-turn estimate is labelled "per step"; the Footer accumulates actuals as the run proceeds.

## 7.6 Session aggregation logic

- In-memory `SessionAggregate { sessionId, inputTokens, outputTokens, costUsd, turns, byModel }` updated on each `usage:recorded`.
- Footer reads `getSessionTotals(currentSessionId)`.
- On session end, the aggregate is finalized; per-turn records remain the source of truth (aggregates are derivable, never the only copy).

## 7.7 localStorage persistence — 30-day rolling window

- **Per-day bucket**: `xl.usage.day.<YYYY-MM-DD>` = `UsageRecord[]`. Writing a record appends to today's bucket.
- **Index**: `xl.usage.index` tracks `{ days:{date:count}, oldest, newest, totalBytesApprox }`.
- **Rolling window**: on each append, delete any `xl.usage.day.*` key with date < (today − 30 days); update index.
- **Eviction on overflow**: the storage wrapper catches `QuotaExceededError` → delete the oldest day bucket(s) until the append succeeds → raise `usage:quota-warning` (surfaced as a dashboard banner) (Risk R7).
- **Why per-day buckets**: bounded write amplification (don't rewrite all history per record) and O(1) window pruning by key.

## 7.8 Dashboard data queries

All derived from day buckets (lazy-loaded for the active range):

| Query | Computation |
|-------|-------------|
| Today / Week / Month / All-time totals | sum cost+tokens over buckets in the range |
| By provider | group records by `provider` |
| By model | group by `model` |
| By day | per-bucket sums (for the time-series chart) |
| By session | group by `sessionId` (join with `xl.sessions.recent` for labels/timestamps) |
| Estimated-vs-actual share | count `estimated:true` vs false |

Queries run in-memory after loading the relevant buckets; for "all-time" the dashboard loads all buckets (bounded to 30 days by the window).

## 7.9 CSV export format

`exportCsv(filter)` returns a UTF-8 CSV (BOM-prefixed for Excel) with header:

```
timestamp,session_id,turn_index,provider,model,input_tokens,output_tokens,
cache_read_tokens,cache_write_tokens,total_tokens,estimated,estimated_cost_usd,pricing_version,tool_calls
```

- One row per `UsageRecord`, RFC-4180 quoting.
- Filtered by the dashboard's active time range/provider/model.
- Delivered via a Blob + `download` link (browser-side; no sidecar needed).
- A second optional "summary" export emits the aggregate table (by day / provider / model).
