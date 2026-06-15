# Document 3 — Data Models and Storage Schema

All shapes are TypeScript-style. IDs are ULIDs unless noted. Timestamps are ISO-8601 UTC strings.

## 3.1 Core data structures

### 3.1.1 `Message`
A single conversation-history entry. Discriminated union on `role`/`kind`.

```ts
type Message =
  | UserMessage | AssistantMessage | ToolCallMessage | ToolResultMessage
  | ConfirmationMessage | SystemNoticeMessage;

interface BaseMessage {
  id: string;                 // ULID
  sessionId: string;
  createdAt: string;
  // provider-format payload is derived at send time; UI keeps the normalized form
}

interface UserMessage extends BaseMessage { role: 'user'; text: string; }

interface AssistantMessage extends BaseMessage {
  role: 'assistant';
  text: string;               // may stream/grow during a run
  toolCalls?: ToolCall[];     // present when the model requested tools this turn
  usageRef?: string;          // UsageRecord.id for this turn
  finishReason?: 'stop'|'tool_calls'|'length'|'error';
}

interface ToolCallMessage extends BaseMessage {  // UI-facing rendering of a single call
  role: 'tool_call'; toolCall: ToolCall; status: 'pending'|'awaiting_confirmation'|'applied'|'failed';
}

interface ToolResultMessage extends BaseMessage {
  role: 'tool'; toolCallId: string; result: ToolResult;
}

interface ConfirmationMessage extends BaseMessage {
  role: 'confirmation'; pendingChangeId: string; decision?: 'apply'|'cancel'|'apply_all';
}

interface SystemNoticeMessage extends BaseMessage {  // truncation, errors, workbook changes
  role: 'system_notice'; level: 'info'|'warn'|'error'; text: string;
}
```

### 3.1.2 `ToolCall` and `ToolResult`

```ts
interface ToolCall {
  id: string;                 // provider tool-call id (or generated for Ollama)
  name: string;               // e.g. "write_range"
  arguments: Record<string, unknown>;  // parsed (post-reassembly) args
  rawArguments?: string;      // original JSON string, for debugging malformed calls
  workbookId: string;         // resolved scope for this call
  mutating: boolean;          // from tool registry
}

interface ToolResult {
  toolCallId: string;
  ok: boolean;
  data?: unknown;             // tool-specific success payload (Doc 4 return shapes)
  error?: {
    code: 'ValidationError'|'WorkbookNotFound'|'RangeError'|'OfficeApiError'|'PermissionDenied'|'Unsupported';
    message: string;          // safe to feed back to the model so it can correct
    details?: unknown;
  };
  snapshotId?: string;        // set when a mutating tool captured a snapshot
  durationMs?: number;
}
```

### 3.1.3 `WorkbookHandle`

```ts
interface WorkbookHandle {
  workbookId: string;         // session-scoped stable id (ULID) — see [D8]
  name: string;               // file name, e.g. "Budget.xlsx"
  isActive: boolean;          // scoped session target
  isHost: boolean;            // the workbook hosting the add-in
  sheets: SheetSummary[];
  lastRefreshed: string;
  capability: 'full'|'host-only';  // degraded mode if multi-workbook unsupported
}

interface SheetSummary {
  name: string; position: number; visible: boolean;
  usedRange?: { address: string; rowCount: number; colCount: number };
  headers?: string[];         // first-row sample, optional
}

type WorkbookManifest = { active: string; workbooks: Omit<WorkbookHandle,'sheets'>[] & { sheets: SheetSummary[] }[] };
```

### 3.1.4 `SnapshotEntry`

```ts
interface SnapshotEntry {
  id: string;
  sessionId: string;
  workbookId: string;
  sheet: string;
  kind: 'range'|'chart'|'pivot'|'sheet';
  target: string;             // range address, chart name, pivot name, sheet name
  before: {
    values?: unknown[][];     // for range
    formulas?: unknown[][];
    numberFormat?: string[][];
    definition?: unknown;     // serialized chart/pivot definition (best-effort)
  };
  payloadRef?: string;        // IndexedDB key when 'before' is large (spilled)
  createdAt: string;
  appliedToolCallId?: string;
  undone: boolean;
  restoreFidelity: 'full'|'values-only'|'structural-coarse';
}
```

### 3.1.5 `UsageRecord` (per-turn)

```ts
interface UsageRecord {
  id: string;
  sessionId: string;
  turnIndex: number;
  timestamp: string;
  provider: 'ollama'|'openai'|'anthropic'|'generic';
  model: string;
  inputTokens: number;
  outputTokens: number;
  cacheReadTokens?: number;   // Anthropic / OpenAI prompt caching
  cacheWriteTokens?: number;
  totalTokens: number;
  estimatedCostUsd: number;   // computed from PricingTable at record time
  pricingVersion?: string;    // which PricingTable entry was used
  estimated: boolean;         // true when token counts came from local estimate, not provider
  toolCallsCount: number;
}
```

### 3.1.6 `ProviderConfig`

```ts
interface ProviderConfig {
  provider: 'ollama'|'openai'|'anthropic'|'generic';
  enabled: boolean;
  baseUrl: string;            // overridable
  model: string;              // selected model id
  knownModels?: ModelInfo[];  // last successful listModels() cache
  authMode: 'apikey'|'oauth'|'none';
  // auth state (tokens stored separately under a guarded key — see §3.2)
  authStateRef: string;       // key into AUTH store
  headers?: Record<string,string>;  // generic extra headers
  contextLimits: {
    maxContextTokens: number; // model window
    historyTokenCap: number;
    maxInlineSheetCells: number;
  };
  temperature?: number;
  maxOutputTokens?: number;
}
```

### 3.1.7 `PricingTable`  **[DECISION REQUIRED D7: bundled static vs fetched-cached]**

```ts
interface PricingTable {
  version: string;            // e.g. "2026-06-01"
  updatedAt: string;
  entries: PricingEntry[];
  defaults: { inputPerMTok: number; outputPerMTok: number };  // fallback when model unknown
}

interface PricingEntry {
  provider: string;
  modelMatch: string;         // exact id or prefix/glob, e.g. "gpt-4o*"
  inputPerMTok: number;       // USD per 1M input tokens
  outputPerMTok: number;
  cacheReadPerMTok?: number;
  cacheWritePerMTok?: number;
  currency: 'USD';
}
```

### 3.1.8 `AgentSession`

```ts
interface AgentSession {
  id: string;
  createdAt: string;
  scope: SessionScope;        // { workbookId } (active workbook at start)
  status: 'idle'|'building'|'calling_llm'|'parsing'|'awaiting_confirmation'|'executing_tool'|'error'|'done'|'stopped';
  iteration: number;
  maxIterations: number;      // safety cap (e.g. 25)
  provider: string; model: string;
  messageIds: string[];       // ordered
  pendingChange?: PendingChange;  // when awaiting_confirmation
  tokenBudget: { used: number; window: number };
  lastError?: { code: string; message: string };
  totals: { inputTokens: number; outputTokens: number; costUsd: number };
}

interface PendingChange {
  id: string; toolCall: ToolCall; snapshotId: string;
  diff: CellDiff[]; severity: 'normal'|'elevated';  // elevated = chart/pivot/large
  workbookName: string; sheet: string;
}
interface CellDiff { address: string; before: unknown; after: unknown; }
```

## 3.2 localStorage schema

Origin is `https://localhost:<port>`. localStorage is ~5 MB per origin; large items spill to IndexedDB.

| Key | Schema | Max-size concern | Eviction strategy |
|-----|--------|------------------|-------------------|
| `xl.config.providers` | `Record<provider, ProviderConfig>` | Small (<10 KB). | Overwrite in place; never evicted. |
| `xl.config.app` | `{ activeProvider, ui prefs, pricingMode, autoApproveSession:false }` | Small. | Overwrite in place. |
| `xl.auth.<provider>` | `EncryptedBlob` of `{ apiKeyMasked?, tokens?: {access, refresh, expiresAt}, state }` **[D3]** | Small. **Plaintext risk** — see Doc 5 §security and [D3]. | Cleared on sign-out; refresh updates in place. |
| `xl.pricing.table` | `PricingTable` (cached) | <50 KB. | Replaced on update; versioned. |
| `xl.usage.index` | `{ days: { 'YYYY-MM-DD': recordCount }, oldest, newest, totalBytesApprox }` | Small. | Drives 30-day window pruning. |
| `xl.usage.day.<YYYY-MM-DD>` | `UsageRecord[]` for that day | Each day small; **30 days aggregate could approach quota** (Risk R7). | **30-day rolling window**: on write, delete day-keys older than 30 days; if `QuotaExceededError`, delete oldest day-keys until the write succeeds, raise `usage:quota-warning`. |
| `xl.sessions.recent` | `AgentSession[]` (last N, e.g. 50, metadata + messageIds) | Medium; messages stored separately. | Ring buffer, cap N; oldest dropped. |
| `xl.messages.<sessionId>` | `Message[]` | **Can be large** for long sessions. | Stored only for recent sessions; pruned with `xl.sessions.recent`. Spill to IndexedDB if >256 KB. |
| `xl.snapshots.index` | `SnapshotEntry[]` (metadata; `before` payloads referenced) | Medium. | Pruned per policy: keep snapshots for the current + last session, or last 24 h; oldest pruned first. |
| IndexedDB `snapshots` store | `{ key, payload }` large `before` blobs | Bounded by snapshot prune policy. | Deleted when the referencing `SnapshotEntry` is pruned. |
| IndexedDB `messages` store | overflow message arrays | Bounded by session ring buffer. | Deleted with the session. |

### Storage rules
1. **Single source of truth per item** — config in `xl.config.*`, never duplicated into session blobs.
2. **Quota guard wrapper** — all writes go through a `storage.put(key, value)` that catches `QuotaExceededError`, runs the relevant eviction, retries once, then raises a typed event.
3. **Schema versioning** — every top-level blob carries `{ _v: number }`; a migration step runs on load.
4. **Secrets** — `xl.auth.*` is the only sensitive blob. **[D3] resolved (implemented):** secrets are sealed with AES-GCM-256 before they reach localStorage, using a non-extractable key held in IndexedDB (`src/auth/secureStore.ts`). This protects the at-rest copy (disk inspection, storage dumps, exports). Same-origin XSS can still call decrypt — only an OS credential vault outside the WebView closes that, which a backend-less add-in cannot offer. This is an acceptable BYOK posture for AppSource ("your key, your traffic"); residual risk is tracked in [Doc 15](15-appsource-readiness.md).
5. **Workbook handles are NOT persisted** **[D8]** — `WorkbookHandle.workbookId` is session-scoped and re-derived on each add-in load via `WorkbookRegistry.refresh()`. Persisting them is unsafe because Office handles do not survive reloads.
