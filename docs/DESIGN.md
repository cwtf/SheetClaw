# SheetClaw — Complete Design & Implementation Specification

**Version:** derived from the current source at v0.1.0.1c (15 July 2026)
**Purpose:** This document is a complete, self-contained specification of SheetClaw. An engineer (or LLM) with no access to the original source should be able to rebuild the entire program from this document alone. It specifies the architecture, every data type, every algorithm with its exact constants, the full tool catalogue with parameter schemas, all wire protocols, storage layouts, UI behavior, and build configuration.

> **Search design note:** the keyless search bundle is **baked into the agent's tool set on every run**, regardless of the Search toggle or Settings. The Search toggle and Settings → Search provider selection govern **only** internet search that requires an API key (keyed BYOK providers) or is billed to the LLM provider (native search).

---

## Table of contents

1. [Product overview](#1-product-overview)
2. [Technology stack & project layout](#2-technology-stack--project-layout)
3. [Runtime environment & hard constraints](#3-runtime-environment--hard-constraints)
4. [Core data model (types)](#4-core-data-model-types)
5. [State management (Zustand store)](#5-state-management-zustand-store)
6. [Persistence layer & storage keys](#6-persistence-layer--storage-keys)
7. [Credential security](#7-credential-security)
8. [OAuth flow (OpenRouter PKCE)](#8-oauth-flow-openrouter-pkce)
9. [LLM provider adapter layer](#9-llm-provider-adapter-layer)
10. [Native web search capabilities](#10-native-web-search-capabilities)
11. [The agent loop](#11-the-agent-loop)
12. [Context builder & compaction](#12-context-builder--compaction)
13. [System prompt](#13-system-prompt)
14. [Workbook layer: registry, executor, snapshots](#14-workbook-layer-registry-executor-snapshots)
15. [Tool catalogue](#15-tool-catalogue)
16. [Web access layer](#16-web-access-layer)
17. [Pricing & usage tracking](#17-pricing--usage-tracking)
18. [User interface](#18-user-interface)
19. [Build, deployment & Office manifest](#19-build-deployment--office-manifest)
20. [Testing strategy](#20-testing-strategy)
21. [Suggested implementation order](#21-suggested-implementation-order)
22. [Key invariants & gotchas](#22-key-invariants--gotchas)

---

## 1. Product overview

SheetClaw is an **agentic AI chat assistant embedded in Microsoft Excel** as an Office.js task-pane add-in. The user chats with an LLM that can read, analyse, and edit the host workbook by calling a large registry of typed tools (read ranges, write values, format cells, create charts/pivots/tables, sort/filter, manage sheets, add comments/shapes, protect sheets, etc.). Optional web tools (`web_search`, `fetch_url`) bring in external data.

Design pillars:

- **BYOK (bring your own key), local-first.** No server component. The task pane talks directly from the browser WebView to the user's chosen LLM provider (Ollama local, OpenAI, Anthropic, or any OpenAI-compatible endpoint). API keys are encrypted at rest in the browser.
- **Human-in-the-loop safety.** By default, every *mutating* tool call pauses the agent for user confirmation, showing a cell-level diff computed from a pre-write snapshot. Users can persistently switch the approval mode to Accept all edits. Snapshots also power an Undo button.
- **Provider-agnostic streaming.** All providers are normalized to one streamed event union (`LLMStreamEvent`) so the agent loop is provider-independent.
- **Cost transparency.** Per-turn token usage is recorded with bundled pricing data; the Usage view shows totals, per-model breakdowns, and a per-day sparkline, exportable to CSV.

The UI is a five-view React SPA (Chat, History, Usage, Settings, About) built with Fluent UI v9 and rendered in Excel's sidebar. Chat is the home view; the other views are reached from a compact overflow menu.

---

## 2. Technology stack & project layout

### Dependencies (`package.json`)

```json
{
  "name": "sheetclaw",
  "private": true,
  "license": "PolyForm-Noncommercial-1.0.0",
  "dependencies": {
    "@fluentui/react-components": "^9.74",
    "react": "^18.3", "react-dom": "^18.3",
    "semver": "^7.8", "ulid": "^3.0", "zustand": "^5.0"
  },
  "devDependencies": {
    "@types/office-js": "^1.0", "@types/react": "^18.3", "@types/react-dom": "^18.3",
    "@vitejs/plugin-react": "^6", "typescript": "^5", "vite": "^8", "vitest": "^4",
    "office-addin-debugging": "^6", "office-addin-dev-certs": "^2", "office-addin-manifest": "^2"
  }
}
```

### npm scripts

| Script | Command | Purpose |
|---|---|---|
| `dev` | `vite` | HTTPS dev server on `localhost:3000` |
| `build` | `tsc && vite build` | Typecheck + bundle to `dist/` |
| `test` | `vitest run` | Unit tests (Node env, `src/**/*.test.ts`) |
| `test:watch` | `vitest` | Watch mode |
| `test:providers` | `vitest run --config vitest.integration.config.ts` | Live-network CORS checks (`src/**/*.integration.ts`, 120 s timeout) |
| `sideload` | `office-addin-debugging start manifest.xml desktop --app excel` | Register add-in with Excel |
| `validate-manifest` | `office-addin-manifest validate manifest.xml` | |
| `install-certs` / `uninstall-certs` | `office-addin-dev-certs install/uninstall` | Dev HTTPS certs |

### Directory layout

```
/                      manifest.xml, taskpane.html, oauth-start.html, oauth-callback.html,
                       vite.config.ts, tsconfig.json, vitest.integration.config.ts, public/
src/
  adapters/            LLM provider adapters: openai.ts, anthropic.ts, ollama.ts,
                       native-search.ts, harness.ts, index.ts (createAdapter factory)
  agent/               loop.ts (AgentLoop), context-builder.ts, system-prompt.ts,
                       choice.ts (request_user_choice), index.ts (singleton)
  auth/                secureStore.ts (AES-GCM), oauthFlow.ts (PKCE), credentials.ts
  lib/                 ulid.ts (re-export)
  pricing/             index.ts (match + cost), pricing.json (bundled table)
  store/               index.ts (combined store), storage.ts (versioned envelope),
                       slices/{config,auth,session,usage}.ts
  taskpane/            index.tsx (Office.onReady bootstrap), App.tsx/App.css (themed view shell),
                       selection.ts (submitted Excel selection), workbookLayer.ts
                       (singleton wiring), components/*.tsx
  types/               index.ts (barrel) + llm, message, provider, session, snapshot,
                       tool, usage, workbook
  usage/               queries.ts (aggregation), export.ts (CSV)
  web/                 fetch.ts (fetch_url), search.ts (web_search), net.ts (guarded fetch),
                       providers/ (15 search adapters + keyless bundle + index)
  workbook/            registry.ts, executor.ts, snapshot.ts, a1notation.ts,
                       unsupported-error.ts, index.ts (createWorkbookLayer),
                       tools/ (specs.ts, write.ts, charts.ts, pivots.ts, tables.ts,
                               sort_filter.ts, sheets.ts, validation.ts,
                               conditional_formats.ts, range_ops.ts, names.ts,
                               comments.ts, shapes.ts, layout_protection.ts,
                               workbook_tools.ts, args.ts)
```

### tsconfig essentials

`target ES2020`, `lib [ES2020, DOM, DOM.Iterable]`, `module ESNext`, `moduleResolution bundler`, `jsx react-jsx`, `strict`, `noEmit`, `types: ["office-js"]`, include `src`.

---

## 3. Runtime environment & hard constraints

1. **Browser WebView inside Excel.** No Node APIs at runtime, no file system, no server. All outbound HTTP is subject to **CORS** enforced by the WebView.
2. **All Office.js calls must be batched inside `Excel.run(async ctx => { ...; await ctx.sync(); })`.** To keep modules testable in Node (where the `Excel` global doesn't exist), an injectable runner type alias is used everywhere:
   ```ts
   type ExcelRunner = <T>(fn: (ctx: Excel.RequestContext) => Promise<T>) => Promise<T>;
   // default: fn => Excel.run(fn); tests substitute a mock.
   ```
3. **Host-only workbook model.** Windows desktop Office cannot enumerate other open workbooks from a task pane. The `WorkbookRegistry` always holds exactly one `WorkbookHandle` (capability `'host-only'`) representing the host workbook.
4. **Tests run in Node** (Vitest, `environment: 'node'`); `Excel`/`Office` globals must be mocked or kept out of tested code paths. Node's fetch does not enforce CORS, so a separate live integration suite (`test:providers`) probes keyless search endpoints for CORS headers.
5. **App version constant.** Vite `define` injects `__APP_VERSION__` (string) from `package.json` version; declared in `src/vite-env.d.ts`.

---

## 4. Core data model (types)

All types live under `src/types/` and are re-exported from a barrel `src/types/index.ts`. Reproduce them exactly.

### 4.1 Tool types (`types/tool.ts`)

```ts
export interface ToolCall {
  id: string;
  name: string;
  arguments: Record<string, unknown>;
  rawArguments?: string;        // unparsed JSON as streamed
  workbookId: string;
  mutating: boolean;
}

export interface ToolResult {
  toolCallId: string;
  ok: boolean;
  data?: unknown;
  error?: {
    code: 'ValidationError' | 'WorkbookNotFound' | 'RangeError' | 'OfficeApiError'
        | 'PermissionDenied' | 'NetworkError' | 'Unsupported';
    message: string;
    details?: unknown;
  };
  snapshotId?: string;
  durationMs?: number;
}

export interface ToolSpec {
  name: string;
  description: string;
  parameters: JSONSchemaObject;
  mutating: boolean;
  runtime?: 'excel' | 'none';   // 'none' = no Excel.run context needed (web tools)
}

export interface JSONSchemaObject {
  type: 'object';
  properties: Record<string, JSONSchemaProperty>;
  required?: string[];
  additionalProperties?: boolean;
}

export type JSONSchemaProperty =
  | { type: 'string'; description?: string; enum?: string[] }
  | { type: 'number'; description?: string }
  | { type: 'boolean'; description?: string }
  | { type: 'array'; items: JSONSchemaProperty; description?: string }
  | JSONSchemaObject
  | { description?: string; [key: string]: unknown };
```

### 4.2 LLM wire types (`types/llm.ts`)

```ts
export interface NormalizedToolCall { id: string; name: string; arguments: Record<string, unknown>; }

export type NormalizedMessage =
  | { role: 'system'; content: string }
  | { role: 'user'; content: string }
  | { role: 'assistant'; content: string; toolCalls?: NormalizedToolCall[] }
  | { role: 'tool'; toolCallId: string; name: string; content: string };

export interface LLMRequest {
  model: string;
  messages: NormalizedMessage[];
  tools: ToolSpec[];
  nativeSearch?: NativeSearchCapability;  // see §10
  system?: string;
  temperature?: number;
  maxTokens?: number;
}

export type LLMStreamEvent =
  | { type: 'text-delta'; delta: string }
  | { type: 'tool-call-start'; index: number; id: string; name: string }
  | { type: 'tool-call-delta'; index: number; argumentsDelta: string }
  | { type: 'tool-call-end'; index: number }
  | { type: 'usage'; inputTokens: number; outputTokens: number;
      cacheRead?: number; cacheWrite?: number; source: 'provider' | 'estimated' }
  | { type: 'done'; finishReason: 'stop' | 'tool_calls' | 'length' | 'error' }
  | { type: 'error'; error: LLMError };

export type LLMError =
  | { code: 'AuthError'; message: string }
  | { code: 'RateLimitError'; message: string; retryAfter?: number }
  | { code: 'NetworkError'; message: string }
  | { code: 'ProviderError'; message: string; status: number; body?: unknown }
  | { code: 'MalformedResponseError'; message: string }
  | { code: 'NotSupported'; message: string };

export interface LLMClient {
  chat(req: LLMRequest, signal: AbortSignal): AsyncIterable<LLMStreamEvent>;
  listModels(): Promise<ModelInfo[]>;
  capabilities(): ProviderCapabilities;
}
```

### 4.3 Provider types (`types/provider.ts`)

```ts
export interface ModelInfo { id: string; name?: string; contextWindow?: number; supportsTools?: boolean; }

export interface ProviderCapabilities {
  supportsTools: boolean; supportsStreaming: boolean; supportsOAuth: boolean;
  nativeUsage: boolean; toolFormat: 'openai' | 'anthropic';
}

export interface ProviderConfig {
  provider: ProviderKey;
  label?: string;
  enabled: boolean;
  baseUrl: string;
  model: string;
  knownModels?: ModelInfo[];
  authMode: 'apikey' | 'oauth' | 'none';
  authStateRef: string;                    // storage key, e.g. 'xl.auth.openai'
  headers?: Record<string, string>;
  contextLimits: { maxContextTokens: number; historyTokenCap: number; maxInlineSheetCells: number; };
  temperature?: number;
  maxOutputTokens?: number;
}

export interface AuthState {
  provider: string;
  state: 'unauthenticated' | 'authenticating' | 'authenticated'
       | 'token-expired' | 'validating' | 'error';
  apiKeyMasked?: string;
  authMode?: 'apikey' | 'oauth' | 'none';
  oauthProvider?: 'openrouter';
  accessToken?: string;
  refreshToken?: string;
  tokenType?: string;
  expiresAt?: string;
  userId?: string;
  error?: string;
  /** Raw key in memory only; persisted AES-GCM-sealed via auth/secureStore. */
  _key?: string;
}

export interface WebAccessConfig {
  provider: WebAccessProvider;   // 'none' | keyed SearchProviderId (tavily, google-cse,
                                 // jina, searxng). 'none' = keyless bundle only; the
                                 // keyless bundle is NOT user-selectable — it is always
                                 // available to the agent as the default search backend.
  baseUrl?: string;              // optional override (self-hosted SearXNG etc.)
  engineId?: string;             // Google CSE cx
  readerFallback: boolean;       // route failed fetches through r.jina.ai
}
```

### 4.4 Usage/pricing types (`types/usage.ts`)

```ts
export type ProviderKey =
  | 'ollama' | 'openai' | 'anthropic' | 'generic' | 'deepseek' | 'groq' | 'mistral'
  | 'together' | 'kimi' | 'glm' | 'qwen' | 'llama' | 'gemini' | 'cerebras'
  | 'cloudflare' | 'huggingface';

export interface UsageRecord {
  id: string; sessionId: string; turnIndex: number; timestamp: string;
  provider: ProviderKey; model: string;
  inputTokens: number; outputTokens: number;
  cacheReadTokens?: number; cacheWriteTokens?: number; totalTokens: number;
  estimatedCostUsd: number; pricingVersion?: string; estimated: boolean;
  toolCallsCount: number;
}

export interface PricingEntry {
  provider: string;          // provider key or '*'
  modelMatch: string;        // exact id, 'prefix*', or '*'
  inputPerMTok: number; outputPerMTok: number;
  cacheReadPerMTok?: number; cacheWritePerMTok?: number;
  currency: 'USD';
}

export interface PricingTable {
  version: string; updatedAt: string;
  entries: PricingEntry[];
  defaults: { inputPerMTok: number; outputPerMTok: number };
}
```

### 4.5 Message types (`types/message.ts`)

The chat transcript is a flat list of discriminated messages; every message has `id` (ULID), `sessionId`, `createdAt` (ISO string).

```ts
export type Message = UserMessage | AssistantMessage | ToolCallMessage
                    | ToolResultMessage | ConfirmationMessage | SystemNoticeMessage;

interface BaseMessage { id: string; sessionId: string; createdAt: string; }
export interface UserMessage extends BaseMessage {
  role: 'user'; text: string; selection?: WorkbookSelection;
}
export interface WorkbookSelection { sheet: string; address: string; }
export interface AssistantMessage extends BaseMessage {
  role: 'assistant'; text: string; toolCalls?: ToolCall[];
  usageRef?: string; finishReason?: 'stop' | 'tool_calls' | 'length' | 'error';
}
export interface ToolCallMessage extends BaseMessage {
  role: 'tool_call'; toolCall: ToolCall;
  status: 'pending' | 'awaiting_confirmation' | 'applied' | 'failed';
}
export interface ToolResultMessage extends BaseMessage {
  role: 'tool'; toolCallId: string; result: ToolResult;
}
export interface ConfirmationMessage extends BaseMessage {
  role: 'confirmation'; pendingChangeId: string; decision?: 'apply' | 'cancel' | 'apply_all';
}
export interface SystemNoticeMessage extends BaseMessage {
  role: 'system_notice'; level: 'info' | 'warn' | 'error'; text: string;
}
```

`selection` is the sheet and local A1 address captured when the message is submitted. It is stored with the message so phrases such as “this cell” remain bound to the original selection even if the user later moves elsewhere in Excel. The UI renders the user text alone; the context builder appends the selection metadata only in the provider-facing message.

`ToolCallMessage`, `ConfirmationMessage`, `SystemNoticeMessage` are **UI-only** — they are skipped when converting the transcript back to LLM wire messages (§12).

### 4.6 Session types (`types/session.ts`)

```ts
export interface SessionScope { workbookId: string; }
export interface CellDiff { address: string; before: unknown; after: unknown; }

export interface PendingChange {
  id: string; toolCall: ToolCall; snapshotId: string; diff: CellDiff[];
  severity: 'normal' | 'elevated';         // elevated when diff.length > 50
  workbookName: string; sheet: string;
}

export type SessionStatus =
  | 'idle' | 'building' | 'calling_llm' | 'parsing' | 'awaiting_confirmation'
  | 'awaiting_choice' | 'executing_tool' | 'error' | 'done' | 'stopped';

export interface AgentSession {
  id: string; createdAt: string;
  scope: SessionScope;
  status: SessionStatus;
  iteration: number; maxIterations: number;
  provider: string; model: string;
  messageIds: string[];
  pendingChange?: PendingChange;
  pendingChoice?: PendingChoice;           // see §11.6
  webSearchEnabled: boolean;
  stopReason?: 'max_iterations';
  tokenBudget: { used: number; window: number };
  lastError?: { code: string; message: string };
  totals: { inputTokens: number; outputTokens: number; costUsd: number };
}
```

### 4.7 Snapshot types (`types/snapshot.ts`)

```ts
export interface SnapshotEntry {
  id: string; sessionId: string; workbookId: string; sheet: string;
  kind: 'range' | 'chart' | 'pivot' | 'sheet';
  target: string;                          // A1 address or object name
  before: {
    values?: unknown[][]; formulas?: unknown[][];
    numberFormat?: string[][]; definition?: unknown;
  };
  payloadRef?: string;
  createdAt: string;
  appliedToolCallId?: string;
  undone: boolean;
  restoreFidelity: 'full' | 'values-only' | 'structural-coarse';
}
```

### 4.8 Workbook types (`types/workbook.ts`)

```ts
export interface SheetSummary {
  name: string; position: number; visible: boolean;
  usedRange?: { address: string; rowCount: number; colCount: number };
  headers?: string[];
}
export interface WorkbookHandle {
  workbookId: string; name: string; isActive: boolean; isHost: boolean;
  sheets: SheetSummary[]; lastRefreshed: string;
  capability: 'full' | 'host-only';
}
export type WorkbookManifest = { active: string; workbooks: WorkbookHandle[]; };
```

---

## 5. State management (Zustand store)

One global store combining four slices (`src/store/index.ts`):

```ts
export type AppStore = ConfigSlice & AuthSlice & SessionSlice & UsageSlice;
export const useStore = create<AppStore>()((...a) => ({
  ...createConfigSlice(...a), ...createAuthSlice(...a),
  ...createSessionSlice(...a), ...createUsageSlice(...a),
}));
```

Selectors: `selectActiveProvider`, `selectActiveProviderConfig`, `selectIsProviderReady(provider)`.

### 5.1 Config slice

State: `providers: Record<ProviderKey, ProviderConfig>` and `appConfig: AppConfig`.

```ts
export interface AppConfig {
  activeProvider: ProviderKey;
  autoApproveSession: boolean;      // persisted approval mode; skip confirmations while true
  pricingMode: 'bundled' | 'custom';
  theme: 'system' | 'light' | 'dark';
  webAccess: WebAccessConfig;
}
// Defaults: { activeProvider: 'ollama', autoApproveSession: false,
//             pricingMode: 'bundled', theme: 'system',
//             webAccess: { provider: 'none', readerFallback: false } }
```

**Default provider configs** (all `contextLimits` default to `{128000, 100000, 5000}` unless noted; all `enabled: false` except ollama; `authStateRef: 'xl.auth.<key>'`):

| key | label | baseUrl | default model | authMode | notes |
|---|---|---|---|---|---|
| ollama | Ollama (local) | `http://localhost:11434` | `llama3.2` | none | enabled by default |
| openai | OpenAI | `https://api.openai.com/v1` | `gpt-4o` | apikey | |
| anthropic | Anthropic | `https://api.anthropic.com` | `claude-sonnet-4-6` | apikey | limits `{200000,160000,8000}` |
| generic | Generic / OpenRouter | `https://openrouter.ai/api/v1` | `openai/gpt-4o-mini` | oauth | |
| deepseek | DeepSeek | `https://api.deepseek.com` | `deepseek-v4-flash` | apikey | |
| groq | Groq | `https://api.groq.com/openai/v1` | `llama-3.3-70b-versatile` | apikey | |
| mistral | Mistral | `https://api.mistral.ai/v1` | `mistral-large-latest` | apikey | |
| together | Together AI | `https://api.together.ai/v1` | `meta-llama/Llama-3.3-70B-Instruct-Turbo` | apikey | |
| kimi | Kimi | `https://api.moonshot.ai/v1` | `kimi-k2.6` | apikey | limits `{256000,200000,8000}` |
| glm | GLM | `https://api.z.ai/api/paas/v4` | `glm-4.7` | apikey | |
| qwen | Qwen | `https://dashscope-intl.aliyuncs.com/compatible-mode/v1` | `qwen-plus` | apikey | |
| llama | Llama | `https://api.llama.com/compat/v1` | `Llama-3.3-70B-Instruct` | apikey | |
| gemini | Google AI Studio | `https://generativelanguage.googleapis.com/v1beta/openai` | `gemini-2.0-flash` | apikey | limits `{1000000,800000,20000}` |
| cerebras | Cerebras | `https://api.cerebras.ai/v1` | `llama-3.3-70b` | apikey | |
| cloudflare | Cloudflare Workers AI | `https://api.cloudflare.com/client/v4/accounts/YOUR_ACCOUNT_ID/ai/v1` | `@cf/meta/llama-3.3-70b-instruct-fp8-fast` | apikey | |
| huggingface | Hugging Face | `https://router.huggingface.co/v1` | `meta-llama/Llama-3.3-70B-Instruct` | apikey | |

Actions:
- `setProvider(key, patch)` — merge patch, persist to `xl.config.providers`.
- `setActiveProvider(key)` — sets `appConfig.activeProvider` **and** flips that provider's `enabled: true`; persists both keys.
- `setAppConfig(patch)` — merge + persist to `xl.config.app`.
- `loadConfigFromStorage()` — hydrate with defaults-merge, then run migrations:
  1. If `webAccess.provider` is an individual keyless source id **or** the legacy `'keyless'` bundle id → rewrite to `'none'` and clear `baseUrl` (keyless search is now baked in and no longer a selectable provider; the provider field selects a keyed provider only).
  2. If stored active provider has `enabled: false` → set true.
  3. If generic provider points at OpenRouter with an empty model → restore default model (`openai/gpt-4o-mini`).

Config hydration is a shallow merge over `DEFAULT_APP_CONFIG`, so configurations saved before theme support automatically receive `theme: 'system'`. Theme changes use the ordinary `setAppConfig` persistence path.

### 5.2 Auth slice

State: `authStates: Record<ProviderKey, AuthState>` and `searchAuthStates: Record<SearchProviderId, AuthState>` (search states use `provider: 'search:<id>'`). All initialized `{provider, state: 'unauthenticated'}`.

Key behaviors:
- `saveApiKey(provider, key)`: trims; sets `state: 'authenticated'`, `authMode: 'apikey'`, `apiKeyMasked` (first 4 + `••••` + last 4, or 8 bullets if ≤ 8 chars), `_key`. Persist sealed.
- `saveOAuthCredential(provider, {accessToken, oauthProvider?, userId?, refreshToken?, tokenType?='Bearer', expiresAt?})`: like above with `authMode: 'oauth'`, stores token in both `accessToken` and `_key`.
- `clearApiKey` / `saveSearchApiKey` / `clearSearchApiKey` — analogous.
- **Ordered persistence queue**: all writes chain through a single `pendingWrites: Promise` so an async encrypt can't let a later `clear` be overwritten by an earlier `save`. Export `flushAuthPersistence(): Promise<void>` for tests.
- Sealing: before persist, encrypt fields `_key`, `accessToken`, `refreshToken` via `encryptSecret` (§7). On load, decrypt; if a value lacks the `enc1:` prefix it's a legacy plaintext entry — decrypt is a no-op and the state is **re-persisted sealed** (migration). On decrypt failure, replace with `{state:'unauthenticated', error:'Saved credential could not be unlocked. Enter it again in Settings.'}`.
- `isProviderReady(p)`: ollama is ready when `unauthenticated` or `authenticated`; others require `authenticated` and, if `expiresAt` set, not expiring within 60 s.
- `isSearchProviderReady(p)`: keyless providers (adapter `requiresKey === false`, includes the bundle) are always ready; keyed providers need `authenticated`.

### 5.3 Session slice

State: `currentSession: AgentSession | null`, `messages: Message[]` (current session only), `chatHistory: ChatHistoryItem[]`, `webSearchEnabled: boolean` (UI toggle).

```ts
export interface ChatHistoryItem {
  id: string; title: string; preview: string;
  createdAt: string; updatedAt: string;
  provider: string; model: string;
  status: AgentSession['status']; messageCount: number;
}
```

Every mutation **persists the full transcript** `{session, messages}` to `xl.chat.history.<sessionId>` and rewrites the index `xl.chat.history.index` (sorted by `updatedAt` desc). Title = first user message compacted (whitespace-collapsed, ≤96 chars with `...`), fallback `'Untitled chat'`; preview = latest non-confirmation message summarized (`Tool: <name>` for tool_calls, `Tool result: OK`/`Tool error: <msg>` for results).

Actions:
- `setSession(session|null)` — replaces current session, clears messages; persists an empty transcript.
- `updateSession(patch)` / `updateSessionById(id, patch)` — the ById variant also patches **non-current** sessions by loading their persisted transcript, patching, and re-persisting (needed because the loop keeps running by session id even if the user switches chats).
- `appendMessage(msg)` / `updateMessage(id, patch)` — same dual path (current in-memory vs. persisted transcript); appendMessage also appends the id to `session.messageIds` (dedup).
- `loadChatHistory()`, `resumeChat(id)` (restores transcript; if the stored status was an active status, downgrade to `'stopped'` with `lastError = {code:'SessionInterrupted', message:'This chat was restored after an interrupted run.'}` and clear pending change/choice), `deleteChat(id)`, `deleteAllChatHistory()` (removes every `xl.chat.history.*` key).

Active statuses set (used in several places): `{'building','calling_llm','parsing','awaiting_confirmation','awaiting_choice','executing_tool'}`.

### 5.4 Usage slice

State: `sessionTotals: {sessionId, inputTokens, outputTokens, costUsd, turns} | null` (in-memory only).

`recordUsage(record)`:
1. Derive `day = timestamp.slice(0,10)`.
2. Prune day buckets older than **30 days** (`ROLLING_DAYS`).
3. Append record to `xl.usage.day.<day>` array.
4. Update `xl.usage.index` (`{days: Record<day, count>, oldest, newest, totalBytesApprox}`).
5. If `sessionTotals.sessionId` matches, accumulate in memory.

`resetSessionTotals(sessionId, totals?)`, `clearSessionTotals()`.

---

## 6. Persistence layer & storage keys

`src/store/storage.ts` wraps `localStorage` with a **versioned envelope** (`SCHEMA_VERSION = 1`):

- `pack(value)`: objects get `{...value, _v: 1}`; arrays are boxed as `{_v: 1, _arr: [...]}` (arrays can't be spread-merged with `_v`).
- `get<T>(key)`: parse; if `_v !== 1` return `null`; unbox `_arr` or strip `_v`.
- `put<T>(key, value)`: on `QuotaExceededError`, evict the **oldest** `xl.usage.day.*` key (ISO dates sort lexicographically) and retry once; if it still fails, dispatch `window` CustomEvent `'xl:quota-warning'` with `{detail:{key}}`.
- `remove(key)`.
- No-ops when `localStorage` is undefined (Node tests).

**Complete storage key inventory:**

| Key | Content |
|---|---|
| `xl.config.providers` | `Record<ProviderKey, ProviderConfig>` |
| `xl.config.app` | `AppConfig` |
| `xl.auth.<providerKey>` | sealed `AuthState` |
| `xl.auth.search:<searchProviderId>` | sealed `AuthState` |
| `xl.chat.history.index` | `ChatHistoryItem[]` |
| `xl.chat.history.<sessionId>` | `{session: AgentSession, messages: Message[]}` |
| `xl.usage.day.<YYYY-MM-DD>` | `UsageRecord[]` |
| `xl.usage.index` | `UsageIndex` |
| `xl.keystore.jwk` | JWK fallback key (only when IndexedDB unavailable) |

IndexedDB: database `xl.keystore`, object store `keys`, record id `primary` holds the `CryptoKey`.

---

## 7. Credential security

`src/auth/secureStore.ts`. Secrets are sealed with **AES-GCM-256** via Web Crypto before touching localStorage.

- Format: `enc1:<base64 iv (12 bytes)>:<base64 ciphertext>`. `isEncryptedSecret(v)` = starts with `enc1:`.
- `encryptSecret(plain)`: random 12-byte IV, `crypto.subtle.encrypt({name:'AES-GCM', iv}, key, utf8(plain))`.
- `decryptSecret(v)`: values without the prefix are returned **as-is** (legacy plaintext support). Malformed prefix/base64 or decrypt failure throws `SecretDecryptError`.
- **Key acquisition** (cached in a module-level promise; `__resetSecureStoreForTests()` clears it):
  1. IndexedDB: get-or-create a **non-extractable** AES-GCM-256 `CryptoKey` stored directly in IDB (structured clone).
  2. Fallback: JWK in `localStorage` (`xl.keystore.jwk`) — generate extractable, export JWK, then re-import **non-extractable** for use.
  3. Last resort (unit tests): ephemeral in-memory non-extractable key.
- Threat model note (keep in code comment): protects the at-rest copy only; same-origin XSS can still call decrypt.

`src/auth/credentials.ts`:

```ts
const EXPIRY_MARGIN_MS = 60_000;
isAuthExpired(auth)  // true if expiresAt parses and expiresAt - 60s <= now
getAuthCredential(auth) // '' if missing/expired, else auth.accessToken ?? auth._key ?? ''
```

---

## 8. OAuth flow (OpenRouter PKCE)

`src/auth/oauthFlow.ts`. Sign-in with OpenRouter via PKCE, working both inside Office (Dialog API) and in a plain browser (popup + postMessage).

Constants: auth URL `https://openrouter.ai/auth`, exchange URL `https://openrouter.ai/api/v1/auth/keys`, timeout 5 min.

Flow (`signInWithOpenRouter()`):
1. `createPKCEPair()`: `codeVerifier` = 64 random bytes base64url; `codeChallenge` = base64url(SHA-256(verifier)); method `S256`.
2. Callback URL: `<origin>/oauth-callback.html?provider=openrouter&state=<24 random bytes base64url>`.
3. Auth URL: `https://openrouter.ai/auth?callback_url=...&code_challenge=...&code_challenge_method=S256`.
4. If `Office.context.ui.displayDialogAsync` exists → open `<origin>/oauth-start.html?to=<authUrl>` in an Office dialog (70% h × 45% w). `oauth-start.html` validates the `to` origin is exactly `https://openrouter.ai` then redirects. The dialog page eventually lands on `oauth-callback.html`, which posts `{type:'xl-oauth-callback', provider, code, state, error}` via `Office.context.ui.messageParent(JSON.stringify(payload))` (and via `window.opener.postMessage` for the popup path), then closes itself after 600 ms.
5. Otherwise open a popup (`window.open(..., 'popup=yes,width=540,height=720')`) and listen for same-origin `message` events; poll popup closure every 500 ms.
6. Validate the callback: `type === 'xl-oauth-callback'`, no `error`, `code` & `state` present, `state` matches — otherwise throw (state mismatch: "OAuth state mismatch. Sign-in was cancelled for safety.").
7. Exchange: POST JSON `{code, code_verifier, code_challenge_method}` to the exchange URL → `{key, user_id?}`. Store via `saveOAuthCredential`.

Office dialog error 12006 = user closed the window.

---

## 9. LLM provider adapter layer

### 9.1 Factory (`adapters/index.ts`)

```ts
export function createAdapter(cfg: ProviderConfig, auth: string | AuthState = ''): LLMClient {
  const apiKey = typeof auth === 'string' ? auth : getAuthCredential(auth);
  switch (cfg.provider) {
    case 'anthropic': return new AnthropicAdapter({ apiKey, baseUrl: cfg.baseUrl, provider: cfg.provider });
    case 'ollama':    return new OllamaAdapter({ baseUrl: cfg.baseUrl });
    default:          return new OpenAIAdapter({ apiKey, baseUrl: cfg.baseUrl,
                        provider: cfg.provider, extraHeaders: cfg.headers });
    // default covers: openai, generic, deepseek, groq, mistral, together, kimi,
    // glm, qwen, llama, gemini, cerebras, cloudflare, huggingface
  }
}
```

### 9.2 OpenAIAdapter (OpenAI-compatible chat completions)

Capabilities: `{supportsTools: true, supportsStreaming: true, supportsOAuth: false, nativeUsage: true, toolFormat: 'openai'}`.

`listModels()`: GET `<baseUrl>/models` with `Authorization: Bearer <key>` → map `data[].id`.

`chat(req, signal)` — POST `<baseUrl>/chat/completions`:

Request body:
```jsonc
{
  "model": req.model,
  "messages": [ /* system message first if req.system, then serialized history */ ],
  "stream": true,
  "stream_options": { "include_usage": true },
  "tools": [ /* function tools + native-search entries (§10) */ ],  // only if non-empty
  "temperature": ..., "max_tokens": ...                              // only if set
  // + native-search body patch (e.g. qwen enable_search), if any
}
```

Message serialization:
- tool result → `{role:'tool', tool_call_id, content}`; **plus `name: '$web_search'` when the tool name is `$web_search`** (Kimi native search echo).
- assistant with toolCalls → `{role:'assistant', content: content || null, tool_calls: [{id, type:'function', function:{name, arguments: JSON.stringify(args)}}]}`.
- otherwise `{role, content}`.

Tool serialization: `{type:'function', function:{name, description, parameters:{type, properties, required}}}` (drop `additionalProperties`).

SSE parsing: split stream on `\n`, yield payloads of lines starting `data: ` (flush trailing buffer too). Stop at `[DONE]`. Per chunk:
- `choices[0].delta.content` → `text-delta`.
- `delta.tool_calls[]`: each has `index`; when `id` **and** `function.name` present it's the first delta for that index → store accumulator `{id, name, argsBuf}` and emit `tool-call-start` (+ a `tool-call-delta` if arguments came along); otherwise append `function.arguments` to that index's buffer and emit `tool-call-delta`.
- `choices[0].finish_reason` (guarded by a `doneSent` flag — some providers e.g. OpenRouter repeat it): emit `tool-call-end` for every accumulated index, then `done` with that reason.
- `chunk.usage` (may arrive on a choice-less chunk): emit `usage` with `prompt_tokens`, `completion_tokens`, `cacheRead = prompt_tokens_details?.cached_tokens`, `source: 'provider'`.

Error mapping: non-OK response → 401 → `AuthError`; 429 → `RateLimitError` (parse `retry-after` header); else `ProviderError {status, body}`. Fetch throw → `NetworkError` (suppressed if aborted). Empty body → `MalformedResponseError`. JSON parse failure of a chunk → `MalformedResponseError` (include first 80 chars) and stop.

### 9.3 AnthropicAdapter

Config: `{apiKey, baseUrl? (default https://api.anthropic.com), provider?}`. API version header `anthropic-version: 2023-06-01`. Capabilities: toolFormat `'anthropic'`.

`listModels()`: static list (no public endpoint): `claude-opus-4-8`, `claude-sonnet-4-6`, `claude-haiku-4-5-20251001`.

`chat`: POST `<base>/v1/messages` with headers `x-api-key`, `anthropic-version`, and **`anthropic-dangerous-direct-browser-access: 'true'`** (required for browser calls). Body:
```jsonc
{ "model", "max_tokens": req.maxTokens ?? 4096, "stream": true,
  "messages": serialize(non-system messages), "system": ...,
  "tools": [ ...{name, description, input_schema:{type,properties,required}},
             ...nativeSearchTool? ],
  "temperature": ... }
```

Message serialization rules:
- system messages are skipped (system goes top-level; prefer `req.system`, else first system message content).
- assistant → content array: optional `{type:'text',text}` + one `{type:'tool_use', id, name, input}` per tool call.
- tool results must be **batched into a single user message**: if the previously emitted message is a user message whose content array starts with a `tool_result` block, push into it; else start a new `{role:'user', content:[{type:'tool_result', tool_use_id, content}]}`.

SSE events (each `data:` line JSON; skip malformed):
- `message_start` → capture `usage.input_tokens`.
- `content_block_start` with `content_block.type === 'tool_use'` → accumulator + `tool-call-start`.
- `content_block_delta`: `text_delta` → `text-delta`; `input_json_delta` → append `partial_json`, `tool-call-delta`.
- `content_block_stop` → `tool-call-end` if that index was a tool block.
- `message_delta` → emit `usage {inputTokens, outputTokens: usage.output_tokens, source:'provider'}` then `done` mapping stop_reason: `tool_use`→`tool_calls`, `max_tokens`→`length`, else `stop`.
- `error` → `ProviderError` (status 0) and return.

HTTP error mapping adds: 529 → `RateLimitError` ("Anthropic overloaded (529)").

### 9.4 OllamaAdapter

Wraps an inner `OpenAIAdapter` pointed at `<base>/v1` with dummy key `'ollama'` (base default `http://localhost:11434`, trailing slashes stripped).

`listModels()`: GET `<base>/api/tags` → `{models:[{name, model}]}` → `{id: model ?? name, name}`. On fetch failure, run a **no-cors reachability probe** against the base URL:
- reachable → the server is up but CORS-blocked. Throw an error starting with the sentinel prefix `"Ollama is reachable, but this add-in cannot read it."` including a PowerShell remedy: `[Environment]::SetEnvironmentVariable('OLLAMA_ORIGINS','<origin>','User')` (single-quote-escaped). Helper exports: `getOllamaBrowserAccessCommand(origin?)`, `isOllamaBrowserAccessError(msg)`.
- not reachable → rethrow original.

`chat`: **buffer all events** from the inner adapter first, accumulating assistant text and noting whether any tool call was emitted. If no tool call, text is non-empty, and tools were requested, run the **lenient tool-call parser**:
- Collect candidates from all fenced ```` ```json {...} ``` ```` blocks (regex `/```(?:json)?\s*(\{[\s\S]*?\})\s*```/g`) plus one bare `{...}` match (`/(\{[\s\S]*\})/`).
- For each, JSON.parse; tool name = `parsed.name` or `parsed.tool`; must be in the known tool list. Arguments = `parsed.arguments ?? parsed.parameters ?? parsed.input ?? parsed` with `name`/`tool` keys deleted.
- On success: emit synthetic `tool-call-start` (id `lenient_<Date.now()>`, index 0), `tool-call-delta` (full JSON), `tool-call-end`, then re-emit any buffered `usage` events, then `done {finishReason:'tool_calls'}` and return.
- Otherwise replay the buffered events verbatim.

### 9.5 Harness (diagnostic canary — `adapters/harness.ts`)

A tool-calling smoke test used by a dev panel: tool `harness_echo` (`{value: string}` required, non-mutating), prompt `'Call the harness_echo tool with value "ping". Do not say anything else.'`. `runHarness(client, model, signal)` is an async generator that yields every raw event plus one final `{type:'result', pass, message, toolCallName?, toolCallArgs?}` after checking: no provider error, a tool call was emitted, it was `harness_echo`, arguments are valid JSON with string `value`.

---

## 10. Native web search capabilities

`adapters/native-search.ts`. Some providers offer server-side web search. When present and enabled, the client-side `web_search` tool is dropped and a provider-specific tool/body patch is injected instead.

```ts
export type NativeSearchKind =
  | 'openrouter-server-tool' | 'anthropic-server-tool' | 'kimi-builtin-function'
  | 'qwen-enable-search' | 'glm-web-search-tool';

export interface NativeSearchCapability {
  provider: ProviderKey; kind: NativeSearchKind;
  supportsModel?: (model: string) => boolean;
  modelSupportLabel?: string;
  costNote: string;
}
```

Registry `NATIVE_SEARCH` (partial record by provider):

| provider | kind | model gate | cost note (verbatim in UI) |
|---|---|---|---|
| generic | openrouter-server-tool | — | "OpenRouter native search is billed to your OpenRouter key, currently about $0.005 per search call." |
| anthropic | anthropic-server-tool | — | "...about $10 per 1,000 searches plus result tokens." |
| kimi | kimi-builtin-function | — | "...about $0.005 per search call." |
| qwen | qwen-enable-search | model must equal or prefix-match `qwen3.5-plus`, `qwen3.5-flash`, `qwen3-max` (label: "qwen3.5-plus, qwen3.5-flash, or qwen3-max in thinking mode") | DashScope pricing note |
| glm | glm-web-search-tool | — | Z.AI pricing note |

Functions:
- `getNativeSearchCapability(provider, model)` — undefined if provider not in registry or model gate fails.
- `resolveSearchTier(provider, model)` → `{tier:'native', capability}` or `{tier:'byok', capability?, nativeUnavailableReason?}`.
- `resolveSearchToggle({provider, model, byokReady})` → tier resolution + `{byokReady, available: tier==='native' || byokReady}`. **Scope:** this toggle machinery governs only the keyed/native search tiers. It has no effect on the keyless bundle, which is always available to the agent (§11, §16).
- UI copy helpers: `getSearchSettingsStatusText`, `getByokSectionNote`, `getProviderNativeSearchCaption`, `getUnavailableSearchToggleHint` (compose the cost notes/reasons into settings captions).

**Request patches** (applied only when `capability.provider === activeProvider`):
- `getOpenAINativeSearchPatch(provider, capability)`:
  - openrouter → `{tools: [{type:'openrouter:web_search'}]}`
  - kimi → `{tools: [{type:'builtin_function', function:{name:'$web_search'}}]}`
  - qwen → `{body: {enable_search: true, search_options: {forced_search: false, search_strategy: 'turbo'}}}`
  - glm → `{tools: [{type:'web_search', web_search:{enable:true, search_engine:'search-prime', search_result:true}}]}`
  - anthropic → `{}` (handled in the Anthropic adapter)
- `getAnthropicNativeSearchTool` → `{type:'web_search_20250305', name:'web_search', max_uses: 5}`.

**Kimi echo protocol:** Kimi's builtin `$web_search` is executed client-side by echoing the raw arguments back as the tool result (see §11.5) — the server does the actual search on the next turn.

---

## 11. The agent loop

`src/agent/loop.ts` — the heart of the app. `MAX_ITERATIONS = 50`.

### 11.1 Class shape

```ts
export type LoopRunner = <T>(fn: (ctx: Excel.RequestContext) => Promise<T>) => Promise<T>;
export interface ChoiceSelection { ids: string[]; otherText?: string; }

export class AgentLoop {
  constructor(registry: WorkbookRegistry, executor: ToolExecutor,
              snapshots: SnapshotManager, runner?: LoopRunner /* default Excel.run */)
  start(instruction, scope, client: LLMClient, cfg: ProviderConfig,
        selection?: WorkbookSelection): Promise<void>
  followUp(instruction, scope, client, cfg,
           selection?: WorkbookSelection): Promise<void>
  continueCurrent(client, cfg, additionalIterations = 50): Promise<void>
  stop(): void                                   // aborts + clears pending resolvers
  resolveConfirmation(decision: 'apply' | 'cancel'): void
  resolveChoice(selection: string[] | ChoiceSelection | 'dismiss'): void
  isRunning(): boolean                           // abortController !== null
}
```

A module singleton (`agent/index.ts`): `getAgentLoop(registry, executor, snapshots)` creates once and reuses.

### 11.2 Session start / follow-up

`start(instruction, scope, client, cfg, selection?)`:
1. Create `AbortController`.
2. Compute web-search wiring. **The keyless bundle is baked in**: `web_search` and `fetch_url` are always part of the run's tool set, with the keyless bundle as the default `web_search` backend. The wiring below only decides whether *keyed/native* search is additionally active:
   - `webProvider = appConfig.webAccess.provider` (keyed provider or `'none'`); `byokReady = webProvider !== 'none' && isSearchProviderReady(webProvider)`.
   - `searchToggle = resolveSearchToggle({provider, model, byokReady})`.
   - `keyedSearchEnabled = store.webSearchEnabled && searchToggle.available` — this becomes `session.webSearchEnabled` and means "keyed BYOK or native search is on for this run"; it does **not** gate the keyless bundle.
3. Build a fresh `AgentSession` (ULID id, `status:'building'`, `iteration:0`, `maxIterations:50`, `tokenBudget:{used:0, window:cfg.contextLimits.maxContextTokens}`, `webSearchEnabled: keyedSearchEnabled`, zero totals). `setSession`, `resetSessionTotals`.
4. Append a `UserMessage` with the instruction and the submitted selection, when available.
5. Run `loop(...)`; on throw: aborted → status `'stopped'`; else status `'error'` with `lastError {code:'LoopError'}` and an error `SystemNoticeMessage` ("Run failed: <msg>"). Finally clear the abort controller.

`followUp`: if no current session → `start`. If current status is active → return (ignore). Otherwise reuse the same session id: reset `iteration:0`, `maxIterations:50`, refresh provider/model/scope/webSearchEnabled/tokenBudget, clear pendingChange/pendingChoice/stopReason/lastError; append the user message with its submitted selection; run the same loop with the same error handling.

`continueCurrent`: only valid when `status === 'done' && stopReason === 'max_iterations'`. Bumps `maxIterations += 50`, clears stopReason, appends info notice "Continuing for 50 more iterations.", re-enters loop (iteration resumes from `session.iteration`).

### 11.3 Main loop

Once per run, compute:
- `toolSpecs = [...executor.getToolSpecs(), REQUEST_USER_CHOICE]`. **No web-tool filtering**: `web_search` (keyless-backed by default) and `fetch_url` are keyless-capable and always advertised, whatever the Search toggle or Settings say. There is no `tool-filter` module in this design.
- `parallelizable` = names of specs with `!mutating && runtime === 'none'` excluding `request_user_choice` (network reads can run concurrently; the choice tool blocks on the user).
- `allowedTools` = set of all advertised names. Calls to non-advertised tools fail fast with a corrective `ValidationError` (see 11.5) — with web tools always advertised, this guard now only catches hallucinated tool names.

Keyed/native search layers on top of the always-on keyless baseline: when `session.webSearchEnabled` is true, the `web_search` handler routes to the configured keyed provider instead of the keyless bundle (§16.3), and native-tier providers additionally get their server-side search patch injected via `LLMRequest.nativeSearch` (§10, §12).

Iteration body (`for iter = session.iteration; iter < maxIterations; iter++`):
1. Abort check. Update session `{iteration: iter+1, status:'building'}`.
2. Collect this session's messages from the store; `req = ctxBuilder.build(session, messages, toolSpecs, cfg)`; status `'calling_llm'`.
3. `sr = await stream(...)` (§11.4).
4. Record a `UsageRecord` (turnIndex = iter, cost from `findPricing`/`computeCost`, `pricingVersion: entry ? 'bundled' : 'default'`, `estimated: input===0 && output===0`, `toolCallsCount`). Add tokens into `session.totals` (read the freshest totals from the store when the current session matches).
5. Map streamed tool calls to `ToolCall[]`, resolving each `mutating` flag from the spec (default false). Finalize the streaming assistant message: `updateMessage(streamMsgId, {toolCalls, finishReason})`.
6. Termination: no calls **or** finishReason `'stop'` → status `'done'`, return. finishReason `'length'` → warn notice "Response cut off at token limit — the model may not have finished.", status `'done'`, return.
7. Status `'executing_tool'`; `executeCalls(...)`; next iteration.

If the loop exhausts `maxIterations`: status `'done'`, `stopReason:'max_iterations'`, warn notice "Stopped after 50 iterations. The task may be incomplete." (UI then offers a **Continue** button.)

### 11.4 Streaming into the transcript

`stream(session, client, req, signal)`:
- Immediately append an empty `AssistantMessage` ("streaming message") and set status `'parsing'`.
- Iterate `client.chat(req, signal)`:
  - `text-delta` → append to accumulated text and `updateMessage` live (UI streams).
  - `tool-call-start` / `tool-call-delta` → index-keyed accumulator `{id, name, argsBuf}`.
  - `usage` → capture input/output tokens.
  - `done` → capture finishReason.
  - `error` → **throw** `Error("<code>: <message>")` (caught by start/followUp).
- After the stream: parse each accumulator's `argsBuf` with JSON.parse (on failure, args = `{}` but keep `rawArgs`). For each tool call, append a `ToolCallMessage` (status `'pending'`, mutating placeholder false).
- Return `{streamMsgId, text, toolCalls: [{id,name,args,rawArgs}], finishReason, inputTokens, outputTokens}`.

### 11.5 Tool execution

`executeCalls(calls, session, signal, parallelizable, allowedTools)` — walk the call list; find maximal runs of ≥2 consecutive parallelizable calls and run them with `Promise.all` via `executor.execute` (results appended in call order to keep the transcript deterministic; batch members are always allowed because parallelizable ⊆ advertised specs). Everything else goes one-by-one through `executeCall`.

`executeCall(call, session, signal, allowedTools)` in order:

1. **Kimi `$web_search` echo**: if `session.provider === 'kimi' && session.webSearchEnabled && call.name === '$web_search'` → append success `ToolResultMessage` whose `data` is `call.rawArguments ?? JSON.stringify(call.arguments)`. Return.
2. **Not advertised**: append failure result with `code:'ValidationError'` and message `'Tool "<name>" is not available in this session. Use only the tools in your tool list.'` Return. (Web tools are always advertised in this design, so this only catches hallucinated names.)
3. **`request_user_choice`**: parse via `parsePendingChoice` (§11.6); parse errors → ValidationError result. Set session `{status:'awaiting_choice', pendingChoice}`; await `waitForChoice(signal)` (promise resolved by `resolveChoice`, rejected with `AbortError` on abort); then `{status:'executing_tool', pendingChoice: undefined}`.
   - dismiss → failure `{code:'PermissionDenied', message:'User dismissed the choice menu'}`.
   - select → success `data = {selected_ids, selected_options (full ChoiceOption objects), other_text?}` (other_text only when non-empty trimmed). Return.
4. **Non-mutating** → `executor.execute(call, scope)`, append result. Return.
5. **Mutating** → snapshot → diff → confirm → apply:
   - `sheet = args.sheet`, `snapshotAddress = args.address ?? args.target_address`.
   - If both present: `snapshots.captureRange(sessionId, workbookId, sheet, snapshotAddress, runner)`. If `args.values` is present, `diff = computeRangeDiff(snapshotAddress, snap.before.values ?? [], proposed)`; else `[]`. Set session `{status:'awaiting_confirmation', pendingChange:{id: ulid(), toolCall, snapshotId, diff, severity: diff.length > 50 ? 'elevated' : 'normal', workbookName: registry.getManifest().workbooks[0]?.name ?? 'Workbook', sheet}}`.
   - Else (no addressable range, e.g. `add_sheet`): just `{status:'awaiting_confirmation'}` with no pendingChange payload.
   - Snapshot capture failure → failure result `{code:'OfficeApiError', message:'Snapshot failed: <msg>'}`. Return.
   - Decision: `appConfig.autoApproveSession ? 'apply' : await waitForConfirmation(signal)`. Then `{status:'executing_tool', pendingChange: undefined}`.
   - cancel → failure `{code:'PermissionDenied', message:'User cancelled the write.'}`. Return.
   - apply → `executor.execute`; attach `snapshotId` if one was captured.
   - **Structural snapshots**: if result ok, no range snapshot, tool is `create_chart`/`create_pivot`, and `result.data.name` is a string → `snapshots.captureStructural(sessionId, workbookId, sheet ?? '', 'chart'|'pivot', data.name, {action: call.name})`; attach its id. (Undo then deletes the created object.)
   - Append the result message.

### 11.6 `request_user_choice` (`agent/choice.ts`)

Injected pseudo-tool so the model asks structured questions instead of prose option menus.

```ts
export interface ChoiceOption { id: string; label: string; description?: string; requiresText?: boolean; }
export interface PendingChoice { id: string; toolCallId: string; question: string;
                                 options: ChoiceOption[]; allowMultiple: boolean; }
```

Spec: name `request_user_choice`, non-mutating, runtime `'none'`, params `{question: string (required), options: array (required), allow_multiple?: boolean}`, `additionalProperties: false`. Description (verbatim): *"Show the user a menu and wait for their selection. Use when the request is ambiguous, when fetching everything would be large, or whenever you would otherwise ask the user to choose Option A/B/C in prose. Derive options from information you actually found; put the short option title in label and the tradeoff/details in description. Always include an 'Other' option so the user can specify custom requirements."*

`parsePendingChoice(toolCallId, args)` validation (throws `ToolValidationError`):
- question: string, trimmed non-empty, ≤ 200 chars.
- options: array of 2–8; each item is either a bare string (id = slugified label: lowercase, non-alphanumerics→`-`, trimmed, ≤60 chars, fallback `option-<n>`) or `{id, label, description?, requiresText?}` with non-empty trimmed id/label; invalid items dropped; ≥2 valid required.
- `ensureOtherOption`: if an option with id/label `other`/`others` (case-insensitive) exists, force `requiresText: true` on it; otherwise append `{id:'other', label:'Other', description:"I'll use your custom requirement or scope.", requiresText:true}`.
- allow_multiple: boolean if present, default false.
- Result id: `choice_<toolCallId>`.

---

## 12. Context builder & compaction

`src/agent/context-builder.ts`. Constants: `CHARS_PER_TOKEN = 4` (estimator: `ceil(len/4)`), `MAX_TOOL_RESULT_CHARS = 24_000` (sized above fetch_url's 20k text cap + JSON overhead).

During transcript normalization, a user message with selection metadata becomes:

```text
<visible user text>

<current_selection>
{"sheet":"<sheet name>","address":"<local A1 address>"}
</current_selection>
```

This provider-facing augmentation does not change the visible chat bubble. Each message carries its own immutable-at-submission selection; older messages are never rebound to the workbook's current selection. The workbook manifest is then appended to the first normalized user message as described below.

```ts
export class ContextBuilder {
  constructor(registry: WorkbookRegistry, getReaderFallback: () => boolean = () => false) {}
  build(session, messages, tools, cfg): LLMRequest
  estimateInputTokens(session, messages, tools, cfg): number
}
```

`build`:
1. `system = buildSystemPrompt(workbookId, {keyedSearchEnabled: session.webSearchEnabled, readerFallback})` (§13).
2. `manifestStr = "\n\n<workbook_manifest>\n" + JSON.stringify(manifest, null, 2) + "\n</workbook_manifest>"`.
3. Budget: `budget = cfg.contextLimits.maxContextTokens`, `maxOutput = cfg.maxOutputTokens ?? 4096`, `fixedTokens = estimateTokens(system + manifestStr + JSON.stringify(tools))`.
4. **Normalize** the transcript (`toNormalized`): user→user; assistant→assistant with mapped toolCalls (also record id→name); `tool_call` messages only feed the id→name map; tool results → `{role:'tool', toolCallId, name, content}` where content = `data` string as-is when the tool is `$web_search` and data is a string, else `JSON.stringify(data)` for ok / `JSON.stringify(error)` for failures; **truncate at 24,000 chars** with suffix `... (<n> chars truncated)`. Skip UI-only roles.
5. **Compact** (below).
6. Append `manifestStr` to the **first user message's content** (so the manifest travels inside the conversation).
7. Return `{model, messages, tools, nativeSearch: session.webSearchEnabled ? getNativeSearchCapability(provider, model) : undefined, system, temperature, maxTokens: maxOutput}`.

**Compaction algorithm** `compact(history, fixedTokens, budget, maxOutput)`:
- `available = budget - fixedTokens - maxOutput`; if ≤ 0 return last 2 messages (emergency floor).
- Identify the **two most recent** tool-result messages (they hold live working data).
- `squash(msgs, oldLimit, recentLimit)`: tool messages get their content sliced to their limit + `'…[truncated]'` if longer.
- Pass 1 (unconditional, age-based): `squash(history, 2000, MAX_SAFE_INTEGER)` — collapse *older* tool results to 2,000 chars even when everything fits (prevents unbounded growth on huge-context models). If it fits (`estimateTokens(JSON.stringify(msgs)) <= available`) return.
- Pass 2 (pressure): `squash(prev, 200, 2000)` — old results to 200, recent to 2,000. If fits, return.
- Pass 3 (drop pairs): keep the first user message; from the remaining tail, while length > 4 and over budget, drop the oldest user-to-next-user span (find first user in tail; drop through the next user message; if no user found drop 2; if no next user keep last 4). Return `[firstUser, ...tail]`.

---

## 13. System prompt

`src/agent/system-prompt.ts`. `buildSystemPrompt(workbookId, web: {keyedSearchEnabled, readerFallback})` renders (reproduce verbatim). Web tools are always present in this design, so the web rules are **always included**; `keyedSearchEnabled` only selects the search-scope rule variant below:

```
You are SheetClaw, an AI workbook assistant embedded in Microsoft Excel via an Office Add-in. You help users read, analyse, and edit their workbook data by calling the tools provided to you.

## Rules — follow these strictly

<numbered list of rules below>

## Workflow

<bulleted workflow below>

When you have finished all requested changes and confirmed they succeeded (via tool results), give a brief summary of what was done.
```

Rules (numbered 1..n in order):
1. `**Read before writing.** Always call `read_range` or `get_sheet_context` before writing to any range. Never assume what is in a cell.`
2. `**Never fabricate addresses.** Only reference addresses you have verified via a tool call or received in `current_selection` metadata.`
3. `**Use the submitted selection.** `current_selection` metadata is the Excel selection captured when that user message was submitted. Resolve phrases such as "this cell", "here", and "the selected range" against that sheet and address. Read the stated range before changing it; do not substitute a later selection.`
4. `**One logical change per write.** Make small, targeted edits. If multiple ranges need changes, write them one at a time.`
5. ``**Active scope.** Your active workbook is `<workbookId>`. Only operate on this workbook unless the user explicitly asks you to switch.``
6. `**Announce before mutating.** Briefly explain what you intend to change before calling a write tool (e.g. "I'll write the totals into column D.").`
7. `**Do not claim success prematurely.** A write is not done until you receive a successful tool result. The user must confirm before the write is applied.`
8. `**Use only listed tools.** Do not invent tool names. If a task requires a capability not in your tool list, say so.`
9. *(web rules — see below, always included)*
10. `**Never ask option menus in prose.** If you are about to write "Option A/B/C", "Which option would you like?", "choose one", or any similar menu, stop and call `request_user_choice` instead. Put the option title in `label` and the tradeoff/details in `description`.`

Web rules — always five rules. The first is the **search-scope rule**, with a variant per `keyedSearchEnabled`:
- keyed search OFF (keyless only): `**Search scope.** `web_search` is backed by keyless public catalogues only (Wikipedia, Wikidata, World Bank, IMF, Eurostat, ECB, UN SDG, CKAN, data.gov.my, data.gov.sg, Open-Meteo) — use the `source` parameter to route each query to the best catalogue. General internet search is NOT available in this session; if the task needs it, tell the user they can configure a search provider in Settings → Web Access and enable Search in Chat.`
- keyed search ON: `**Search scope.** `web_search` runs on the user's configured search provider and covers the general internet; the keyless public catalogues remain available through the same tool. Prefer specific queries; searches may be billed to the user's key.`

Then the four rules below:
- `**External data workflow.** When web tools are available and the user asks for external data, search first, then read previews before full fetches. Never paste large raw payloads into your reply; write useful data to the workbook with tools.`
- `**Clarify scope structurally.** Before fetching external data in full, if the request could map to more than one distinct source, table, or granularity, or a preview shows more data than the task needs, call `request_user_choice` with options built only from information you actually found. Do not enumerate those options as plain text.`
- `**Do not browse by trial and error.** If a `fetch_url` preview is truncated, or a plausible public site cannot be fetched because of network or CORS limits, do not keep trying unrelated URLs. Use `request_user_choice` when there are multiple found sources, endpoints, tables, or narrowing strategies that could satisfy the request.`
- CORS reality rule, two variants:
  - readerFallback ON: `**Browser/CORS/proxy reality — read before using web tools.** You run in a browser taskpane; many servers block cross-origin requests (CORS). `fetch_url` automatically retries any failed request through a reader proxy — you do **not** control this fallback and cannot avoid it. Never retry a URL to "try without the proxy"; the proxy is always used automatically on failure. The runtime also caches CORS-blocked hosts and fast-fails repeat calls to them, so retrying a blocked host is always wasted. If a fetch fails or a preview is too short, switch to a different source or call `request_user_choice`.`
  - readerFallback OFF: `**Browser/CORS reality — read before using web tools.** You run in a browser taskpane; many servers block cross-origin requests (CORS), and there is NO automatic proxy fallback in this session, so `fetch_url` on ordinary web pages will usually fail. Prefer `web_search` results and JSON/CSV APIs that allow browser requests. The runtime caches CORS-blocked hosts and fast-fails repeat calls, so retrying a blocked host is always wasted. If a fetch fails, switch to a different source, call `request_user_choice`, or tell the user they can enable the reader-proxy fallback in Settings → Web Access.`

Workflow bullets:
- `To understand the workbook, call `list_sheets` then `get_sheet_context` for relevant sheets.`
- `To read data, call `read_range` with a specific address.`
- `To change workbook content or presentation, use the specific listed tool: `write_range` for values, `format_range` for cell styling/autofit, table tools for Excel tables, sort/filter tools for ordering and filtering, sheet tools for worksheet structure, validation/conditional-format tools for rules, chart/pivot tools for summaries and visuals, and shape/comment/protection/page-layout tools when needed. The user will review and confirm mutating changes before they are applied.`
- `To bring in external data, call `web_search` for discovery and `fetch_url` for bounded previews/full reads.` *(always included — web tools are always listed)*
- `To undo, the user clicks the Undo button in the add-in.`

---

## 14. Workbook layer: registry, executor, snapshots

### 14.1 WorkbookRegistry (`workbook/registry.ts`)

Host-only model. `refresh(runner)`: generate a stable `hostId` (ULID) once per session; inside the runner load `workbook.name` and `worksheets` (`name, position, visibility`); build a single `WorkbookHandle` (`isActive: true, isHost: true, capability: 'host-only'`, `visible = visibility === Excel.SheetVisibility.visible`); store in a map; first refresh sets `activeId`. `resolve(id)` throws `WorkbookNotFoundError` (message `Workbook not found: "<id>"`). `setActive(id)` (resolves first), `getManifest()` → `{active, workbooks}`, `getHostId()`, `getActiveId()`.

### 14.2 ToolExecutor (`workbook/executor.ts`)

```ts
export type ToolHandler = (args: Record<string, unknown>, ctx: Excel.RequestContext,
                           registry: WorkbookRegistry) => Promise<unknown>;
```

Error classes: `ToolValidationError`, `ToolNetworkError` (with `details?`), plus `ToolUnsupportedError` (re-exported from `workbook/unsupported-error.ts` to avoid import cycles).

`register(spec, handler)` into a Map; `getToolSpecs()`.

`execute(call, scope)`:
1. Unknown tool → `{ok:false, error:{code:'ValidationError', message:'Unknown tool: "<name>"'}}`.
2. Lightweight arg validation against the spec: every `required` field present (not undefined/null) — message `Missing required argument: "<field>"`; then for present fields with a declared scalar `type`, check `string`/`number`/`boolean`/`array` — message `"<key>" must be a <type>`. (Deliberately not full JSON Schema — enough for the model to self-correct.)
3. If `workbook_id` present in args, `registry.resolve` it → `WorkbookNotFound` error result on failure.
4. Run the handler: `runtime === 'none'` → call directly with `ctx = undefined`; else wrap in the injected `ExcelRunner`.
5. Success → `{toolCallId, ok:true, data, durationMs}`.
6. Catch mapping — check `instanceof` **or** `e.name` (robust across Vitest module boundaries): ToolValidationError→ValidationError, ToolUnsupportedError→Unsupported, ToolNetworkError→NetworkError, WorkbookNotFoundError→WorkbookNotFound, everything else→OfficeApiError with the message.

### 14.3 SnapshotManager (`workbook/snapshot.ts`)

In-memory Map of `SnapshotEntry`.

- `captureRange(sessionId, workbookId, sheet, address, runner)`: load `values, formulas, numberFormat` of the range; entry kind `'range'`, fidelity `'full'`.
- `captureStructural(sessionId, workbookId, sheet, kind: 'chart'|'pivot', target, definition)`: synchronous; `before = {definition}`, fidelity `'structural-coarse'`.
- `undo(snapshotId, runner)`: throws if missing or already undone.
  - chart + `definition.action === 'create_chart'` → delete `worksheets.getItem(sheet).charts.getItem(target)`.
  - pivot + `create_pivot` → delete `workbook.pivotTables.getItem(target)`.
  - range → restore `formulas` if captured (covers values too), else `values`; then `numberFormat`. Mark `undone: true`.
- `get(id)`, `list(sessionId)`, `lastUndoable(sessionId)` (last non-undone entry in insertion order).

### 14.4 A1 utilities (`workbook/a1notation.ts`)

- `colIndexToLetter(col)` 0-based → `A..Z, AA..`; `cellAddress(col, row)` → e.g. `C5`.
- `stripSheetPrefix('Sheet1!A1:B2')` → `'A1:B2'`.
- `parseRangeTopLeft(address)` → `{col, row}` 0-based (regex `^([A-Z]+)(\d+)$` on the part before `:`; throws on mismatch).
- `computeRangeDiff(rangeAddress, before[][], after[][])`: iterate the max row/col extent; normalize missing cells to `null`; strict `!==` comparison; produce `CellDiff{address, before, after}` with absolute cell addresses offset from the range's top-left.

### 14.5 Layer factory (`workbook/index.ts`)

`createWorkbookLayer(): {registry, executor, snapshots}` — constructs the three objects and registers **all** workbook tools (spec→handler pairs listed in §15). The taskpane wraps this in a singleton (`taskpane/workbookLayer.ts`):

```ts
getTaskpaneWorkbookLayer()  // creates once; additionally registers the web tools:
  executor.register(FETCH_URL, createFetchUrlHandler({
    readerFallback: () => store.appConfig.webAccess.readerFallback }));
  executor.register(WEB_SEARCH, createWebSearchHandler({
    // Effective backend: the configured keyed provider only when the Search
    // toggle is on and the provider is ready; otherwise the keyless bundle.
    getProvider: () => {
      const p = store.appConfig.webAccess.provider;
      return p !== 'none' && store.webSearchEnabled && store.isSearchProviderReady(p)
        ? p : KEYLESS_BUNDLE_ID;
    },
    getApiKey: id => getAuthCredential(store.searchAuthStates[id]),
    getBaseUrl: () => store.appConfig.webAccess.baseUrl,
    getEngineId: () => store.appConfig.webAccess.engineId }));
getTaskpaneAgentLoop()      // getAgentLoop(registry, executor, snapshots)
```

---

## 15. Tool catalogue

Shared arg-parsing helpers (`tools/args.ts`), all throwing `ToolValidationError`:
`stringArg` (non-empty string), `optionalStringArg` (undefined/null/'' → undefined), `numberArg`/`optionalNumberArg` (finite), `optionalBooleanArg`, `stringArrayArg`/`optionalStringArrayArg`, `matrixArg` (2-D array), and `enumArg(args, key, allowed, fallback?)` which **normalizes** input by lowercasing and converting spaces/hyphens to underscores before matching (error: `Invalid "<key>" value "<v>". Supported: ...`).

Conventions: every workbook tool takes `workbook_id` (validated by the executor); range tools take `sheet` + `address` (A1). All mutating specs say "Requires user confirmation." in the description. Handlers run **after** snapshot+confirmation and contain no confirmation logic themselves. Handlers `load(...)` then `await ctx.sync()` per Office.js batching.

### 15.1 Read tools (`tools/specs.ts` + `tools/range.ts`, `tools/workbook_tools.ts`) — all non-mutating

| Tool | Required params | Optional | Behavior |
|---|---|---|---|
| `read_range` | workbook_id, sheet, address | `include: ('values'\|'formulas'\|'numberFormat'\|'text')[]` default `['values']` | Loads rowCount/columnCount/address + requested props. **Rejects ranges > 10,000 cells** (`ValidationError`: "Range is too large (N cells, max 10000). Narrow the address and try again."). Returns `{address, rowCount, colCount, values?, formulas?, numberFormat?, text?}`. |
| `list_sheets` | workbook_id | | `[{name, position, visible}]`. |
| `get_sheet_context` | workbook_id, sheet | `sample_rows` default 5, max 20 | Uses `getUsedRangeOrNullObject`. Empty sheet → `{usedRange:null, headers:null, sampleValues:[]}`. Else loads first `min(sample_rows+1, usedRows)` rows via `getRangeByIndexes`; returns `{usedRange:{address,rowCount,colCount}, headers: row0, sampleValues: rows 1..}`. |
| `get_selection` | workbook_id | | `getSelectedRange()` → `{sheet, address, rowCount, colCount, values}`. |
| `list_workbooks` | — | | Returns the registry manifest. |
| `get_active_workbook` | — | | `{workbook_id, name}` of the active handle (nulls if none). |
| `set_scope_workbook` | workbook_id | | `registry.setActive(id)`; returns `{workbook_id}`. |
| `get_named_ranges` | workbook_id | | `workbook.names` → `[{name, refersTo (formula), type, scope, comment?}]`. |

### 15.2 Write/format tools (`tools/write.ts`) — all mutating

**`write_range`** (workbook_id, sheet, address, values; optional `as_text` default false). Validates `values` is 2-D. Loads the range's dimensions and **requires an exact match** with the values array (error message spells out both shapes: `Dimension mismatch: range <addr> is R×C but values are r×c. Adjust address or values.`). If `as_text` → assign `range.values`; else assign `range.formulas` (handles both `=FORMULA` strings and plain values; nulls become `''`). Returns `{address, written:{rows, cols}}`.

**`clear_range`** (workbook_id, sheet, address; optional `apply_to`: `'contents'` (default) | `'formats'` | `'all'`) → `range.clear(mapped Excel.ClearApplyTo)`.

**`copy_range_format`** (workbook_id, sheet, source_address, target_address; optional `copy_column_width` default true). Loads both ranges; **requires identical shape**; `target.copyFrom(source, Excel.RangeCopyType.formats)`; optionally copies `format.columnWidth`. Note the loop uses `target_address` as the snapshot key in the agent loop (mutating tools snapshot `address ?? target_address`).

**`format_range`** (workbook_id, sheet, address + any of): `number_format`, `bold`, `italic`, `font_color`, `fill_color`, `font_size`, `horizontal_alignment` (enum general/left/center/right/fill/justify/center_across_selection/distributed), `vertical_alignment` (top/middle/center/bottom/justify/distributed — `middle` maps to Center), `wrap_text`, `border_style` (none/continuous/dash/dashed/dash_dot/dash_dot_dot/dot/dotted/double), `border_color`, `border_weight` (hairline/thin/medium/thick), `column_width`, `row_height`, `autofit_columns`, `autofit_rows`. Requires ≥1 option (`Provide at least one formatting option`). numberFormat is applied as a full matrix of the same code. Borders: apply to all six `BorderIndex` values (EdgeTop/Bottom/Left/Right, InsideVertical/InsideHorizontal); style defaults to `Continuous` if only color/weight given. Returns `{sheet, address, formatted: string[] of applied props}`.

### 15.3 Chart tools (`tools/charts.ts`)

Chart type allow-list (else `ToolUnsupportedError` "Unsupported chart_type ..."): column→columnClustered, bar→barClustered, line, pie, area, scatter→xyscatter, doughnut, radar. `series_by`: rows/columns/auto.

| Tool | Mutating | Params (req; opt) | Behavior |
|---|---|---|---|
| `list_charts` | no | workbook_id, sheet | `[{name, chartType}]` |
| `create_chart` | yes | workbook_id, sheet, chart_type, data_range; title, series_by | `sheet.charts.add(type, range, seriesBy)`; optional title (visible). Returns `{name, created:true}` → triggers structural snapshot. |
| `modify_chart` | yes | workbook_id, sheet, chart_name; title, chart_type, data_range, series_by | Applies present props; `chart.setData` for range. Returns `{name, applied[]}` |
| `delete_chart` | yes | workbook_id, sheet, chart_name | delete |
| `set_chart_data` | yes | workbook_id, sheet, chart_name, data_range; series_by | `chart.setData` |
| `format_chart` | yes | workbook_id, sheet, chart_name; title, show_title, show_legend, legend_position (Top/Bottom/Left/Right/Corner), style (number), left, top, width, height | Applies present props (style/position via untyped access for older typings). Returns `{name, applied[]}` |
| `set_chart_axes` | yes | workbook_id, sheet, chart_name; category_axis_title, value_axis_title, category_axis_visible, value_axis_visible, value_axis_number_format | Axis titles set text + visible=true |
| `set_chart_labels` | yes | workbook_id, sheet, chart_name; show_value, show_category, show_series_name, show_percentage, position | `chart.dataLabels.*` |
| `add_trendline` | yes | workbook_id, sheet, chart_name; series_index (default 0), trendline_type (default Linear; Exponential/Logarithmic/MovingAverage/Polynomial/Power), name, show_equation, show_r_squared | `series.getItemAt(i).trendlines.add(type)`; returns trendline JSON |

### 15.4 Pivot tools (`tools/pivots.ts`)

Guard `assertPivotApi(ctx)`: if `!('pivotTables' in ctx.workbook)` throw `ToolUnsupportedError('Pivot tables require ExcelApi 1.8+')`. Aggregations: sum (default), count, average, max, min, product.

| Tool | Mutating | Params | Behavior |
|---|---|---|---|
| `list_pivots` | no | workbook_id; sheet | Per-sheet or workbook-wide `[{name, sheet?}]` |
| `get_pivot` | no | workbook_id, name | `{name, availableFields (hierarchies), rows, columns, data:[{name, summarizeBy}], filters}` |
| `create_pivot` | yes | workbook_id, sheet, source_range, destination; dest_sheet, name | `destSheet.pivotTables.add(name ?? '', sourceRange, destRange)` → `{name, created:true}` → structural snapshot |
| `add_pivot_field` | yes | workbook_id, name, field, area (row/column/data/filter); function | Loads hierarchies, validates field exists (error lists available), adds to the area collection; data area sets `summarizeBy` |
| `refresh_pivot` | yes | workbook_id, name | `pivot.refresh()` |
| `remove_pivot_field` | yes | workbook_id, name, field, area | Finds item in area collection (error `Field "<f>" is not in the <area> area.`); `collection.remove(item)` or `item.delete()`, else Unsupported |
| `set_pivot_style` | yes | workbook_id, name, style | `pivot.style = style` |
| `set_pivot_layout` | yes | workbook_id, name; layout_type (Compact/Outline/Tabular), show_row_grand_totals, show_column_grand_totals, show_banded_rows, show_banded_columns | Applies present props |
| `refresh_all_pivots` | yes | workbook_id | `workbook.pivotTables.refreshAll()` |
| `delete_pivot` | yes | workbook_id, name | delete |

### 15.5 Table tools (`tools/tables.ts`)

| Tool | Mutating | Params | Behavior |
|---|---|---|---|
| `list_tables` | no | workbook_id; sheet | `[{name, sheet, address, style, showHeaders, showTotals, showBandedRows, showBandedColumns, showFilterButton}]` (two-sync pattern: load tables, then load each `getRange().address`) |
| `create_table` | yes | workbook_id, sheet, address; has_headers (default true), name, style | `sheet.tables.add`; optional rename/style; returns `{name, address, created:true}` |
| `resize_table` | yes | workbook_id, table, address | `table.resize(address)` |
| `add_table_rows` | yes | workbook_id, table, values; index, always_insert (default true) | `table.rows.add(index, values, alwaysInsert)` |
| `add_table_columns` | yes | workbook_id, table, name; values, index | `table.columns.add(index, values ?? [[name]], name)` |
| `set_table_style` | yes | workbook_id, table; style, show_headers, show_totals, show_banded_rows, show_banded_columns, show_filter_button | Applies present props |
| `set_table_totals` | yes | workbook_id, table, column; function (none/sum/average/count/count_numbers/min/max/standard_deviation/variance/custom), label, formula | Forces `showTotals=true`; label → total cell value; formula → total cell formula; else builtin → `=SUBTOTAL(<code>,[<column>])` with codes {average:101, count:103, count_numbers:102, max:104, min:105, standard_deviation:107, sum:109, variance:110} |
| `delete_table` | yes | workbook_id, table; convert_to_range (default true) | `convertToRange()` or `delete()` |

### 15.6 Sort/filter tools (`tools/sort_filter.ts`)

Sort fields parse: non-empty array of objects with `key` (number 0-based column index for ranges; name or index for tables), `ascending` default true, `dataOption` `'text_as_number'` → `TextAsNumber` else `Normal`.

| Tool | Mutating | Params | Behavior |
|---|---|---|---|
| `sort_range` | yes | workbook_id, sheet, address, fields; has_headers (default true), match_case (default false), orientation (rows/columns, default rows) | `range.sort.apply(fields, matchCase, hasHeaders, orientation)` |
| `sort_table` | yes | workbook_id, table, fields; match_case | `table.sort.apply` |
| `apply_filter` | yes | workbook_id, column; sheet, address, table, criterion1, criterion2, operator (and/or/top_items/bottom_items/top_percent/bottom_percent/values/dynamic), values (string[]), filter_on (custom/values/top_items/bottom_items/top_percent/bottom_percent/dynamic/cell_color/font_color/icon), dynamic_criteria | Builds `Excel.FilterCriteria` (filterOn defaults to `values` if a values list given else `custom`; snake_case→camelCase conversion for filterOn/dynamicCriteria; operator maps and/or → And/Or). Table path: `table.columns.getItem(column).filter.apply`. Range path: requires numeric column; `worksheet.autoFilter.apply(address, column, criteria)` |
| `clear_filters` | yes | workbook_id; sheet, address, table | Table: `table.autoFilter.clearCriteria()`; else `worksheet.autoFilter.clearCriteria()` |
| `reapply_filters` | yes | workbook_id; sheet, table | `.autoFilter.reapply()` |

### 15.7 Sheet tools (`tools/sheets.ts`)

| Tool | Mutating | Params | Behavior |
|---|---|---|---|
| `add_sheet` | yes | workbook_id; name, activate (default true) | `worksheets.add(name?)`; returns `{name, position, visible}` |
| `rename_sheet` | yes | workbook_id, sheet, new_name | |
| `delete_sheet` | yes | workbook_id, sheet | |
| `copy_sheet` | yes | workbook_id, sheet; new_name, position (none/before/after/beginning/end, default end), relative_to | `sheet.copy(position, relativeSheet?)` |
| `move_sheet` | yes | workbook_id, sheet, position (number) | `ws.position = n` |
| `hide_sheet` | yes | workbook_id, sheet; visibility (visible/hidden/very_hidden, default hidden) | |
| `activate_sheet` | yes | workbook_id, sheet | |
| `freeze_panes` | yes | workbook_id, sheet, mode (rows/columns/at/none); count (default 1), address | `freezeRows/freezeColumns/freezeAt/unfreeze` |

### 15.8 Data validation tools (`tools/validation.ts`)

**`set_data_validation`** (mutating; workbook_id, sheet, address, type) — type ∈ list/whole_number/decimal/date/time/text_length/custom. Rule construction:
- list → `{list:{source (required), inCellDropDown: show_dropdown ?? true}}`
- custom → `{custom:{formula: formula1 (required)}}`
- others → `{<camelType>: {formula1: formula1 ?? '', formula2?, operator (pascal-cased, default Between)}}` — operators between/not_between/equal_to/not_equal_to/greater_than/less_than/greater_than_or_equal_to/less_than_or_equal_to.
Also: `ignore_blanks` default true; optional prompt (`prompt_title`/`prompt_message` → `{showPrompt:true,...}`); optional error alert (`error_title`/`error_message`, `error_style` stop/warning/information default stop).

**`clear_data_validation`** (mutating) → `.dataValidation.clear()`. **`get_data_validation`** (non-mutating) → load `type,valid,rule,ignoreBlanks,prompt,errorAlert`, return `toJSON()`.

### 15.9 Conditional format tools (`tools/conditional_formats.ts`)

**`add_conditional_format`** (mutating; workbook_id, sheet, address, type ∈ cell_value/contains_text/top_bottom/color_scale/data_bar):
- cell_value → `cf.cellValue.rule = {formula1 (req), formula2?, operator pascal (default GreaterThan)}` + optional fill/font colors.
- contains_text → `cf.textComparison.rule = {operator:'Contains', text (req)}` + colors.
- top_bottom → `cf.topBottom.rule = {rank: rank ?? 10, type:'TopItems'}` + colors.
- color_scale → criteria `{minimum:{type:'LowestValue', color: lowest_color ?? '#F8696B'}, midpoint:{type:'Percentile', formula:'50', color: midpoint_color ?? '#FFEB84'}, maximum:{type:'HighestValue', color: highest_color ?? '#63BE7B'}}`.
- data_bar → `barColor = bar_color ?? '#638EC6'`.
Returns `{sheet, address, id, type}`.

**`list_conditional_formats`** (non-mutating) → `[{id, type, priority}]`. **`clear_conditional_formats`** (mutating) → `conditionalFormats.clearAll()`.

### 15.10 Range structure ops (`tools/range_ops.ts`)

| Tool | Mutating | Params | Behavior |
|---|---|---|---|
| `insert_cells` | yes | +shift (down/right, default down) | `range.insert('Down'/'Right')` |
| `delete_cells` | yes | +shift (up/left, default up) | `range.delete('Up'/'Left')` |
| `insert_rows` / `delete_rows` | yes | workbook_id, sheet, address | `getEntireRow().insert('Down')` / `.delete('Up')` |
| `insert_columns` / `delete_columns` | yes | workbook_id, sheet, address | `getEntireColumn().insert('Right')` / `.delete('Left')` |
| `merge_range` | yes | +across (default false) | `range.merge(across)` |
| `unmerge_range` | yes | | |
| `find_replace` | yes | workbook_id, sheet, find, replace; address, complete_match, match_case | On the range if address given else the worksheet: `replaceAll(find, replace, {completeMatch, matchCase})`; returns `{replaced: count}` |
| `get_special_cells` | **no** | workbook_id, sheet, address, cell_type (blanks/constants/formulas/visible/data_validations/conditional_formats/same_data_validation/same_conditional_format); value_type | `getSpecialCellsOrNullObject(PascalType, PascalValueType?)` → `{address, areaCount}` or `{address:null, areaCount:0}` |

### 15.11 Named range tools (`tools/names.ts`) — all mutating

Collection = `worksheet.names` if `sheet` given else `workbook.names`.
- `create_named_range` (workbook_id, name, reference; comment, sheet, visible default true) → `names.add(name, reference, comment)`, returns item JSON.
- `update_named_range` (workbook_id, name; reference, comment, visible, sheet) — sets `formula`/`comment`/`visible` when present.
- `delete_named_range` (workbook_id, name; sheet).

### 15.12 Comment/note tools (`tools/comments.ts`)

Address helper: prefix `sheet!` unless the address already contains `!`. Comment lookup by `comment_id` else by cell.
- `list_comments` (no; workbook_id; sheet) → `[{id, content, authorName, creationDate, resolved}]`.
- `add_comment` (yes; workbook_id, sheet, address, content) → `workbook.comments.add(fullAddress, content, 'Plain')`.
- `reply_to_comment` (yes; workbook_id, content; comment_id or sheet+address).
- `delete_comment` (yes; workbook_id; comment_id or sheet+address).
- `add_note` (yes; workbook_id, sheet, address, content; visible, width, height) → `worksheet.notes.add`; returns note JSON.
- `delete_note` (yes; workbook_id, sheet, address).

### 15.13 Shape tools (`tools/shapes.ts`)

Common optional position props: left/top/width/height (points), name.
- `list_shapes` (no; workbook_id, sheet) → shape JSONs (`id,name,type,left,top,width,height,visible`).
- `add_image` (yes; +base64_image — no data-URL prefix) → `shapes.addImage(base64)`.
- `add_textbox` (yes; +text) → `shapes.addTextBox(text)`.
- `add_shape` (yes; shape_type default Rectangle; fill_color → `fill.setSolidColor`, line_color → `lineFormat.color`) → `shapes.addGeometricShape`.
- `move_shape` / `resize_shape` / `delete_shape` (yes; shape = name or ID).

### 15.14 Layout & protection tools (`tools/layout_protection.ts`) — all mutating

- `protect_sheet` (workbook_id, sheet; password, allow_sort, allow_filter, allow_format_cells, allow_insert_rows, allow_delete_rows, allow_pivots) → `protection.protect(options, password?)`.
- `unprotect_sheet` (…; password), `protect_workbook` (workbook_id; password), `unprotect_workbook`.
- `set_print_area` (workbook_id, sheet, address) → `pageLayout.setPrintArea`; `clear_print_area` → `setPrintArea('')`.
- `set_page_layout` (workbook_id, sheet; orientation Portrait/Landscape, paper_size, print_gridlines, print_headings, center_horizontally, center_vertically, margin_unit default Inches, top/bottom/left/right_margin, fit_to_width, fit_to_height, scale 10–400) — margins go through `setPrintMargins(unit, margins)`; fit/scale through `layout.zoom = {horizontalFitToPages, verticalFitToPages, scale}`.
- `add_page_break` (workbook_id, sheet, address, direction horizontal/vertical) → `horizontalPageBreaks.add(address)` / `verticalPageBreaks.add`; returns rowIndex/columnIndex.

### 15.15 Registration order

`createWorkbookLayer` registers, in order: 8 read specs, 4 write specs, 9 chart, 10 pivot, 8 table, 5 sort/filter, 8 sheet, 3 validation, 3 conditional-format, 10 range-op, 3 name, 6 comment, 7 shape, 8 layout/protection = **92 workbook tools**, plus `fetch_url` and `web_search` registered by the taskpane layer, plus the loop-injected `request_user_choice`.

---

## 16. Web access layer

### 16.1 Guarded fetch (`web/net.ts`)

Constants: `FETCH_TIMEOUT_MS = 15_000`, `MAX_BODY_BYTES = 1_000_000`, probe timeout 5 000 ms, allowed ports `{'', '80', '443', '8080'}`.

`validatePublicHttpUrl(value)` throws `ToolValidationError` unless: absolute http(s) URL; no embedded credentials; allowed port; hostname not empty/localhost/`*.localhost`/`*.local`; not a private/loopback/link-local IP. IPv4 blocks: `0.*`, `10.*`, `127.*`, `100.64–127.*`, `169.254.*`, `172.16–31.*`, `192.168.*`, `198.18–19.*`. IPv6 blocks: `::1`, `::`, `fe80:*`, `fc*`, `fd*` (brackets stripped, lowercased).

`fetchTextWithGuards(url, opts)`:
- **Session-scoped CORS-blocked host cache** (`corsBlockedHosts: Set<string>`, clearable via `clearHostStatusCache()`). If host is cached, fail fast with `ToolNetworkError`: `Blocked by CORS (cached): <host> was already confirmed CORS-blocked earlier in this session. Do not retry this host; choose a different source or ask the user how to proceed.`
- GET with `redirect:'follow'`, own AbortController + timeout, chained to caller's signal. **Re-validate the final (post-redirect) URL.**
- Stream the body with a byte cap: if `truncateAtMaxBytes` keep the first `maxBytes` and flag `truncated`; else throw `Response body exceeded <n> bytes.`
- Returns `{url, finalUrl, status, statusText, contentType, bytesFetched, truncatedByBytes, text}` (note: non-2xx statuses are returned, not thrown).
- Error taxonomy on fetch throw: caller-aborted → rethrow raw; AbortError → timeout message (`Request timed out after <ms> ms. The host did not respond in time; one retry or a different source is reasonable.`); `TypeError` → run a **no-cors reachability probe** (GET `mode:'no-cors'`, cancel body): reachable → cache host + `Blocked by CORS: <host> is reachable but does not allow browser requests from the add-in. Do not retry this host; choose a different source, or the user can enable the reader fallback in Settings.`; unreachable → `Network unreachable: could not connect to <host> (DNS failure or no connectivity). Do not retry immediately.`; anything else → `Network request failed: <msg>`.

### 16.2 `fetch_url` tool (`web/fetch.ts`)

Spec: non-mutating, runtime `'none'`, `additionalProperties: false`. Params: `url` (required), `mode` preview|full (default preview), `format` auto|text|json|csv (default auto), `max_chars` (clamped 1 000–20 000; defaults: preview 4 000, full 16 000). Description mentions preview-first workflow.

Handler (`createFetchUrlHandler({readerFallback})` — readerFallback may be a boolean or getter):
1. Direct `fetchTextWithGuards(url, {truncateAtMaxBytes:true})`.
2. On `ToolNetworkError` **and** readerFallback enabled: retry via the reader proxy `https://r.jina.ai/<url>` (same guards); result marked `source:'reader'` vs `'direct'`.
3. Format the result. Auto-detection: content-type containing `json` or body starting `{`/`[` → json; content-type `csv` or URL path ending `.csv` → csv; else text.
   - Base fields: `{url, finalUrl, status, contentType, source, format, bytesFetched}`.
   - **text**: HTML gets scripts/styles/tags stripped and whitespace collapsed; then `cleanTruncate` (cut at maxChars, back up to the last `\n`/space/comma if it's ≥ 80% of the cut) → `{returnedChars, truncated, text}`.
   - **json**: parse; unparseable + byte-truncated → `{dataPreview, truncated:true, hint:'JSON response exceeded the network byte cap before it could be parsed; refetch with a narrower API query or a smaller dataset.'}`; unparseable otherwise → `ToolValidationError('Response is not valid JSON.')`. Parseable: pretty-print; fits → `{data}` (structured); truncated → `{dataPreview, hint:'<N array items|N top-level keys>; refetch with a narrower API query or ask the user to choose a scope.'}`.
   - **csv**: RFC-style quoted-field parser (handles `""` escapes, CRLF). Always returns `{headerRow, previewRows (first 20 data rows), approxRowCount}`; `mode:'full'` additionally returns capped `text`.

### 16.3 `web_search` tool (`web/search.ts`)

Spec: non-mutating, runtime `'none'`. Params: `query` (required, ≤400 chars), `max_results` (1–10, default 5), `include_content` (default false), `source` (enum `'auto'` + all keyless source ids, with per-source hint text baked into the description — honoured when the query routes to the keyless bundle, which is the default backend; ignored when a keyed provider handles the query). Description emphasizes preferring it over fetch_url for CORS-likely pages and using `request_user_choice` when results are ambiguous.

Handler (`createWebSearchHandler(options)`): `options.getProvider()` returns the **effective backend** — the configured keyed provider when keyed search is enabled and ready, otherwise the keyless bundle (§14.5). Because the keyless bundle is always a valid fallback, the handler can never be left without a provider. Keyed providers need a key (`Missing API key for <id>.`); call `provider.search(query, {maxResults, apiKey, baseUrl?, engineId?, includeContent, source, signal, fetchImpl?})`; validate the result is an array (`ToolNetworkError('<id> returned an invalid result list.')`); return `{query, provider: id, results: results.slice(0, maxResults)}` so the model can see which backend answered.

### 16.4 Search provider adapter contract (`web/providers/index.ts`)

```ts
export interface SearchResult {
  title: string; url: string; snippet?: string; publishedAt?: string;
  content?: string;              // extracted page text (≤ MAX_RESULT_CONTENT_CHARS = 4000)
  source?: SearchProviderId;     // set by the keyless bundle only
}
export interface SearchProviderAdapter {
  id: SearchProviderId | 'keyless';
  label: string;
  requiresKey: boolean;
  requiresEngineId?: boolean;    // Google CSE
  selfHosted?: boolean;          // SearXNG — excluded from keyless bundle
  endpoint: string; signupUrl: string;
  search(query, opts): Promise<SearchResult[]>;
}
```

Registry `SEARCH_PROVIDERS` (15 providers):

| id | label | key? | endpoint |
|---|---|---|---|
| tavily | Tavily | yes | `https://api.tavily.com/search` |
| google-cse | Google Programmable Search | yes (+engineId) | `https://www.googleapis.com/customsearch/v1` |
| jina | Jina Search | yes | `https://s.jina.ai/` |
| searxng | SearXNG (self-hosted) | no, selfHosted | `http://localhost:8080/search` |
| wikipedia | Wikipedia (keyless, encyclopedic only) | no | `https://en.wikipedia.org/w/api.php` |
| wikidata | Wikidata (keyless entities) | no | `https://www.wikidata.org/w/api.php` |
| worldbank | World Bank Indicators (keyless) | no | `https://api.worldbank.org/v2/indicator` |
| ckan | CKAN catalog (keyless, data.gov.au) | no | `https://data.gov.au/data/api/3/action/package_search` |
| data-gov-my | data.gov.my APIs (keyless) | no | `https://api.data.gov.my/data-catalogue` |
| data-gov-sg | data.gov.sg datasets (keyless) | no | `https://api-production.data.gov.sg/v2/public/api/datasets` |
| ecb | ECB Data Portal (keyless) | no | `https://data-api.ecb.europa.eu/service/dataflow/ECB/ALL/latest` |
| eurostat | Eurostat datasets (keyless) | no | `https://ec.europa.eu/eurostat/api/dissemination/sdmx/2.1/dataflow/ESTAT/all/latest` |
| imf | IMF DataMapper (keyless) | no | `https://www.imf.org/external/datamapper/api/v1/indicators` |
| open-meteo | Open-Meteo (keyless) | no | `https://api.open-meteo.com/v1/forecast` |
| un-sdg | UN SDG API (keyless) | no | `https://unstats.un.org/sdgapi/v1/sdg/Indicator/List` |

Each provider module implements `search()` against its public JSON API and normalizes to `SearchResult[]`; each declares a `signupUrl` for the settings link. Keyed providers pass the key per their API (Tavily: JSON body `api_key`; Google CSE: `key` + `cx` query params; Jina: `Authorization: Bearer`). Providers honour `opts.baseUrl` via `resolveBaseUrl(baseUrl, fallback)` which only accepts absolute http(s) URLs.

**User-selectable vs. internal:** only the keyed/self-hosted providers (tavily, google-cse, jina, searxng) appear in Settings → Search. The keyless sources are internal — reachable exclusively through the always-on keyless bundle (§16.5) and its `source` routing parameter.

Also exported: `READER_PROVIDER_ENDPOINT = 'https://r.jina.ai/'`, host allowlists derived from provider endpoints/signup URLs (used by the CORS integration test), `getSearchProvider(id)` (`'none'`→null, `'keyless'`→bundle), `isKeylessSearchProvider(id)` (adapter exists and `requiresKey === false`).

### 16.5 Keyless bundle (`web/providers/keyless.ts`)

Meta-provider `id: 'keyless'`, label "All keyless sources (bundled)", `requiresKey: false`, empty endpoint/signupUrl. **Not user-selectable**: it is the permanent default backend for `web_search` — used whenever keyed search is off, unconfigured, or not ready — and never appears in the Settings provider list.

- `keylessSourceIds()` = registry ids with `requiresKey === false && !selfHosted` (computed lazily due to an import cycle with index.ts).
- `KEYLESS_SOURCE_HINTS` (surfaced in the web_search schema): wikipedia "encyclopedic articles", wikidata "structured entities and identifiers", worldbank "World Bank development indicators, country-level statistics", ckan "Australian open-data catalogue (data.gov.au)", data-gov-my "Malaysia official open-data APIs including OpenDOSM statistics", data-gov-sg "Singapore official datasets", ecb "European Central Bank euro-area finance and statistics", eurostat "EU official statistics", imf "IMF macroeconomic indicators by country", open-meteo "weather forecasts for coordinates", un-sdg "UN Sustainable Development Goal indicators".
- `search()`: target = the requested `source` if valid, else `AUTO_SOURCES = ['wikipedia','wikidata','worldbank']`. Never forwards `baseUrl`. `Promise.allSettled` across targets; tag each result with its `source`; **round-robin interleave** buckets so every source appears near the top. All-failed: single targeted source rethrows its original error; multi-source throws `ToolNetworkError('All keyless sources failed - <id>: <msg|no results>; ...')`.

---

## 17. Pricing & usage tracking

### 17.1 Pricing (`pricing/index.ts` + `pricing.json`)

Matching: entry matches when (`entry.provider === provider` or `'*'`) and model matches `modelMatch` (`'*'` any; `prefix*` startsWith; else exact). `findPricing` prefers exact (non-glob) matches over globs; returns null if none. `computeCost(record, entry, defaults)`:

```
regularInput = max(0, inputTokens - cacheReadTokens)
cost = regularInput/1e6 * inputRate
     + cacheRead/1e6   * (cacheReadPerMTok ?? inputRate)
     + cacheWrite/1e6  * (cacheWritePerMTok ?? inputRate)
     + outputTokens/1e6 * outputRate
```

Bundled table (version 2026-06-08; defaults input $1.00 / output $3.00 per MTok): openai gpt-4o 2.50/10.00 (cacheRead 1.25), gpt-4o-mini* 0.15/0.60 (0.075), gpt-4.1* 2.00/8.00 (0.50), o3* 10/40, o4-mini* 1.10/4.40; anthropic claude-opus-4* 15/75 (cr 1.50, cw 18.75), claude-sonnet-4* 3/15 (0.30/3.75), claude-haiku-4* 0.80/4.00 (0.08/1.00), claude-3-5-sonnet* 3/15, claude-3-5-haiku* 0.80/4.00; ollama * 0/0; generic deepseek* 0.27/1.10; generic qwen* 0.50/1.50; gemini/cerebras/cloudflare/huggingface * 0/0.

### 17.2 Usage queries (`usage/queries.ts`)

`loadBuckets(fromDay, toDay)` scans localStorage for `xl.usage.day.*` keys within the lexicographic day range. Ranges: today, week (last 7 days), month (last 30), all (from 2000-01-01). Aggregations: `queryTotals` (sums + `estimatedCount`), `queryByProvider`, `queryByModel` (key `provider/model`), `queryByDay` (sorted ascending).

### 17.3 CSV export (`usage/export.ts`)

Header `timestamp,session_id,turn_index,provider,model,input_tokens,output_tokens,cache_read_tokens,cache_write_tokens,total_tokens,estimated,tool_calls` with CRLF line endings, RFC-4180 quoting, **UTF-8 BOM prefix** so Excel opens it correctly. Download via Blob + temporary `<a download="usage-<from>-to-<to>.csv">`.

---

## 18. User interface

### 18.1 Bootstrap (`taskpane/index.tsx`)

Inside `Office.onReady`: `loadConfigFromStorage()`, `loadChatHistory()`, `void loadAuthFromStorage()` (async decrypt; store re-renders when it lands); then `createRoot(#root).render(<App/>)`.

### 18.2 App shell (`App.tsx`)

`FluentProvider` is a full-height flex column using the persisted `appConfig.theme`. `light` and `dark` select Fluent's `webLightTheme` and `webDarkTheme`; `system` resolves `(prefers-color-scheme: dark)` and listens for changes while the app is mounted. The provider also sets the CSS `color-scheme`, foreground, and background from the resolved theme.

The fixed header contains an optional back-to-Chat button (shown on every non-Chat view), 🦞 "SheetClaw", the product caption, and an overflow navigation menu for History, Usage, Settings, and About. The active non-Chat item has a check mark. Selecting Settings from the menu clears any targeted Settings sub-tab; `ChatPanel.onOpenSettings(target?: 'search')` can still navigate directly to a requested sub-tab. `HistoryPanel` receives `onOpenChat`.

Only one view is mounted at a time. Its key is the current view, so changing views remounts the surface and applies the `tab-surface-enter` animation (180 ms fade/6 px rise); `prefers-reduced-motion: reduce` disables the animation. A persistent `Footer` sits below the active surface.

### 18.3 ChatPanel — behavior spec

State/selectors: current session, messages, providers, appConfig, auth states, `webSearchEnabled`, `isProviderReady(active)`, `isSearchProviderReady(webAccess.provider)`, and the currently observed `WorkbookSelection | null`.

Derived:
- `isRunning` = status ∈ {building, calling_llm, parsing, executing_tool}.
- `providerReady` = active provider `enabled` && authenticated && model non-empty; otherwise a warning MessageBar with a Settings button ("No provider enabled..." / "Active provider is not authenticated..." / "Select a model in Settings before chatting.").
- `searchToggle = resolveSearchToggle(...)`; the pill's active state = `webSearchEnabled && searchToggle.available`. (The keyless bundle is always on in the background and has no pill state — the pill strictly reflects keyed/native search.)

Effects: on mount, `registry.refresh()` (errors → error MessageBar); scroll to bottom on new messages; if the toggle is on but search became unavailable, switch it off. When Office/Excel is present, read the selected range immediately and subscribe to `Office.EventType.DocumentSelectionChanged`; remove the same handler on unmount. Selection-read failures clear the badge without surfacing an error, and request ids prevent a slower read from replacing a newer selection.

Composer: if a selection is available, show a compact badge above the input with the sheet and local A1 address. The input is a single auto-growing Textarea (height measured from scrollHeight, clamped 32–200 px; Enter sends, Shift+Enter inserts a newline; disabled while running/awaiting). Action row: **New chat** icon button (stops the loop and clears session, totals, and input), icon-only **Search** pill (🌐 — governs keyed/native internet search only; keyless catalogue search is always available and needs no pill. Clicking when unavailable — no native tier and no ready keyed provider — shows `getUnavailableSearchToggleHint` and forces the toggle off; otherwise toggles), an icon-only **edit approval mode** menu, and Send (▶) / **Stop** while running. The approval menu is radio-like: **Ask before edits** sets `autoApproveSession:false`; **Accept all edits** sets it to `true`. Its trigger icon and tooltip reflect the selected mode.

Send: build the adapter via `createAdapter(cfg, authState)`, refresh the registry, and re-read the Excel selection at submission time (rather than trusting the event-driven badge). Normalize the selection in `taskpane/selection.ts`: load `address,worksheet/name`, remove the worksheet prefix and all absolute-reference `$` characters, and return `{sheet, address}`. Scope is `{workbookId: registry.getActiveId() ?? 'host'}`; pass the captured selection to `followUp` if a session exists, otherwise to `start`.

Session strip (when a session exists): `"<model> | iter i/max | <in+out> tok"` caption + **Continue** button (only when `status==='done' && stopReason==='max_iterations'`; re-creates the adapter from the *session's* provider and calls `continueCurrent`) + **Undo last write** (calls `snapshots.lastUndoable(session.id)` then `snapshots.undo(id, Excel.run)`).

Message rendering:
- Transcript is grouped into render items: user messages break segments; within a segment, all tool activity (assistant `toolCalls`, `tool_call` and `tool` messages) is collapsed into one **ToolCallChain** card — a collapsible "Tool calls (n) — k/n complete[, f failed]" header expanding to per-call rows (`running` cyan / `ok` green / `error` red-with-truncated-message, left accent border, monospace). Assistant messages that only carry tool calls (empty text) are hidden. The chain is inserted after the last assistant message that owns the calls.
- User bubbles right-aligned brand-colored; assistant bubbles left neutral. Assistant text goes through a **minimal markdown renderer**: GitHub-style pipe tables (header + `---` separator rows with `:` alignment, rendered as a horizontally-scrollable table) and `**bold**` inline. Everything else is pre-wrap plain text.
- `system_notice` renders as a MessageBar (info/warning/error). `confirmation` messages render nothing.
- Empty state: "Ready for this workbook" + three example prompt buttons ("Summarize the active sheet", "Sum B2:B13 into B14", "Make a bar chart from A1:B12"), or setup CTA when no provider.
- While running/awaiting-choice: spinner + status label (building→"Preparing context", calling_llm→"Calling model", parsing→"Reading response", executing_tool→"Running workbook tool", awaiting_confirmation→"Awaiting confirmation", awaiting_choice→"Awaiting selection").

**ConfirmationBlock** (when `awaiting_confirmation` with a pendingChange): title `Confirm change - <workbook> / <sheet>`; if severity elevated a red "Large change - review carefully"; first 10 diffs as `ADDR: before → after` (empty rendered `(empty)`, values truncated at 40 chars) with "...and N more cells" overflow, or "(no cell values change)"; **Apply** / **Cancel** → `loop.resolveConfirmation`.

**ChoiceBlock** (when `awaiting_choice`): question + numbered option buttons (label bold, description caption; toggle select, radio vs multi per `allowMultiple`); options with `requiresText` reveal a textarea when selected; **Continue** disabled until a selection exists and any required text is non-empty → `resolveChoice({ids, otherText?})`; **Dismiss** → `resolveChoice('dismiss')`.

### 18.4 HistoryPanel

Lists `chatHistory` (newest first): title, status chip (Done/Error/Stopped/Confirming/Choosing/Running/Idle), preview line, model, `MMM d, h:mm` timestamp | message count. Click → stop loop, `resumeChat(id)`, seed `resetSessionTotals` from the restored session's totals, switch to Chat. Per-item 🗑️ with inline Delete/Cancel confirm; header "Delete all" with the same two-step confirm.

### 18.5 UsageDashboard

Range selector (Today/7d/30d/All) + **Export CSV**. Empty state "No usage yet — start a chat." Otherwise: a Tokens totals card (formatted 1.2k/3.45M), a per-day bar sparkline (title tooltip `day: n tok`), a "By model" table (provider/model | tokens), a "By provider" table when >1 provider. Footer **Reset history** (confirm() then remove all `xl.usage.*` keys).

### 18.6 SettingsPanel

Four sub-tabs — **Ollama**, **API**, **Search**, and **Appearance** — with a `*` marker on the provider-owning tab.

**Provider form** (shared): "Set as active"/"Active provider" button; provider signup link; native-search caption (`getProviderNativeSearchCaption`); Base URL input (commit on blur, sets enabled); Model field — a freeform Combobox when a model list is available (with Refresh) else a plain input; model lists come from `knownModels` → static fallback lists per provider → live `listModels()` fetch. Fetch behavior: auto-fetch on mount when possible (ollama and anthropic don't need a key; others need a stored credential); OpenAI results are filtered to chat models (prefixes `gpt-`, `o1`, `o3`, `o4`, `chatgpt-`); when the configured model is empty, choose a per-provider preferred default (falling back to OpenRouter-preferred ids or `ids[0]`), and persist `knownModels`. API-key field (password with Show/Clear; Save key button) for all but ollama; **Sign in with OpenRouter** button when provider is `generic` and base URL origin is `https://openrouter.ai` (runs the §8 flow, then saves the credential and refreshes models). Auth status line (green authenticated / "no auth needed" for ollama / red error). **Test connection** = fetch models and report count. Ollama-specific failure help: if the error carries the browser-access sentinel, show the `OLLAMA_ORIGINS` PowerShell command with a Copy button, else suggest `ollama serve`.

**API tab**: provider dropdown (15 entries, with a "Free" checkbox filtering to free-API providers: gemini, groq, cerebras, cloudflare, huggingface, generic) + the provider form. Selecting a provider while the API tab is active immediately makes it the active provider.

**Search tab** (`SearchSettingsForm`): a permanent informational caption stating that the keyless catalogues (Wikipedia, Wikidata, World Bank, IMF, Eurostat, ECB, UN SDG, CKAN, data.gov.my, data.gov.sg, Open-Meteo) are **always available to the agent and need no setup** — this tab only adds keyed internet search. Then: status MessageBar from `getSearchSettingsStatusText` (native tier → success intent); provider Select with options None (= keyless only) / the keyed-or-self-hosted providers (tavily, google-cse, jina, searxng) — **no keyless bundle option**; per-provider caption (keyed: "Search uses your own provider key. It is off for each new session until you enable it in Chat."); signup link; API-key field when `requiresKey`; Engine ID (cx) field when `requiresEngineId`; Base URL override; "Allow reader fallback for fetched URLs" checkbox + explanation; **Save key** (also sets the provider active in webAccess) and **Test key/Test search** (runs a 1-result search for `'spreadsheet public data'` and reports). Clearing a key resets webAccess.provider to `'none'` and turns the chat toggle off (keyless search continues regardless).

**Appearance tab**: a Theme select with **System (default)**, **Light**, and **Dark**. Changes persist immediately through `setAppConfig` and affect the root Fluent theme without a reload.

### 18.7 Footer

Hidden when no session/totals. Left: `<n> tok` (session in+out, monospace). Right: context pressure `NN% ctx` shown only above 70% (red above 90%), then the model id (ellipsized).

### 18.8 AboutPanel / HarnessPanel / WorkbookScopeStrip

About: branding, provider list blurb, license (PolyForm Noncommercial 1.0.0), commercial contact, privacy-policy + GitHub links, `Version {__APP_VERSION__}`. HarnessPanel: dev-only canary tester (§9.5) with provider/baseUrl/model/key inputs, live event log, PASS/FAIL bar (not mounted in the shipping view set). WorkbookScopeStrip: standalone strip showing "<name> - host workbook" with a Refresh button (also not mounted in the current App shell).

---

## 19. Build, deployment & Office manifest

### 19.1 Vite (`vite.config.ts`)

Async config: `base: '/SheetClaw/'` for builds (GitHub Pages), `'/'` for dev. `define.__APP_VERSION__` from package.json. React plugin. Dev & preview servers on **port 3000 with HTTPS** using `getHttpsServerOptions()` from `office-addin-dev-certs`. Build inputs: `taskpane.html`, `oauth-start.html`, `oauth-callback.html` → `dist/`. Vitest config embedded: node environment, `src/**/*.test.ts`.

### 19.2 HTML shells

`taskpane.html`: CSP meta — `default-src 'self'; script-src 'self' https://appsforoffice.microsoft.com; style-src 'self' 'unsafe-inline'; connect-src *; img-src 'self' data: https:; font-src 'self' data:; object-src 'none'; base-uri 'self'`. Loads `https://appsforoffice.microsoft.com/lib/1/hosted/office.js`, styles html/body/#root to 100% height with hidden overflow, mounts `/src/taskpane/index.tsx`.

`oauth-start.html` / `oauth-callback.html`: see §8 (start validates the redirect origin; callback posts `xl-oauth-callback` payload to opener and/or `Office.context.ui.messageParent`, then self-closes after 600 ms).

### 19.3 Office manifest (`manifest.xml`)

TaskPaneApp, Id GUID, Host `Workbook`, `Permissions: ReadWriteDocument`, SourceLocation `https://cwtf.github.io/SheetClaw/taskpane.html` (production; sideloading points Excel at the manifest whose URLs resolve to the dev server during development), AppDomain `https://cwtf.github.io`, icon URLs under `/assets/`, VersionOverrides adding a Home-tab ribbon button "Open SheetClaw" that shows the task pane.

Deployment model: static hosting (GitHub Pages) — `npm run build` output published under `/SheetClaw/`.

---

## 20. Testing strategy

- **Unit tests** (Vitest, Node): adapters (SSE parsing incl. fragmented tool-call deltas, error mapping, Ollama lenient parser, native-search patches), agent (loop state machine with mocked executor/runner/store, context-builder compaction and selection injection, choice parsing), auth (secureStore roundtrip + fallbacks + tamper cases, PKCE/oauth validation), store (auth persistence ordering/migration, session history persistence, theme defaults and persistence), task pane (selection address normalization and injected Office runner), pricing (matching + cache-aware cost), web (net URL validation + CORS classification with fake fetch, fetch formatting, search arg validation, keyless bundle interleave/fallback), workbook (executor validation/error mapping, snapshot undo, diff math, chart/pivot handlers against a mock Excel context, tool-registration completeness).
- **Integration** (`src/web/__integration__/provider-cors.integration.ts`, `npm run test:providers`): hits every keyless search endpoint through a CORS-enforcing fetch wrapper to catch providers that drop `Access-Control-Allow-Origin` — Node fetch can't catch this in unit tests. Run when adding/changing a keyless provider.
- Excel-dependent handlers use an injected mock `Excel.RequestContext`; error-class mapping intentionally checks `e.name` as well as `instanceof` to survive Vitest module duplication.

---

## 21. Suggested implementation order

1. **Types + storage**: `src/types/*`, storage envelope, ULID helper.
2. **Secure store + credentials** (§7) with tests.
3. **Store slices** (config with defaults/migrations, auth with sealed persistence queue, session with transcript persistence, usage buckets).
4. **Adapters**: OpenAIAdapter first (it serves 14 providers), then Anthropic, then Ollama wrapper + lenient parser; `createAdapter`; harness canary.
5. **Workbook layer**: registry → executor → a1notation → snapshot; then the tool modules in the §15 order; `createWorkbookLayer` registration table.
6. **Agent**: system prompt → context builder/compaction → choice tool → AgentLoop → singleton.
7. **Web layer**: net guards → fetch_url → provider adapters (start with wikipedia/tavily as templates) → keyless bundle → web_search → native-search module → taskpane wiring.
8. **Pricing + usage** queries/export.
9. **UI**: bootstrap, App shell, ChatPanel (largest), Settings, History, Usage, Footer, About.
10. **OAuth** flow + HTML shells; manifest + vite config; sideload and verify in Excel.

---

## 22. Key invariants & gotchas

1. **Never mutate the workbook without a snapshot-first flow.** The loop, not the handlers, owns snapshot + diff + confirmation; handlers must be pure "do the write" functions.
2. **Session updates must go through `updateSessionById`** — the loop may outlive the currently displayed session (user can browse History mid-run), and the ById path patches persisted transcripts.
3. **The executor keeps all tools registered even when filtered out of a run**; allowed-set enforcement in the loop is what produces the corrective "not available in this session" errors that stop models from spinning on web tools.
4. **The keyless search bundle is baked into every run** — `web_search` and `fetch_url` are always in the agent's tool set, with the keyless bundle as the default `web_search` backend, no Settings or toggle required. The Search toggle and Settings → Search provider strictly govern internet search that requires an API key (keyed BYOK) or is billed to the LLM provider (native); keyed search resets to off each new session.
5. **CORS is the defining network constraint**: guarded fetch classifies TypeErrors via a no-cors probe, caches blocked hosts per session, and both the system prompt and tool error strings explicitly instruct the model not to retry blocked hosts.
6. **Anthropic tool results must be merged into a single user message**; the browser call requires the `anthropic-dangerous-direct-browser-access` header.
7. **OpenAI streams may repeat `finish_reason`** (OpenRouter) — emit `done` once. Usage may arrive on a chunk without choices.
8. **Tool-call arguments stream as fragments** keyed by `index`; a delta with `id`+`name` starts a new accumulator, everything else appends. Unparseable final JSON degrades to `{}` while keeping `rawArguments`.
9. **Compaction always squashes old tool results** (even under budget) but preserves the two most recent, and always preserves the first user message; the workbook manifest rides inside the first user message, not the system prompt.
10. **`write_range` writes through `range.formulas`** unless `as_text`, so `=SUM(...)` strings become live formulas; dimension mismatches are hard validation errors, not silent clipping.
11. **Undo fidelity**: range snapshots restore formulas+numberFormat fully; chart/pivot creations get coarse structural snapshots whose undo deletes the created object; other mutations (e.g. sheet ops) currently have no undo.
12. **Storage is quota-bounded**: usage day-buckets are the designated eviction victims; usage history rolls at 30 days.
13. **Auth persistence is a serialized queue** — encryption is async and writes must land in call order.
14. **All UI-only message roles** (`tool_call`, `confirmation`, `system_notice`) must be excluded from LLM serialization or providers will reject the payload.
15. **Kimi native search**: the `$web_search` tool call is answered by echoing its raw arguments back; its tool result is sent with `name: '$web_search'` and passes through context-building untruncated as a raw string.
16. **Selection references are message-scoped**: capture the active sheet/address again at submit time, persist it on that `UserMessage`, and serialize that stored value as `current_selection`. Never resolve an older phrase such as “this cell” against a newer live selection.
