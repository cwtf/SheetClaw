# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Commands

```bash
npm run dev              # Start Vite dev server (HTTPS on localhost:3000)
npm run build            # tsc + vite build (outputs to dist/)
npm test                 # Run Vitest test suite (single pass)
npm run test:watch       # Vitest in watch mode
npm run test:providers   # Live-network CORS checks for keyless search providers (needs internet)
npm run sideload         # Register add-in with Excel via office-addin-debugging
npm run validate-manifest  # Validate manifest.xml
npm run install-certs    # Install dev HTTPS certs (run once on a new machine)
```

Run a single test file:
```bash
npx vitest run src/agent/__tests__/loop.test.ts
```

The dev server must be running before sideloading. Excel loads the task pane from `https://localhost:3000/taskpane.html`.

## Architecture

SheetClaw is an Office.js Excel task pane add-in: a React SPA embedded inside Excel's sidebar. It runs in a WebView within the Office runtime — no Node.js at runtime, no server, no network proxy of its own.

### Entry point

[src/taskpane/index.tsx](src/taskpane/index.tsx) bootstraps inside `Office.onReady`. The root component is [src/taskpane/App.tsx](src/taskpane/App.tsx), which renders a five-tab Fluent UI layout: Chat, History, Usage, Settings, About.

### Layer diagram

```
App.tsx (tab shell)
  └─ ChatPanel / HistoryPanel / UsageDashboard / SettingsPanel / AboutPanel

taskpane/workbookLayer.ts       ← singleton factory; wires the three core objects
  ├─ WorkbookRegistry           ← tracks the host workbook (always exactly one on Windows desktop)
  ├─ ToolExecutor               ← registry of ToolSpec → ToolHandler, validates args, runs Excel.run()
  └─ SnapshotManager            ← captures before-state for undo; range snapshots + structural snapshots

agent/loop.ts (AgentLoop)       ← main agentic loop (max 25 iterations)
  ├─ ContextBuilder             ← builds the LLM request (system prompt + trimmed message history)
  ├─ LLMClient (via adapter)    ← streams SSE events: text-delta / tool-call-start / tool-call-delta / usage / done
  └─ ToolExecutor               ← executes tool calls; mutating tools pause for user confirmation

store/index.ts (Zustand)        ← single global store, four slices:
  ├─ config    ← provider settings, active provider, web-access config, UI flags
  ├─ auth      ← per-provider auth state (encrypted API keys, OAuth tokens)
  ├─ session   ← current and past AgentSessions, messages, pending confirmations
  └─ usage     ← per-turn usage records, rolling history
```

### Provider adapters

[src/adapters/](src/adapters/) contains three concrete adapters: `OpenAIAdapter`, `AnthropicAdapter`, `OllamaAdapter`. `createAdapter()` routes all OpenAI-compatible providers (openai, deepseek, groq, mistral, together, kimi, glm, qwen, llama, generic) through `OpenAIAdapter`. All adapters implement `LLMClient` (`src/types/llm.ts`) and yield a typed async iterator of streaming events.

### Tool system

Each tool is a `ToolSpec` (name, description, JSON Schema parameters, `mutating: boolean`, `runtime: 'excel' | 'none'`) paired with a `ToolHandler`. All workbook tools live in [src/workbook/tools/](src/workbook/tools/). Web tools (`fetch_url`, `web_search`) are registered separately in `workbookLayer.ts`.

Mutating tools (writes, clears, chart/pivot creation, etc.) pause the agent loop at `awaiting_confirmation`; the UI shows a diff and the user must click Apply or Cancel before execution continues.

Non-mutating tools with `runtime: 'none'` (web search, fetch) run concurrently when consecutive calls appear in the same turn.

### Credentials and storage

API keys are encrypted at rest with AES-GCM-256 via Web Crypto ([src/auth/secureStore.ts](src/auth/secureStore.ts)). The encryption key is a non-extractable `CryptoKey` in IndexedDB; ciphertext is in `localStorage`. Zustand slices are persisted to `localStorage` via a versioned envelope ([src/store/storage.ts](src/store/storage.ts)).

### Pricing

[src/pricing/index.ts](src/pricing/index.ts) holds bundled per-model pricing data. Usage records are appended to the store each turn and summarised in the Usage tab.

### Web access

Optional. [src/web/](src/web/) contains `fetch.ts` (fetch_url tool with Jina reader proxy fallback) and `search.ts` (web_search routing to Tavily / SearXNG / Google CSE / Jina plus keyless sources such as Wikipedia, Wikidata, World Bank, CKAN, data.gov.my, data.gov.sg, IMF, Eurostat, ECB, Open-Meteo, and UN SDG). Provider selection and BYOK keys are configured in Settings → Web Access.

## Key constraints

- The app runs in a browser WebView inside Excel — no Node APIs, no file system, no arbitrary outbound connections without CORS headers.
- `WorkbookRegistry` always holds exactly one workbook on Windows desktop Office (host-only model). Multi-workbook enumeration requires a sidecar that is not implemented.
- All Office.js calls must be batched inside `Excel.run()` / `ctx.sync()`. The `LoopRunner` type alias (`(fn) => Excel.run(fn)`) is injected for testability — tests substitute a mock runner.
- Tests use Vitest and run in a Node environment; Office.js globals (`Excel`, `Office`) are not available in tests and must be mocked or kept out of tested modules.
- Node fetch does not enforce CORS, so unit tests cannot catch a provider endpoint that stops sending `Access-Control-Allow-Origin` (which breaks every fetch from the task pane). `npm run test:providers` probes the live keyless search endpoints through a CORS-enforcing fetch wrapper — run it when adding or changing a keyless provider endpoint.
