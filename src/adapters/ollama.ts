import type {
  LLMClient,
  LLMRequest,
  LLMStreamEvent,
  ModelInfo,
  ProviderCapabilities,
  ToolSpec,
} from '../types';
import { OpenAIAdapter } from './openai';

export interface OllamaAdapterConfig {
  baseUrl?: string;
}

const DEFAULT_BASE = 'http://localhost:11434';
const BROWSER_ACCESS_ERROR_PREFIX = 'Ollama is reachable, but this add-in cannot read it.';

function browserOrigin(): string | undefined {
  if (typeof window === 'undefined') return undefined;
  return window.location?.origin;
}

export function getOllamaBrowserAccessCommand(origin = browserOrigin()): string {
  const allowedOrigin = origin ?? '<add-in-origin>';
  return `$env:OLLAMA_ORIGINS='${allowedOrigin}'; ollama serve`;
}

export function isOllamaBrowserAccessError(message: string): boolean {
  return message.includes(BROWSER_ACCESS_ERROR_PREFIX);
}

async function canReachOllamaWithoutCors(baseUrl: string): Promise<boolean> {
  try {
    await fetch(baseUrl, {
      method: 'GET',
      mode: 'no-cors',
      cache: 'no-store',
    });
    return true;
  } catch {
    return false;
  }
}

async function mapOllamaFetchFailure(baseUrl: string, cause: unknown): Promise<Error> {
  const reachable = await canReachOllamaWithoutCors(baseUrl);
  if (reachable) {
    const origin = browserOrigin();
    const originText = origin ? ` (${origin})` : '';
    return new Error(
      `${BROWSER_ACCESS_ERROR_PREFIX} Allow the add-in origin${originText} by quitting Ollama and restarting it with: ${getOllamaBrowserAccessCommand(origin)}`
    );
  }
  return cause instanceof Error ? cause : new Error(String(cause));
}

// ── Lenient fallback parser (§6.3) ─────────────────────────────────────────
// Some Ollama models emit tool calls as plain text/JSON rather than the OpenAI
// tool_calls field. We scan assistant content for a fenced JSON block that
// matches a known tool name and promote it to a synthetic tool-call-start/end.

const FENCE_RE = /```(?:json)?\s*(\{[\s\S]*?\})\s*```/g;
const JSON_RE = /(\{[\s\S]*\})/;

export function parseLenientToolCall(
  text: string,
  knownTools: ToolSpec[]
): { name: string; arguments: Record<string, unknown> } | null {
  const toolNames = new Set(knownTools.map(t => t.name));
  const candidates: string[] = [];

  // Collect all fenced JSON blocks
  let m: RegExpExecArray | null;
  FENCE_RE.lastIndex = 0;
  while ((m = FENCE_RE.exec(text)) !== null) candidates.push(m[1]);

  // Also try a bare top-level JSON object in the text
  const bareMatch = JSON_RE.exec(text);
  if (bareMatch) candidates.push(bareMatch[1]);

  for (const candidate of candidates) {
    try {
      const parsed = JSON.parse(candidate) as Record<string, unknown>;
      const name = typeof parsed.name === 'string' ? parsed.name :
                   typeof parsed.tool === 'string' ? parsed.tool : null;
      if (name && toolNames.has(name)) {
        const args = (parsed.arguments ?? parsed.parameters ?? parsed.input ?? parsed) as Record<string, unknown>;
        const cleaned: Record<string, unknown> = { ...args };
        delete cleaned.name;
        delete cleaned.tool;
        return { name, arguments: cleaned };
      }
    } catch { /* skip */ }
  }
  return null;
}

// ── Adapter ────────────────────────────────────────────────────────────────

export class OllamaAdapter implements LLMClient {
  private base: string;
  // Delegate to OpenAI adapter for the OpenAI-compatible endpoint
  private inner: OpenAIAdapter;

  constructor(cfg: OllamaAdapterConfig = {}) {
    this.base = (cfg.baseUrl ?? DEFAULT_BASE).replace(/\/+$/, '');
    this.inner = new OpenAIAdapter({ baseUrl: `${this.base}/v1`, apiKey: 'ollama' });
  }

  capabilities(): ProviderCapabilities {
    return {
      supportsTools: true, // assumed; degraded per model at runtime
      supportsStreaming: true,
      supportsOAuth: false,
      nativeUsage: true,
      toolFormat: 'openai',
    };
  }

  async listModels(): Promise<ModelInfo[]> {
    let res: Response;
    try {
      res = await fetch(`${this.base}/api/tags`);
    } catch (e) {
      throw await mapOllamaFetchFailure(this.base, e);
    }
    if (!res.ok) throw new Error(`listModels failed: ${res.status}`);
    const data = (await res.json()) as { models: Array<{ name: string; model: string }> };
    return data.models.map(m => ({ id: m.model ?? m.name, name: m.name }));
  }

  async *chat(req: LLMRequest, signal: AbortSignal): AsyncIterable<LLMStreamEvent> {
    const events: LLMStreamEvent[] = [];
    let assistantText = '';
    let hadToolCall = false;

    // First pass: collect all events from the inner OpenAI-compat adapter
    for await (const ev of this.inner.chat(req, signal)) {
      events.push(ev);
      if (ev.type === 'text-delta') assistantText += ev.delta;
      if (ev.type === 'tool-call-start') hadToolCall = true;
    }

    // If no tool calls were emitted but we got text, try lenient parser
    if (!hadToolCall && assistantText && req.tools.length) {
      const found = parseLenientToolCall(assistantText, req.tools);
      if (found) {
        const syntheticId = `lenient_${Date.now()}`;
        const argsStr = JSON.stringify(found.arguments);
        yield { type: 'tool-call-start', index: 0, id: syntheticId, name: found.name };
        yield { type: 'tool-call-delta', index: 0, argumentsDelta: argsStr };
        yield { type: 'tool-call-end', index: 0 };
        // Re-emit usage + done with tool_calls reason if present, else synthesize
        for (const ev of events) {
          if (ev.type === 'usage') yield ev;
        }
        yield { type: 'done', finishReason: 'tool_calls' };
        return;
      }
    }

    // Normal path: emit what the inner adapter produced
    for (const ev of events) yield ev;
  }
}
