import type {
  LLMClient,
  LLMRequest,
  LLMStreamEvent,
  LLMError,
  NormalizedMessage,
  ToolSpec,
  ModelInfo,
  ProviderCapabilities,
  ProviderKey,
} from '../types';
import { getOpenAINativeSearchPatch } from './native-search';

// ── Provider config injected at construction ───────────────────────────────

export interface OpenAIAdapterConfig {
  baseUrl: string;
  apiKey: string;
  provider?: ProviderKey;
  extraHeaders?: Record<string, string>;
}

// ── Wire types (OpenAI SSE) ────────────────────────────────────────────────

interface OAIToolCallDelta {
  index: number;
  id?: string;
  type?: 'function';
  function?: { name?: string; arguments?: string };
}

interface OAIChunk {
  choices?: Array<{
    index: number;
    delta?: {
      role?: string;
      content?: string | null;
      tool_calls?: OAIToolCallDelta[];
      annotations?: unknown[];
    };
    finish_reason?: string | null;
    message?: { annotations?: unknown[] };
  }>;
  usage?: {
    prompt_tokens: number;
    completion_tokens: number;
    prompt_tokens_details?: { cached_tokens?: number };
    [key: string]: unknown;
  };
}

// ── Serialization helpers ──────────────────────────────────────────────────

function serializeTools(tools: ToolSpec[]) {
  return tools.map(t => ({
    type: 'function' as const,
    function: {
      name: t.name,
      description: t.description,
      parameters: { type: t.parameters.type, properties: t.parameters.properties, required: t.parameters.required },
    },
  }));
}

function serializeToolsWithNative(tools: ToolSpec[], nativeTools: unknown[] = []) {
  return [...serializeTools(tools), ...nativeTools];
}

function serializeMessages(messages: NormalizedMessage[]): unknown[] {
  return messages.map(m => {
    if (m.role === 'tool') {
      return {
        role: 'tool',
        tool_call_id: m.toolCallId,
        content: m.content,
        ...(m.name === '$web_search' ? { name: m.name } : {}),
      };
    }
    if (m.role === 'assistant' && m.toolCalls?.length) {
      return {
        role: 'assistant',
        content: m.content || null,
        tool_calls: m.toolCalls.map(tc => ({
          id: tc.id,
          type: 'function',
          function: { name: tc.name, arguments: JSON.stringify(tc.arguments) },
        })),
      };
    }
    return { role: m.role, content: m.content };
  });
}

// ── Error mapping ──────────────────────────────────────────────────────────

async function mapHttpError(res: Response): Promise<LLMError> {
  let body: unknown;
  try { body = await res.json(); } catch { body = await res.text(); }
  if (res.status === 401) return { code: 'AuthError', message: 'Invalid API key (401)' };
  if (res.status === 429) {
    const retryAfter = Number(res.headers.get('retry-after') ?? NaN);
    return { code: 'RateLimitError', message: 'Rate limited (429)', retryAfter: isNaN(retryAfter) ? undefined : retryAfter };
  }
  return { code: 'ProviderError', message: `HTTP ${res.status}`, status: res.status, body };
}

// ── SSE parser ─────────────────────────────────────────────────────────────

async function* parseSSE(body: ReadableStream<Uint8Array>): AsyncIterable<string> {
  const reader = body.getReader();
  const decoder = new TextDecoder();
  let buf = '';
  try {
    while (true) {
      const { done, value } = await reader.read();
      if (done) break;
      buf += decoder.decode(value, { stream: true });
      const lines = buf.split('\n');
      buf = lines.pop() ?? '';
      for (const line of lines) {
        if (line.startsWith('data: ')) yield line.slice(6);
      }
    }
    // flush remaining
    if (buf.startsWith('data: ')) yield buf.slice(6);
  } finally {
    reader.releaseLock();
  }
}

// ── Adapter ────────────────────────────────────────────────────────────────

export class OpenAIAdapter implements LLMClient {
  constructor(private cfg: OpenAIAdapterConfig) {}

  capabilities(): ProviderCapabilities {
    return {
      supportsTools: true,
      supportsStreaming: true,
      supportsOAuth: false,
      nativeUsage: true,
      toolFormat: 'openai',
    };
  }

  async listModels(): Promise<ModelInfo[]> {
    const res = await fetch(`${this.cfg.baseUrl}/models`, {
      headers: this.#authHeaders(),
    });
    if (!res.ok) throw new Error(`listModels failed: ${res.status}`);
    const data = (await res.json()) as { data: Array<{ id: string }> };
    return data.data.map(m => ({ id: m.id }));
  }

  async *chat(req: LLMRequest, signal: AbortSignal): AsyncIterable<LLMStreamEvent> {
    const body: Record<string, unknown> = {
      model: req.model,
      messages: [
        ...(req.system ? [{ role: 'system', content: req.system }] : []),
        ...serializeMessages(req.messages),
      ],
      stream: true,
      stream_options: { include_usage: true },
    };
    const nativePatch = getOpenAINativeSearchPatch(this.cfg.provider, req.nativeSearch);
    if (req.tools.length || nativePatch.tools?.length) body.tools = serializeToolsWithNative(req.tools, nativePatch.tools);
    if (nativePatch.body) Object.assign(body, nativePatch.body);
    if (req.temperature !== undefined) body.temperature = req.temperature;
    if (req.maxTokens !== undefined) body.max_tokens = req.maxTokens;

    let res: Response;
    try {
      res = await fetch(`${this.cfg.baseUrl}/chat/completions`, {
        method: 'POST',
        headers: { 'Content-Type': 'application/json', ...this.#authHeaders() },
        body: JSON.stringify(body),
        signal,
      });
    } catch (e) {
      yield { type: 'error', error: { code: 'NetworkError', message: String(e) } };
      return;
    }

    if (!res.ok) {
      yield { type: 'error', error: await mapHttpError(res) };
      return;
    }

    if (!res.body) {
      yield { type: 'error', error: { code: 'MalformedResponseError', message: 'Empty response body' } };
      return;
    }

    // Index-keyed accumulator for reassembling fragmented tool-call arguments
    const toolCallAccum: Record<number, { id: string; name: string; argsBuf: string }> = {};
    // Guard against providers (e.g. OpenRouter) that send finish_reason on multiple chunks
    let doneSent = false;

    try {
      for await (const raw of parseSSE(res.body)) {
        if (raw === '[DONE]') break;
        if (!raw.trim()) continue;

        let chunk: OAIChunk;
        try {
          chunk = JSON.parse(raw) as OAIChunk;
        } catch {
          yield { type: 'error', error: { code: 'MalformedResponseError', message: `JSON parse error: ${raw.slice(0, 80)}` } };
          return;
        }

        const choice = chunk.choices?.[0];
        if (choice) {
          const delta = choice.delta ?? {};

          // Text delta
          if (delta.content) {
            yield { type: 'text-delta', delta: delta.content };
          }

          // Tool call deltas
          for (const tc of delta.tool_calls ?? []) {
            const idx = tc.index;

            if (tc.id && tc.function?.name) {
              // First delta for this index — emit start + begin accumulating
              toolCallAccum[idx] = { id: tc.id, name: tc.function.name, argsBuf: tc.function.arguments ?? '' };
              yield { type: 'tool-call-start', index: idx, id: tc.id, name: tc.function.name };
              if (tc.function.arguments) {
                yield { type: 'tool-call-delta', index: idx, argumentsDelta: tc.function.arguments };
              }
            } else if (tc.function?.arguments) {
              // Continuation delta
              toolCallAccum[idx].argsBuf += tc.function.arguments;
              yield { type: 'tool-call-delta', index: idx, argumentsDelta: tc.function.arguments };
            }
          }

          // Finish — guard against providers that send finish_reason on multiple chunks
          if (choice.finish_reason && !doneSent) {
            doneSent = true;
            for (const idx of Object.keys(toolCallAccum).map(Number)) {
              yield { type: 'tool-call-end', index: idx };
            }
            const reason = choice.finish_reason as 'stop' | 'tool_calls' | 'length';
            yield { type: 'done', finishReason: reason };
          }
        }

        // Usage (may arrive on a chunk without choices, or the last chunk)
        if (chunk.usage) {
          const u = chunk.usage;
          yield {
            type: 'usage',
            inputTokens: u.prompt_tokens,
            outputTokens: u.completion_tokens,
            cacheRead: u.prompt_tokens_details?.cached_tokens,
            source: 'provider',
          };
        }
      }
    } catch (e) {
      if (signal.aborted) return;
      yield { type: 'error', error: { code: 'NetworkError', message: String(e) } };
    }
  }

  #authHeaders(): Record<string, string> {
    // A local gateway with auth disabled can reject an empty bearer outright,
    // so send no Authorization at all rather than `Bearer `. Remote providers
    // would have failed on an empty key either way.
    return {
      ...(this.cfg.apiKey ? { Authorization: `Bearer ${this.cfg.apiKey}` } : {}),
      ...this.cfg.extraHeaders,
    };
  }
}
