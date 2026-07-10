import type {
  LLMRequest,
  NormalizedMessage,
  NormalizedToolCall,
  ToolSpec,
  Message,
  AgentSession,
  ProviderConfig,
} from '../types';
import type { WorkbookRegistry } from '../workbook/registry';
import { buildSystemPrompt } from './system-prompt';
import { getNativeSearchCapability } from '../adapters/native-search';

// ── Token estimation ───────────────────────────────────────────────────────
// Rough character-to-token ratio; good enough for budget decisions.
const CHARS_PER_TOKEN = 4;
// Sized above fetch_url's 20k-char text cap (plus JSON-encoding overhead) so a
// full-mode fetch reaches the model intact; compact() degrades older results
// when the context budget is under pressure.
const MAX_TOOL_RESULT_CHARS = 24_000;

export function estimateTokens(text: string): number {
  return Math.ceil(text.length / CHARS_PER_TOKEN);
}

// ── Message conversion ─────────────────────────────────────────────────────

function toNormalized(messages: Message[]): NormalizedMessage[] {
  const out: NormalizedMessage[] = [];
  const toolNames = new Map<string, string>();
  for (const m of messages) {
    switch (m.role) {
      case 'user':
        out.push({ role: 'user', content: m.text });
        break;
      case 'assistant': {
        const toolCalls: NormalizedToolCall[] = (m.toolCalls ?? []).map(tc => ({
          id: tc.id,
          name: tc.name,
          arguments: tc.arguments,
        }));
        for (const tc of toolCalls) toolNames.set(tc.id, tc.name);
        out.push({ role: 'assistant', content: m.text, toolCalls: toolCalls.length ? toolCalls : undefined });
        break;
      }
      case 'tool_call':
        toolNames.set(m.toolCall.id, m.toolCall.name);
        break;
      case 'tool': {
        const toolName = toolNames.get(m.toolCallId) ?? m.result.toolCallId;
        const payload = m.result.ok
          ? toolName === '$web_search' && typeof m.result.data === 'string'
            ? m.result.data
            : JSON.stringify(m.result.data)
          : JSON.stringify(m.result.error);
        // Truncate large payloads to protect context budget
        const content = payload ?? '';
        const truncated = content.length > MAX_TOOL_RESULT_CHARS
          ? content.slice(0, MAX_TOOL_RESULT_CHARS) + `... (${content.length - MAX_TOOL_RESULT_CHARS} chars truncated)`
          : content;
        out.push({ role: 'tool', toolCallId: m.toolCallId, name: toolName, content: truncated });
        break;
      }
      // ToolCallMessage, ConfirmationMessage, SystemNoticeMessage are UI-only — skip
    }
  }
  return out;
}

// ── Compaction ─────────────────────────────────────────────────────────────

function compact(
  history: NormalizedMessage[],
  fixedTokens: number,  // system + manifest
  budget: number,
  maxOutput: number
): NormalizedMessage[] {
  const available = budget - fixedTokens - maxOutput;
  if (available <= 0) return history.slice(-2); // emergency floor

  // Identify the two most-recent tool results — they hold the data the model
  // is actively working with; squashing them forces a refetch and more iterations.
  const toolIndices = history.map((m, i) => (m.role === 'tool' ? i : -1)).filter(i => i >= 0);
  const recent = new Set(toolIndices.slice(-2));
  const squash = (msgs: NormalizedMessage[], oldLimit: number, recentLimit: number) =>
    msgs.map((m, i) => {
      if (m.role !== 'tool') return m;
      const limit = recent.has(i) ? recentLimit : oldLimit;
      return m.content.length > limit
        ? { ...m, content: m.content.slice(0, limit) + '…[truncated]' }
        : m;
    });

  // Age-based pass (unconditional): collapse tool results older than the two
  // most recent ones even when the full transcript fits the context budget.
  // Prevents token count from growing linearly across long sessions on
  // large-context models where fits() would otherwise always return true.
  let squashed = squash(history, 2000, Number.MAX_SAFE_INTEGER);

  const fits = (msgs: NormalizedMessage[]) => estimateTokens(JSON.stringify(msgs)) <= available;
  if (fits(squashed)) return squashed;

  // Budget-pressure passes: further degrade when the age-based pass isn't enough.
  squashed = squash(squashed, 200, 2000);
  if (fits(squashed)) return squashed;

  // Step 3: drop oldest pairs, always keep first user message + last 4 messages
  const firstUser = squashed.findIndex(m => m.role === 'user');
  const kept = firstUser >= 0 ? [squashed[firstUser]] : [];
  let tail = squashed.slice(firstUser + 1);

  while (tail.length > 4 && estimateTokens(JSON.stringify([...kept, ...tail])) > available) {
    // Drop the oldest user+assistant pair (skip tool results attached to it)
    const firstUserInTail = tail.findIndex(m => m.role === 'user');
    if (firstUserInTail < 0) { tail = tail.slice(2); break; }
    // Drop from firstUserInTail through the next user message
    const nextUser = tail.findIndex((m, i) => i > firstUserInTail && m.role === 'user');
    tail = nextUser < 0 ? tail.slice(-4) : tail.slice(nextUser);
  }
  return [...kept, ...tail];
}

// ── ContextBuilder ─────────────────────────────────────────────────────────

export class ContextBuilder {
  constructor(
    private registry: WorkbookRegistry,
    // Injected so this module stays store-free (and testable in Node).
    private getReaderFallback: () => boolean = () => false
  ) {}

  build(
    session: AgentSession,
    messages: Message[],
    tools: ToolSpec[],
    cfg: ProviderConfig
  ): LLMRequest {
    const system = buildSystemPrompt(session.scope.workbookId, {
      keyedSearchEnabled: session.webSearchEnabled,
      readerFallback: this.getReaderFallback(),
    });
    const manifest = this.registry.getManifest();
    const manifestStr = `\n\n<workbook_manifest>\n${JSON.stringify(manifest, null, 2)}\n</workbook_manifest>`;

    const budget = cfg.contextLimits.maxContextTokens;
    const maxOutput = cfg.maxOutputTokens ?? 4096;
    const fixedTokens = estimateTokens(system + manifestStr + JSON.stringify(tools));

    const rawHistory = toNormalized(messages);
    const history = compact(rawHistory, fixedTokens, budget, maxOutput);

    // Prepend manifest as a system-level note in the first user message slot
    const firstUser = history.findIndex(m => m.role === 'user');
    const withManifest: NormalizedMessage[] =
      firstUser >= 0
        ? [
            ...history.slice(0, firstUser),
            { role: 'user', content: history[firstUser].content + manifestStr },
            ...history.slice(firstUser + 1),
          ]
        : history;

    return {
      model: cfg.model,
      messages: withManifest,
      tools,
      nativeSearch: session.webSearchEnabled ? getNativeSearchCapability(cfg.provider, cfg.model) : undefined,
      system,
      temperature: cfg.temperature,
      maxTokens: maxOutput,
    };
  }

  estimateInputTokens(
    session: AgentSession,
    messages: Message[],
    tools: ToolSpec[],
    cfg: ProviderConfig
  ): number {
    const req = this.build(session, messages, tools, cfg);
    return estimateTokens(
      (req.system ?? '') +
      JSON.stringify(req.messages) +
      JSON.stringify(req.tools)
    );
  }
}
