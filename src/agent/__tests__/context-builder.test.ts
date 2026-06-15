import { describe, it, expect } from 'vitest';
import { ContextBuilder, estimateTokens } from '../context-builder';
import { WorkbookRegistry } from '../../workbook/registry';
import type { AgentSession, Message, ProviderConfig } from '../../types';

// ── Fixtures ───────────────────────────────────────────────────────────────

function makeRegistry(): WorkbookRegistry {
  const r = new WorkbookRegistry();
  (r as unknown as { handles: Map<string, unknown>; activeId: string; hostId: string }).handles.set('wb1', {
    workbookId: 'wb1', name: 'Test.xlsx', isActive: true, isHost: true,
    sheets: [{ name: 'Sheet1', position: 0, visible: true }],
    lastRefreshed: new Date().toISOString(), capability: 'host-only',
  });
  (r as unknown as { activeId: string; hostId: string }).activeId = 'wb1';
  (r as unknown as { hostId: string }).hostId = 'wb1';
  return r;
}

const SESSION: AgentSession = {
  id: 's1', createdAt: '', scope: { workbookId: 'wb1' },
  status: 'building', iteration: 1, maxIterations: 25,
  provider: 'ollama', model: 'llama3.2',
  messageIds: [], tokenBudget: { used: 0, window: 128000 },
  webSearchEnabled: false,
  totals: { inputTokens: 0, outputTokens: 0, costUsd: 0 },
};

const CFG: ProviderConfig = {
  provider: 'ollama', enabled: true,
  baseUrl: 'http://localhost:11434', model: 'llama3.2',
  authMode: 'none', authStateRef: '',
  contextLimits: { maxContextTokens: 128000, historyTokenCap: 100000, maxInlineSheetCells: 5000 },
  maxOutputTokens: 4096,
};

function makeMessages(...pairs: Array<[string, string]>): Message[] {
  const msgs: Message[] = [];
  let i = 0;
  for (const [user, assistant] of pairs) {
    msgs.push({ id: `u${i}`, sessionId: 's1', createdAt: '', role: 'user', text: user });
    msgs.push({ id: `a${i}`, sessionId: 's1', createdAt: '', role: 'assistant', text: assistant });
    i++;
  }
  return msgs;
}

// ── Tests ──────────────────────────────────────────────────────────────────

describe('estimateTokens', () => {
  it('returns ceil(length/4)', () => {
    expect(estimateTokens('abcd')).toBe(1);
    expect(estimateTokens('abcde')).toBe(2);
    expect(estimateTokens('')).toBe(0);
  });
});

describe('ContextBuilder.build', () => {
  it('returns an LLMRequest with the user instruction embedded', () => {
    const cb = new ContextBuilder(makeRegistry());
    const msgs = makeMessages(['Sum column A', 'I will do that.']);
    const req = cb.build(SESSION, msgs, [], CFG);

    expect(req.model).toBe('llama3.2');
    expect(req.system).toContain('SheetClaw');
    expect(req.messages.some(m => m.role === 'user' && (m.content as string).includes('Sum column A'))).toBe(true);
  });

  it('includes workbook manifest in first user message', () => {
    const cb = new ContextBuilder(makeRegistry());
    const msgs = makeMessages(['Hello', 'Hi']);
    const req = cb.build(SESSION, msgs, [], CFG);
    const firstUser = req.messages.find(m => m.role === 'user');
    expect(firstUser?.content).toContain('Test.xlsx');
  });

  it('normalizes tool result messages', () => {
    const cb = new ContextBuilder(makeRegistry());
    const msgs: Message[] = [
      { id: 'u1', sessionId: 's1', createdAt: '', role: 'user', text: 'Read A1' },
      { id: 'a1', sessionId: 's1', createdAt: '', role: 'assistant', text: '', toolCalls: [{ id: 'tc1', name: 'read_range', arguments: {}, workbookId: 'wb1', mutating: false }] },
      { id: 'tr1', sessionId: 's1', createdAt: '', role: 'tool', toolCallId: 'tc1', result: { toolCallId: 'tc1', ok: true, data: { values: [['hello']] } } },
    ];
    const req = cb.build(SESSION, msgs, [], CFG);
    const toolMsg = req.messages.find(m => m.role === 'tool');
    expect(toolMsg).toBeDefined();
    expect(toolMsg?.content).toContain('hello');
  });

  it('keeps tool results up to the fetch_url cap intact', () => {
    const cb = new ContextBuilder(makeRegistry());
    const fetchSized = 'x'.repeat(20_000);
    const msgs: Message[] = [
      { id: 'u1', sessionId: 's1', createdAt: '', role: 'user', text: 'Read' },
      { id: 'a1', sessionId: 's1', createdAt: '', role: 'assistant', text: '' },
      { id: 'tr1', sessionId: 's1', createdAt: '', role: 'tool', toolCallId: 'tc1', result: { toolCallId: 'tc1', ok: true, data: fetchSized } },
    ];
    const req = cb.build(SESSION, msgs, [], CFG);
    const toolMsg = req.messages.find(m => m.role === 'tool');
    expect(toolMsg?.content).not.toContain('truncated');
    expect((toolMsg?.content as string).length).toBeGreaterThanOrEqual(20_000);
  });

  it('truncates oversized tool result payloads', () => {
    const cb = new ContextBuilder(makeRegistry());
    const bigData = 'x'.repeat(30_000);
    const msgs: Message[] = [
      { id: 'u1', sessionId: 's1', createdAt: '', role: 'user', text: 'Read' },
      { id: 'a1', sessionId: 's1', createdAt: '', role: 'assistant', text: '' },
      { id: 'tr1', sessionId: 's1', createdAt: '', role: 'tool', toolCallId: 'tc1', result: { toolCallId: 'tc1', ok: true, data: bigData } },
    ];
    const req = cb.build(SESSION, msgs, [], CFG);
    const toolMsg = req.messages.find(m => m.role === 'tool');
    expect((toolMsg?.content as string).length).toBeLessThan(24_100);
    expect(toolMsg?.content).toContain('truncated');
  });

  it('compaction squashes older tool results before the most recent ones', () => {
    const cb = new ContextBuilder(makeRegistry());
    const tightCfg: ProviderConfig = {
      ...CFG,
      maxOutputTokens: 16,
      contextLimits: { ...CFG.contextLimits, maxContextTokens: 15_600 },
    };
    const big = (ch: string) => ch.repeat(20_000);
    const msgs: Message[] = [
      { id: 'u1', sessionId: 's1', createdAt: '', role: 'user', text: 'Fetch four pages' },
      { id: 'a1', sessionId: 's1', createdAt: '', role: 'assistant', text: '' },
      ...(['a', 'b', 'c', 'd'] as const).map((ch, i): Message => ({
        id: `tr${i}`, sessionId: 's1', createdAt: '', role: 'tool', toolCallId: `tc${i}`,
        result: { toolCallId: `tc${i}`, ok: true, data: big(ch) },
      })),
    ];
    const req = cb.build(SESSION, msgs, [], tightCfg);
    const toolMsgs = req.messages.filter(m => m.role === 'tool');
    expect(toolMsgs).toHaveLength(4);
    // Older results squashed to ~2000 chars; the last two stay full-size.
    expect(toolMsgs[0].content).toContain('[truncated]');
    expect(toolMsgs[1].content).toContain('[truncated]');
    expect((toolMsgs[0].content as string).length).toBeLessThan(2100);
    expect((toolMsgs[1].content as string).length).toBeLessThan(2100);
    expect((toolMsgs[2].content as string).length).toBeGreaterThan(19_000);
    expect((toolMsgs[3].content as string).length).toBeGreaterThan(19_000);
  });

  it('unconditionally squashes old tool results even when the full history fits the budget', () => {
    const cb = new ContextBuilder(makeRegistry());
    // Very large budget — everything fits without compaction under the old logic.
    const largeCfg: ProviderConfig = {
      ...CFG,
      contextLimits: { ...CFG.contextLimits, maxContextTokens: 1_000_000 },
    };
    const msgs: Message[] = [
      { id: 'u1', sessionId: 's1', createdAt: '', role: 'user', text: 'Fetch pages' },
      { id: 'a1', sessionId: 's1', createdAt: '', role: 'assistant', text: '' },
      ...(['a', 'b', 'c', 'd'] as const).map((_, i): Message => ({
        id: `tr${i}`, sessionId: 's1', createdAt: '', role: 'tool', toolCallId: `tc${i}`,
        result: { toolCallId: `tc${i}`, ok: true, data: 'x'.repeat(20_000) },
      })),
    ];
    const req = cb.build(SESSION, msgs, [], largeCfg);
    const toolMsgs = req.messages.filter(m => m.role === 'tool');
    expect(toolMsgs).toHaveLength(4);
    // First two (old) results must be squashed to ≤2000 chars.
    expect((toolMsgs[0].content as string).length).toBeLessThan(2100);
    expect((toolMsgs[1].content as string).length).toBeLessThan(2100);
    expect(toolMsgs[0].content).toContain('…[truncated]');
    expect(toolMsgs[1].content).toContain('…[truncated]');
    // Last two (recent) results must be preserved verbatim.
    expect((toolMsgs[2].content as string).length).toBeGreaterThan(19_000);
    expect((toolMsgs[3].content as string).length).toBeGreaterThan(19_000);
    expect(toolMsgs[3].content).not.toContain('[truncated]');
  });

  it('input token estimate grows sublinearly as old tool results accumulate', () => {
    const cb = new ContextBuilder(makeRegistry());
    const largeCfg: ProviderConfig = {
      ...CFG,
      contextLimits: { ...CFG.contextLimits, maxContextTokens: 1_000_000 },
    };

    function makeGrowingTranscript(n: number): Message[] {
      const msgs: Message[] = [
        { id: 'u0', sessionId: 's1', createdAt: '', role: 'user', text: 'task' },
      ];
      for (let i = 0; i < n; i++) {
        msgs.push({
          id: `a${i}`, sessionId: 's1', createdAt: '', role: 'assistant', text: `reply ${i}`,
          toolCalls: [{ id: `tc${i}`, name: 'fetch_url', arguments: {}, workbookId: 'wb1', mutating: false }],
        });
        msgs.push({
          id: `tr${i}`, sessionId: 's1', createdAt: '', role: 'tool', toolCallId: `tc${i}`,
          result: { toolCallId: `tc${i}`, ok: true, data: 'x'.repeat(20_000) },
        });
      }
      return msgs;
    }

    const tokens5 = cb.estimateInputTokens(SESSION, makeGrowingTranscript(5), [], largeCfg);
    const tokens30 = cb.estimateInputTokens(SESSION, makeGrowingTranscript(30), [], largeCfg);
    // Without proactive compaction, 30 large results would cost ~6× what 5 results cost.
    // With it, old results collapse to ≤2000 chars so growth is sublinear.
    expect(tokens30).toBeLessThan(tokens5 * 3);
  });

  it('estimateInputTokens returns a positive number', () => {
    const cb = new ContextBuilder(makeRegistry());
    const msgs = makeMessages(['test', 'ok']);
    const tokens = cb.estimateInputTokens(SESSION, msgs, [], CFG);
    expect(tokens).toBeGreaterThan(10);
  });
});
