import { describe, it, expect, beforeEach } from 'vitest';
import { AgentLoop } from '../loop';
import { WorkbookRegistry } from '../../workbook/registry';
import { SnapshotManager } from '../../workbook/snapshot';
import { useStore } from '../../store/index';
import type { LLMClient, LLMStreamEvent, LLMRequest, ProviderConfig } from '../../types';
import type { ToolSpec, ToolResult } from '../../types';
import type { ToolExecutor } from '../../workbook/executor';
import { WEB_SEARCH } from '../../web/search';

// ── Helpers ────────────────────────────────────────────────────────────────

function makeRegistry(wbId = 'wb1'): WorkbookRegistry {
  const r = new WorkbookRegistry();
  (r as unknown as { handles: Map<string, unknown>; activeId: string; hostId: string }).handles.set(wbId, {
    workbookId: wbId, name: 'Test.xlsx', isActive: true, isHost: true,
    sheets: [{ name: 'Sheet1', position: 0, visible: true }],
    lastRefreshed: new Date().toISOString(), capability: 'host-only',
  });
  (r as unknown as { activeId: string; hostId: string }).activeId = wbId;
  (r as unknown as { hostId: string }).hostId = wbId;
  return r;
}

function makeExecutor(specs: ToolSpec[] = [], execResult?: ToolResult): Partial<ToolExecutor> {
  return {
    getToolSpecs: () => specs,
    execute: async (call) => execResult ?? { toolCallId: call.id, ok: true, data: { result: 'ok' } },
  };
}

async function* textStream(text: string): AsyncIterable<LLMStreamEvent> {
  yield { type: 'text-delta', delta: text };
  yield { type: 'usage', inputTokens: 10, outputTokens: 5, source: 'estimated' };
  yield { type: 'done', finishReason: 'stop' };
}

async function* toolCallStream(toolId: string, toolName: string, args: string): AsyncIterable<LLMStreamEvent> {
  yield { type: 'tool-call-start', index: 0, id: toolId, name: toolName };
  yield { type: 'tool-call-delta', index: 0, argumentsDelta: args };
  yield { type: 'tool-call-end', index: 0 };
  yield { type: 'usage', inputTokens: 20, outputTokens: 8, source: 'estimated' };
  yield { type: 'done', finishReason: 'tool_calls' };
}

function makeSingleTurnClient(events: AsyncIterable<LLMStreamEvent>): LLMClient {
  return {
    chat: () => events,
    listModels: async () => [],
    capabilities: () => ({ supportsTools: true, supportsStreaming: true, supportsOAuth: false, nativeUsage: false, toolFormat: 'openai' as const }),
  };
}

function makeTwoTurnClient(
  firstEvents: AsyncIterable<LLMStreamEvent>,
  secondEvents: AsyncIterable<LLMStreamEvent>
): LLMClient {
  let call = 0;
  return {
    chat: (_req: LLMRequest) => {
      call++;
      return call === 1 ? firstEvents : secondEvents;
    },
    listModels: async () => [],
    capabilities: () => ({ supportsTools: true, supportsStreaming: true, supportsOAuth: false, nativeUsage: false, toolFormat: 'openai' as const }),
  };
}

function makeCapturingClient(
  requests: LLMRequest[],
  events: AsyncIterable<LLMStreamEvent> = textStream('ok')
): LLMClient {
  return {
    chat: (req: LLMRequest) => {
      requests.push(req);
      return events;
    },
    listModels: async () => [],
    capabilities: () => ({ supportsTools: true, supportsStreaming: true, supportsOAuth: false, nativeUsage: false, toolFormat: 'openai' as const }),
  };
}

function makeMultiTurnClient(events: AsyncIterable<LLMStreamEvent>[]): LLMClient {
  let call = 0;
  return {
    chat: () => events[Math.min(call++, events.length - 1)],
    listModels: async () => [],
    capabilities: () => ({ supportsTools: true, supportsStreaming: true, supportsOAuth: false, nativeUsage: false, toolFormat: 'openai' as const }),
  };
}

const CFG: ProviderConfig = {
  provider: 'ollama', enabled: true,
  baseUrl: 'http://localhost:11434', model: 'llama3.2',
  authMode: 'none', authStateRef: '',
  contextLimits: { maxContextTokens: 128000, historyTokenCap: 100000, maxInlineSheetCells: 5000 },
  maxOutputTokens: 4096,
};

const GENERIC_CFG: ProviderConfig = {
  ...CFG,
  provider: 'generic',
  label: 'Generic / OpenRouter',
  baseUrl: 'https://openrouter.ai/api/v1',
  model: 'openai/gpt-4o-mini',
  authMode: 'oauth',
  authStateRef: 'xl.auth.generic',
};

const KIMI_CFG: ProviderConfig = {
  ...CFG,
  provider: 'kimi',
  label: 'Kimi',
  baseUrl: 'https://api.moonshot.ai/v1',
  model: 'kimi-k2.6',
  authMode: 'apikey',
  authStateRef: 'xl.auth.kimi',
};

const SCOPE = { workbookId: 'wb1' };

async function waitForStatus(status: string): Promise<void> {
  const start = Date.now();
  while (Date.now() - start < 1000) {
    if (useStore.getState().currentSession?.status === status) return;
    await new Promise(r => setTimeout(r, 5));
  }
  const state = useStore.getState();
  throw new Error(`Timed out waiting for status ${status}; current=${state.currentSession?.status}; messages=${JSON.stringify(state.messages)}`);
}

function noop<T>(fn: (ctx: Excel.RequestContext) => Promise<T>): Promise<T> {
  return fn({} as Excel.RequestContext);
}

// ── Reset store between tests ──────────────────────────────────────────────

beforeEach(() => {
  useStore.getState().setSession(null);
  useStore.getState().clearMessages();
  useStore.getState().setWebSearchEnabled(false);
  useStore.getState().setAppConfig({ webAccess: { provider: 'none', readerFallback: false } });
});

// ── Tests ──────────────────────────────────────────────────────────────────

describe('AgentLoop — text-only run', () => {
  it('creates session, appends user + assistant message, ends with status done', async () => {
    const loop = new AgentLoop(
      makeRegistry(),
      makeExecutor() as unknown as ToolExecutor,
      new SnapshotManager(),
      noop
    );

    const client = makeSingleTurnClient(textStream('Hello from LLM!'));
    await loop.start('Say hi', SCOPE, client, CFG);

    const state = useStore.getState();
    expect(state.currentSession?.status).toBe('done');

    const msgs = state.messages;
    const userMsg = msgs.find(m => m.role === 'user');
    const assistantMsg = msgs.find(m => m.role === 'assistant');

    expect(userMsg).toBeDefined();
    expect((userMsg as { text: string }).text).toBe('Say hi');
    expect(assistantMsg).toBeDefined();
    expect((assistantMsg as { text: string }).text).toBe('Hello from LLM!');
  });

  it('accumulates token totals', async () => {
    const loop = new AgentLoop(
      makeRegistry(),
      makeExecutor() as unknown as ToolExecutor,
      new SnapshotManager(),
      noop
    );

    await loop.start('Hello', SCOPE, makeSingleTurnClient(textStream('Hi')), CFG);

    const totals = useStore.getState().currentSession?.totals;
    expect(totals?.inputTokens).toBe(10);
    expect(totals?.outputTokens).toBe(5);
  });

  it('adds follow-up prompts to the same session instead of replacing the chat', async () => {
    const requests: LLMRequest[] = [];
    let call = 0;
    const client: LLMClient = {
      chat: (req: LLMRequest) => {
        requests.push(req);
        return call++ === 0 ? textStream('First answer') : textStream('Second answer');
      },
      listModels: async () => [],
      capabilities: () => ({ supportsTools: true, supportsStreaming: true, supportsOAuth: false, nativeUsage: false, toolFormat: 'openai' as const }),
    };
    const loop = new AgentLoop(
      makeRegistry(),
      makeExecutor() as unknown as ToolExecutor,
      new SnapshotManager(),
      noop
    );

    await loop.start('First question', SCOPE, client, CFG);
    const sessionId = useStore.getState().currentSession?.id;
    await loop.followUp('Follow-up question', SCOPE, client, CFG);

    const state = useStore.getState();
    expect(state.currentSession?.id).toBe(sessionId);
    expect(state.messages.filter(m => m.role === 'user').map(m => (m as { text: string }).text)).toEqual([
      'First question',
      'Follow-up question',
    ]);
    expect(state.messages.filter(m => m.role === 'assistant').map(m => (m as { text: string }).text)).toEqual([
      'First answer',
      'Second answer',
    ]);
    expect(requests[1].messages.filter(m => m.role === 'user').map(m => m.content)).toEqual([
      expect.stringContaining('First question'),
      'Follow-up question',
    ]);
  });

  it('isRunning returns false after completion', async () => {
    const loop = new AgentLoop(
      makeRegistry(),
      makeExecutor() as unknown as ToolExecutor,
      new SnapshotManager(),
      noop
    );
    await loop.start('Test', SCOPE, makeSingleTurnClient(textStream('ok')), CFG);
    expect(loop.isRunning()).toBe(false);
  });
});

describe('AgentLoop - web search gating', () => {
  const readRange: ToolSpec = {
    name: 'read_range',
    description: 'Read cells',
    parameters: { type: 'object', properties: {}, required: [] },
    mutating: false,
  };
  const fetchUrl: ToolSpec = {
    name: 'fetch_url',
    description: 'Fetch URL',
    parameters: { type: 'object', properties: {}, required: [] },
    mutating: false,
    runtime: 'none',
  };
  const webTools = [readRange, WEB_SEARCH, fetchUrl];

  it('keeps Doc 11 behavior for BYOK-tier providers with a configured search provider', async () => {
    const requests: LLMRequest[] = [];
    const loop = new AgentLoop(
      makeRegistry(),
      makeExecutor(webTools) as unknown as ToolExecutor,
      new SnapshotManager(),
      noop
    );

    useStore.getState().setAppConfig({ webAccess: { provider: 'wikipedia', readerFallback: false } });
    useStore.getState().setWebSearchEnabled(true);

    await loop.start('Search with BYOK tier', SCOPE, makeCapturingClient(requests), CFG);

    expect(requests[0].tools.map(t => t.name)).toEqual([
      'read_range',
      'web_search',
      'fetch_url',
      'request_user_choice',
    ]);
    expect(useStore.getState().currentSession?.webSearchEnabled).toBe(true);
  });

  it('suppresses client web_search on the native tier even when a BYOK provider is configured', async () => {
    const requests: LLMRequest[] = [];
    const loop = new AgentLoop(
      makeRegistry(),
      makeExecutor(webTools) as unknown as ToolExecutor,
      new SnapshotManager(),
      noop
    );

    useStore.getState().setAppConfig({ webAccess: { provider: 'wikipedia', readerFallback: false } });
    useStore.getState().setWebSearchEnabled(true);

    await loop.start('Search with native tier', SCOPE, makeCapturingClient(requests), GENERIC_CFG);

    expect(requests[0].tools.map(t => t.name)).toEqual([
      'read_range',
      'fetch_url',
      'request_user_choice',
    ]);
    expect(useStore.getState().currentSession?.webSearchEnabled).toBe(true);
  });

  it('removes web tools when BYOK-tier search is unavailable even if the session toggle was on', async () => {
    const requests: LLMRequest[] = [];
    const loop = new AgentLoop(
      makeRegistry(),
      makeExecutor(webTools) as unknown as ToolExecutor,
      new SnapshotManager(),
      noop
    );

    useStore.getState().setWebSearchEnabled(true);

    await loop.start('Search unavailable', SCOPE, makeCapturingClient(requests), CFG);

    expect(requests[0].tools.map(t => t.name)).toEqual(['read_range', 'request_user_choice']);
    expect(useStore.getState().currentSession?.webSearchEnabled).toBe(false);
  });

  it('exposes client web tools for keyless BYOK providers even when the toggle is off', async () => {
    const byokRequests: LLMRequest[] = [];

    useStore.getState().setAppConfig({ webAccess: { provider: 'wikipedia', readerFallback: false } });

    await new AgentLoop(
      makeRegistry(),
      makeExecutor(webTools) as unknown as ToolExecutor,
      new SnapshotManager(),
      noop
    ).start('BYOK off', SCOPE, makeCapturingClient(byokRequests), CFG);

    expect(byokRequests[0].tools.map(t => t.name)).toEqual(['read_range', 'web_search', 'fetch_url', 'request_user_choice']);
    expect(useStore.getState().currentSession?.webSearchEnabled).toBe(false);
  });

  it('removes web tools when the toggle is off for native and keyed BYOK tiers', async () => {
    const keyedRequests: LLMRequest[] = [];
    const nativeRequests: LLMRequest[] = [];

    useStore.getState().setAppConfig({ webAccess: { provider: 'tavily', readerFallback: false } });

    await new AgentLoop(
      makeRegistry(),
      makeExecutor(webTools) as unknown as ToolExecutor,
      new SnapshotManager(),
      noop
    ).start('Keyed BYOK off', SCOPE, makeCapturingClient(keyedRequests), CFG);

    await new AgentLoop(
      makeRegistry(),
      makeExecutor(webTools) as unknown as ToolExecutor,
      new SnapshotManager(),
      noop
    ).start('Native off', SCOPE, makeCapturingClient(nativeRequests), GENERIC_CFG);

    expect(keyedRequests[0].tools.map(t => t.name)).toEqual(['read_range', 'request_user_choice']);
    expect(nativeRequests[0].tools.map(t => t.name)).toEqual(['read_range', 'request_user_choice']);
  });

  it('echoes Kimi $web_search arguments without executing a local tool', async () => {
    const requests: LLMRequest[] = [];
    const executor = {
      getToolSpecs: () => webTools,
      execute: async () => {
        throw new Error('Kimi $web_search should not reach the executor');
      },
    };
    let turn = 0;
    const client: LLMClient = {
      chat: (req: LLMRequest) => {
        requests.push(req);
        return turn++ === 0
          ? toolCallStream('kimi_search_1', '$web_search', '{"query":"latest SheetClaw news"}')
          : textStream('Search complete.');
      },
      listModels: async () => [],
      capabilities: () => ({ supportsTools: true, supportsStreaming: true, supportsOAuth: false, nativeUsage: false, toolFormat: 'openai' as const }),
    };

    useStore.getState().setWebSearchEnabled(true);

    await new AgentLoop(
      makeRegistry(),
      executor as unknown as ToolExecutor,
      new SnapshotManager(),
      noop
    ).start('Search with Kimi', SCOPE, client, KIMI_CFG);

    const toolMessage = useStore.getState().messages.find(m => m.role === 'tool') as { result: ToolResult };
    expect(toolMessage.result.ok).toBe(true);
    expect(toolMessage.result.data).toBe('{"query":"latest SheetClaw news"}');

    const secondRequestToolResult = requests[1].messages.find(m => m.role === 'tool');
    expect(secondRequestToolResult).toEqual({
      role: 'tool',
      toolCallId: 'kimi_search_1',
      name: '$web_search',
      content: '{"query":"latest SheetClaw news"}',
    });
  });
});

describe('AgentLoop — non-mutating tool call', () => {
  it('executes tool, appends result, makes second LLM call, ends done', async () => {
    const spec: ToolSpec = {
      name: 'read_range',
      description: 'Read cells',
      parameters: { type: 'object', properties: { workbook_id: { type: 'string' }, sheet: { type: 'string' }, address: { type: 'string' } }, required: ['workbook_id', 'sheet', 'address'] },
      mutating: false,
    };
    const executor = makeExecutor([spec], { toolCallId: 'tc1', ok: true, data: { values: [['A']] } });

    const client = makeTwoTurnClient(
      toolCallStream('tc1', 'read_range', JSON.stringify({ workbook_id: 'wb1', sheet: 'Sheet1', address: 'A1' })),
      textStream('Cell A1 contains "A".')
    );

    const loop = new AgentLoop(
      makeRegistry(),
      executor as unknown as ToolExecutor,
      new SnapshotManager(),
      noop
    );

    await loop.start('What is in A1?', SCOPE, client, CFG);

    const state = useStore.getState();
    expect(state.currentSession?.status).toBe('done');

    const toolResultMsg = state.messages.find(m => m.role === 'tool');
    expect(toolResultMsg).toBeDefined();
    expect((toolResultMsg as { result: ToolResult }).result.ok).toBe(true);

    const assistantMsgs = state.messages.filter(m => m.role === 'assistant');
    const finalAssistant = assistantMsgs[assistantMsgs.length - 1];
    expect((finalAssistant as { text: string }).text).toBe('Cell A1 contains "A".');
  });

  it('can continue the same session after the iteration cap is reached', async () => {
    const spec: ToolSpec = {
      name: 'read_range',
      description: 'Read cells',
      parameters: { type: 'object', properties: { workbook_id: { type: 'string' }, sheet: { type: 'string' }, address: { type: 'string' } }, required: ['workbook_id', 'sheet', 'address'] },
      mutating: false,
    };
    const executor = makeExecutor([spec], { toolCallId: 'tc', ok: true, data: { values: [['A']] } });
    const events = [
      ...Array.from({ length: 26 }, (_, i) =>
        toolCallStream(`tc${i}`, 'read_range', JSON.stringify({ workbook_id: 'wb1', sheet: 'Sheet1', address: 'A1' }))
      ),
      textStream('Now complete.'),
    ];
    const client = makeMultiTurnClient(events);
    const loop = new AgentLoop(
      makeRegistry(),
      executor as unknown as ToolExecutor,
      new SnapshotManager(),
      noop
    );

    await loop.start('Long task', SCOPE, client, CFG);

    let state = useStore.getState();
    expect(state.currentSession?.status).toBe('done');
    expect(state.currentSession?.stopReason).toBe('max_iterations');
    expect(state.currentSession?.iteration).toBe(25);
    expect(state.currentSession?.maxIterations).toBe(25);

    await loop.continueCurrent(client, CFG);

    state = useStore.getState();
    expect(state.currentSession?.status).toBe('done');
    expect(state.currentSession?.stopReason).toBeUndefined();
    expect(state.currentSession?.maxIterations).toBe(50);
    const assistantMessages = state.messages.filter(m => m.role === 'assistant') as Array<{ text: string }>;
    const final = assistantMessages[assistantMessages.length - 1];
    expect(final.text).toBe('Now complete.');
    expect(state.messages.filter(m => m.role === 'user')).toHaveLength(1);
  });
});

describe('AgentLoop — parallel network tool execution', () => {
  it('runs consecutive read-only network calls concurrently and appends results in call order', async () => {
    const fetchSpec: ToolSpec = {
      name: 'fetch_url',
      description: 'Fetch URL',
      parameters: { type: 'object', properties: {}, required: [] },
      mutating: false,
      runtime: 'none',
    };

    let active = 0;
    let maxActive = 0;
    const executor: Partial<ToolExecutor> = {
      getToolSpecs: () => [fetchSpec],
      execute: async (call) => {
        active++;
        maxActive = Math.max(maxActive, active);
        await new Promise(r => setTimeout(r, 25));
        active--;
        return { toolCallId: call.id, ok: true, data: { id: call.id } };
      },
    };

    async function* twoCalls(): AsyncIterable<LLMStreamEvent> {
      yield { type: 'tool-call-start', index: 0, id: 'p1', name: 'fetch_url' };
      yield { type: 'tool-call-delta', index: 0, argumentsDelta: '{"url":"https://finance.example/"}' };
      yield { type: 'tool-call-start', index: 1, id: 'p2', name: 'fetch_url' };
      yield { type: 'tool-call-delta', index: 1, argumentsDelta: '{"url":"https://weather.example/"}' };
      yield { type: 'usage', inputTokens: 20, outputTokens: 8, source: 'estimated' };
      yield { type: 'done', finishReason: 'tool_calls' };
    }

    useStore.getState().setAppConfig({ webAccess: { provider: 'wikipedia', readerFallback: false } });
    useStore.getState().setWebSearchEnabled(true);

    const loop = new AgentLoop(
      makeRegistry(),
      executor as unknown as ToolExecutor,
      new SnapshotManager(),
      noop
    );
    await loop.start('Fetch two pages', SCOPE, makeTwoTurnClient(twoCalls(), textStream('Both fetched.')), CFG);

    expect(useStore.getState().currentSession?.status).toBe('done');
    expect(maxActive).toBe(2);
    const toolMsgs = useStore.getState().messages.filter(m => m.role === 'tool') as Array<{ toolCallId: string }>;
    expect(toolMsgs.map(m => m.toolCallId)).toEqual(['p1', 'p2']);
  });
});

describe('AgentLoop — stop()', () => {
  it('stops a running loop and isRunning() returns false after the promise settles', async () => {
    let yieldControl!: () => void;
    async function* slowStream(): AsyncIterable<LLMStreamEvent> {
      await new Promise<void>(r => { yieldControl = r; });
      yield { type: 'done', finishReason: 'stop' };
    }

    const loop = new AgentLoop(
      makeRegistry(),
      makeExecutor() as unknown as ToolExecutor,
      new SnapshotManager(),
      noop
    );

    expect(loop.isRunning()).toBe(false);
    const p = loop.start('Slow task', SCOPE, makeSingleTurnClient(slowStream()), CFG);
    expect(loop.isRunning()).toBe(true);

    loop.stop();
    yieldControl();
    await p;

    // After stop + resolve, the loop must have terminated
    expect(loop.isRunning()).toBe(false);
    // Status is either 'done' (abort caught mid-stream non-throwing path) or 'stopped'
    const status = useStore.getState().currentSession?.status;
    expect(['done', 'stopped']).toContain(status);
  });
});

describe('AgentLoop — error handling', () => {
  it('catches LLM error event and sets session status to error', async () => {
    async function* errorStream(): AsyncIterable<LLMStreamEvent> {
      yield { type: 'error', error: { code: 'AuthError', message: 'Invalid API key' } };
    }

    const loop = new AgentLoop(
      makeRegistry(),
      makeExecutor() as unknown as ToolExecutor,
      new SnapshotManager(),
      noop
    );

    await loop.start('Hello', SCOPE, makeSingleTurnClient(errorStream()), CFG);

    const state = useStore.getState();
    expect(state.currentSession?.status).toBe('error');
    expect(state.currentSession?.lastError?.message).toContain('AuthError');
  });
});

describe('AgentLoop - request_user_choice', () => {
  it('feeds malformed choice args back as ValidationError, then succeeds after corrected args', async () => {
    const client = makeMultiTurnClient([
      toolCallStream('bad_choice', 'request_user_choice', JSON.stringify({
        question: 'Choose one',
        options: ['Only one'],
      })),
      toolCallStream('good_choice', 'request_user_choice', JSON.stringify({
        question: 'Choose one',
        options: [{ id: 'a', label: 'A' }, { id: 'b', label: 'B' }],
      })),
      textStream('Continuing with A.'),
    ]);
    const loop = new AgentLoop(
      makeRegistry(),
      makeExecutor() as unknown as ToolExecutor,
      new SnapshotManager(),
      noop
    );

    const run = loop.start('Need a choice', SCOPE, client, CFG);
    await waitForStatus('awaiting_choice');
    loop.resolveChoice(['a']);
    await run;

    const toolMessages = useStore.getState().messages.filter(m => m.role === 'tool') as Array<{ result: ToolResult }>;
    expect(toolMessages[0].result.ok).toBe(false);
    expect(toolMessages[0].result.error?.code).toBe('ValidationError');
    expect(toolMessages[1].result.ok).toBe(true);
    expect(toolMessages[1].result.data).toMatchObject({ selected_ids: ['a'] });
  });

  it('leaves prose-only clarification as a normal text turn with no menu', async () => {
    const loop = new AgentLoop(
      makeRegistry(),
      makeExecutor() as unknown as ToolExecutor,
      new SnapshotManager(),
      noop
    );

    await loop.start('Ambiguous request', SCOPE, makeSingleTurnClient(textStream('Which source should I use?')), CFG);

    const state = useStore.getState();
    expect(state.currentSession?.status).toBe('done');
    expect(state.currentSession?.pendingChoice).toBeUndefined();
    expect(state.messages.some(m => m.role === 'tool')).toBe(false);
  });

  it('selection and dismiss resolve as tool results without synthetic user messages', async () => {
    const choiceArgs = JSON.stringify({
      question: 'Pick a source',
      options: [{ id: 'left', label: 'Left' }, { id: 'right', label: 'Right' }],
    });
    const selectLoop = new AgentLoop(
      makeRegistry(),
      makeExecutor() as unknown as ToolExecutor,
      new SnapshotManager(),
      noop
    );
    let run = selectLoop.start('Pick', SCOPE, makeTwoTurnClient(
      toolCallStream('choice_select', 'request_user_choice', choiceArgs),
      textStream('Done.')
    ), CFG);
    await waitForStatus('awaiting_choice');
    selectLoop.resolveChoice(['right']);
    await run;

    let state = useStore.getState();
    let userMessages = state.messages.filter(m => m.role === 'user') as Array<{ text: string }>;
    expect(userMessages.map(m => m.text)).toEqual(['Pick']);
    let toolMessage = state.messages.find(m => m.role === 'tool') as { result: ToolResult };
    expect(toolMessage.result.data).toMatchObject({ selected_ids: ['right'] });

    useStore.getState().setSession(null);
    useStore.getState().clearMessages();

    const dismissLoop = new AgentLoop(
      makeRegistry(),
      makeExecutor() as unknown as ToolExecutor,
      new SnapshotManager(),
      noop
    );
    run = dismissLoop.start('Pick', SCOPE, makeTwoTurnClient(
      toolCallStream('choice_dismiss', 'request_user_choice', choiceArgs),
      textStream('Stopped.')
    ), CFG);
    await waitForStatus('awaiting_choice');
    dismissLoop.resolveChoice('dismiss');
    await run;

    state = useStore.getState();
    userMessages = state.messages.filter(m => m.role === 'user') as Array<{ text: string }>;
    expect(userMessages.map(m => m.text)).toEqual(['Pick']);
    toolMessage = state.messages.find(m => m.role === 'tool') as { result: ToolResult };
    expect(toolMessage.result.ok).toBe(false);
    expect(toolMessage.result.error?.code).toBe('PermissionDenied');
  });

  it('adds a required Other option and returns its custom text when selected', async () => {
    const choiceArgs = JSON.stringify({
      question: 'Pick a scope',
      options: [{ id: 'country', label: 'Country' }, { id: 'region', label: 'Region' }],
    });
    const loop = new AgentLoop(
      makeRegistry(),
      makeExecutor() as unknown as ToolExecutor,
      new SnapshotManager(),
      noop
    );

    const run = loop.start('Pick', SCOPE, makeTwoTurnClient(
      toolCallStream('choice_other', 'request_user_choice', choiceArgs),
      textStream('Done.')
    ), CFG);
    await waitForStatus('awaiting_choice');

    const pendingChoice = useStore.getState().currentSession?.pendingChoice;
    const lastOption = pendingChoice ? pendingChoice.options[pendingChoice.options.length - 1] : undefined;
    expect(lastOption).toMatchObject({
      id: 'other',
      label: 'Other',
      requiresText: true,
    });

    loop.resolveChoice({ ids: ['other'], otherText: 'Use my custom country list.' });
    await run;

    const toolMessage = useStore.getState().messages.find(m => m.role === 'tool') as { result: ToolResult };
    expect(toolMessage.result.data).toMatchObject({
      selected_ids: ['other'],
      other_text: 'Use my custom country list.',
    });
  });
});
