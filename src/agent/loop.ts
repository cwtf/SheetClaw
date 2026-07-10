import { ulid } from 'ulid';
import type {
  LLMClient,
  ToolCall,
  AgentSession,
  SessionScope,
  ProviderConfig,
  Message,
  UserMessage,
  AssistantMessage,
  ToolCallMessage,
  ToolResultMessage,
  SystemNoticeMessage,
} from '../types';
import type { WorkbookRegistry } from '../workbook/registry';
import { ToolExecutor } from '../workbook/executor';
import { SnapshotManager } from '../workbook/snapshot';
import { ContextBuilder } from './context-builder';
import { computeRangeDiff } from '../workbook/a1notation';
import { useStore } from '../store/index';
import { findPricing, computeCost } from '../pricing/index';
import { filterToolsForRun } from './tool-filter';
import { REQUEST_USER_CHOICE, parsePendingChoice } from './choice';
import { ToolValidationError } from '../workbook/executor';
import { resolveSearchToggle } from '../adapters/native-search';
import { isKeylessSearchProvider, type WebAccessProvider } from '../web/providers';

const MAX_ITERATIONS = 50;
const ACTIVE_STATUSES = new Set<AgentSession['status']>([
  'building',
  'calling_llm',
  'parsing',
  'awaiting_confirmation',
  'awaiting_choice',
  'executing_tool',
]);
export type LoopRunner = <T>(fn: (ctx: Excel.RequestContext) => Promise<T>) => Promise<T>;
export interface ChoiceSelection {
  ids: string[];
  otherText?: string;
}

// ── Tool-call accumulator for streaming ────────────────────────────────────

interface StreamResult {
  streamMsgId: string;
  text: string;
  toolCalls: Array<{ id: string; name: string; args: Record<string, unknown>; rawArgs: string }>;
  finishReason: string;
  inputTokens: number;
  outputTokens: number;
}

// ── AgentLoop ──────────────────────────────────────────────────────────────

export class AgentLoop {
  private abortController: AbortController | null = null;
  private confirmationResolve: ((d: 'apply' | 'cancel') => void) | null = null;
  private choiceResolve: ((d: { kind: 'select'; selection: ChoiceSelection } | { kind: 'dismiss' }) => void) | null = null;
  private runner: LoopRunner;

  constructor(
    private registry: WorkbookRegistry,
    private executor: ToolExecutor,
    private snapshots: SnapshotManager,
    runner?: LoopRunner
  ) {
    this.runner = runner ?? (fn => Excel.run(fn));
  }

  // ── Public ─────────────────────────────────────────────────────────────

  async start(
    instruction: string,
    scope: SessionScope,
    client: LLMClient,
    cfg: ProviderConfig
  ): Promise<void> {
    const ac = new AbortController();
    this.abortController = ac;
    const store = useStore.getState();
    const webProvider = store.appConfig.webAccess.provider;
    const byokReady = webProvider !== 'none' && store.isSearchProviderReady(webProvider);
    const searchToggle = resolveSearchToggle({ provider: cfg.provider, model: cfg.model, byokReady });
    const forceClientWebSearch = shouldForceClientWebSearch(webProvider, byokReady);
    const manualWebSearchEnabled = !forceClientWebSearch && store.webSearchEnabled && searchToggle.available;

    const session: AgentSession = {
      id: ulid(),
      createdAt: new Date().toISOString(),
      scope,
      status: 'building',
      iteration: 0,
      maxIterations: MAX_ITERATIONS,
      provider: cfg.provider,
      model: cfg.model,
      messageIds: [],
      tokenBudget: { used: 0, window: cfg.contextLimits.maxContextTokens },
      webSearchEnabled: manualWebSearchEnabled,
      totals: { inputTokens: 0, outputTokens: 0, costUsd: 0 },
    };

    store.setSession(session);
    store.resetSessionTotals(session.id);

    this.append(session.id, msg<UserMessage>(session.id, { role: 'user', text: instruction }));

    const ctxBuilder = new ContextBuilder(this.registry, readerFallbackEnabled);

    try {
      await this.loop(session, client, cfg, ctxBuilder, ac.signal);
    } catch (e) {
      if (ac.signal.aborted) {
        useStore.getState().updateSessionById(session.id, { status: 'stopped' });
      } else {
        const message = e instanceof Error ? e.message : String(e);
        useStore.getState().updateSessionById(session.id, { status: 'error', lastError: { code: 'LoopError', message } });
        this.append(session.id, msg<SystemNoticeMessage>(session.id, { role: 'system_notice', level: 'error', text: `Run failed: ${message}` }));
      }
    } finally {
      this.abortController = null;
    }
  }

  async followUp(
    instruction: string,
    scope: SessionScope,
    client: LLMClient,
    cfg: ProviderConfig
  ): Promise<void> {
    const current = useStore.getState().currentSession;
    if (!current) {
      await this.start(instruction, scope, client, cfg);
      return;
    }
    if (ACTIVE_STATUSES.has(current.status)) return;

    const ac = new AbortController();
    this.abortController = ac;
    const store = useStore.getState();
    const webProvider = store.appConfig.webAccess.provider;
    const byokReady = webProvider !== 'none' && store.isSearchProviderReady(webProvider);
    const searchToggle = resolveSearchToggle({ provider: cfg.provider, model: cfg.model, byokReady });
    const forceClientWebSearch = shouldForceClientWebSearch(webProvider, byokReady);
    const manualWebSearchEnabled = !forceClientWebSearch && store.webSearchEnabled && searchToggle.available;
    const session: AgentSession = {
      ...current,
      scope,
      status: 'building',
      iteration: 0,
      maxIterations: MAX_ITERATIONS,
      provider: cfg.provider,
      model: cfg.model,
      pendingChange: undefined,
      pendingChoice: undefined,
      stopReason: undefined,
      lastError: undefined,
      tokenBudget: { used: 0, window: cfg.contextLimits.maxContextTokens },
      webSearchEnabled: manualWebSearchEnabled,
    };

    store.updateSessionById(session.id, {
      scope: session.scope,
      status: session.status,
      iteration: session.iteration,
      maxIterations: session.maxIterations,
      provider: session.provider,
      model: session.model,
      pendingChange: undefined,
      pendingChoice: undefined,
      stopReason: undefined,
      lastError: undefined,
      tokenBudget: session.tokenBudget,
      webSearchEnabled: session.webSearchEnabled,
    });
    this.append(session.id, msg<UserMessage>(session.id, { role: 'user', text: instruction }));

    const ctxBuilder = new ContextBuilder(this.registry, readerFallbackEnabled);
    try {
      await this.loop(session, client, cfg, ctxBuilder, ac.signal);
    } catch (e) {
      if (ac.signal.aborted) {
        useStore.getState().updateSessionById(session.id, { status: 'stopped' });
      } else {
        const message = e instanceof Error ? e.message : String(e);
        useStore.getState().updateSessionById(session.id, { status: 'error', lastError: { code: 'LoopError', message } });
        this.append(session.id, msg<SystemNoticeMessage>(session.id, { role: 'system_notice', level: 'error', text: `Run failed: ${message}` }));
      }
    } finally {
      this.abortController = null;
    }
  }

  stop(): void {
    this.abortController?.abort();
    this.confirmationResolve = null;
    this.choiceResolve = null;
  }

  async continueCurrent(
    client: LLMClient,
    cfg: ProviderConfig,
    additionalIterations = MAX_ITERATIONS
  ): Promise<void> {
    const current = useStore.getState().currentSession;
    if (!current || current.status !== 'done' || current.stopReason !== 'max_iterations') return;

    const ac = new AbortController();
    this.abortController = ac;
    const resumed: AgentSession = {
      ...current,
      status: 'building',
      maxIterations: current.maxIterations + additionalIterations,
      stopReason: undefined,
    };
    useStore.getState().updateSessionById(resumed.id, {
      status: resumed.status,
      maxIterations: resumed.maxIterations,
      stopReason: undefined,
    });
    this.append(resumed.id, msg<SystemNoticeMessage>(resumed.id, {
      role: 'system_notice',
      level: 'info',
      text: `Continuing for ${additionalIterations} more iterations.`,
    }));

    const ctxBuilder = new ContextBuilder(this.registry, readerFallbackEnabled);
    try {
      await this.loop(resumed, client, cfg, ctxBuilder, ac.signal);
    } catch (e) {
      if (ac.signal.aborted) {
        useStore.getState().updateSessionById(resumed.id, { status: 'stopped' });
      } else {
        const message = e instanceof Error ? e.message : String(e);
        useStore.getState().updateSessionById(resumed.id, { status: 'error', lastError: { code: 'LoopError', message } });
        this.append(resumed.id, msg<SystemNoticeMessage>(resumed.id, { role: 'system_notice', level: 'error', text: `Run failed: ${message}` }));
      }
    } finally {
      this.abortController = null;
    }
  }

  resolveConfirmation(decision: 'apply' | 'cancel'): void {
    this.confirmationResolve?.(decision);
    this.confirmationResolve = null;
  }

  resolveChoice(selection: string[] | ChoiceSelection | 'dismiss'): void {
    if (selection === 'dismiss') {
      this.choiceResolve?.({ kind: 'dismiss' });
    } else {
      this.choiceResolve?.({
        kind: 'select',
        selection: Array.isArray(selection) ? { ids: selection } : selection,
      });
    }
    this.choiceResolve = null;
  }

  isRunning(): boolean { return this.abortController !== null; }

  // ── Loop ───────────────────────────────────────────────────────────────

  private async loop(
    session: AgentSession,
    client: LLMClient,
    cfg: ProviderConfig,
    ctxBuilder: ContextBuilder,
    signal: AbortSignal
  ): Promise<void> {
    const store = useStore.getState();
    const webProvider = store.appConfig.webAccess.provider;
    const byokReady = webProvider !== 'none' && store.isSearchProviderReady(webProvider);
    const searchToggle = resolveSearchToggle({ provider: cfg.provider, model: cfg.model, byokReady });
    const forceClientWebSearch = shouldForceClientWebSearch(webProvider, byokReady);
    const toolSpecs = [
      ...filterToolsForRun(this.executor.getToolSpecs(), session.webSearchEnabled, searchToggle, { forceClientWebSearch }),
      REQUEST_USER_CHOICE,
    ];
    // Read-only tools that run outside the Excel runtime (web_search, fetch_url)
    // are safe to execute concurrently; each network call can block for seconds.
    // request_user_choice blocks on user input, so it stays sequential.
    const parallelizable = new Set(
      toolSpecs
        .filter(s => !s.mutating && s.runtime === 'none' && s.name !== REQUEST_USER_CHOICE.name)
        .map(s => s.name)
    );
    // The executor keeps every tool registered; only names advertised in this
    // run may execute, so calls to filtered-out tools (e.g. web tools while web
    // access is unconfigured) fail fast with a corrective message.
    const allowedTools = new Set(toolSpecs.map(s => s.name));

    for (let iter = session.iteration; iter < session.maxIterations; iter++) {
      if (signal.aborted) return;
      useStore.getState().updateSessionById(session.id, { iteration: iter + 1, status: 'building' });

      const messages = useStore.getState().messages.filter(
        m => (m as Message & { sessionId: string }).sessionId === session.id
      );
      const req = ctxBuilder.build(session, messages, toolSpecs, cfg);
      useStore.getState().updateSessionById(session.id, { status: 'calling_llm' });

      const sr = await this.stream(session, client, req, signal);

      // Record per-turn usage and update session totals
      const pricingEntry = findPricing(cfg.provider, cfg.model);
      const costUsd = computeCost({ inputTokens: sr.inputTokens, outputTokens: sr.outputTokens }, pricingEntry);
      useStore.getState().recordUsage({
        id: ulid(),
        sessionId: session.id,
        turnIndex: iter,
        timestamp: new Date().toISOString(),
        provider: cfg.provider as import('../types').ProviderKey,
        model: cfg.model,
        inputTokens: sr.inputTokens,
        outputTokens: sr.outputTokens,
        totalTokens: sr.inputTokens + sr.outputTokens,
        estimatedCostUsd: costUsd,
        pricingVersion: pricingEntry ? 'bundled' : 'default',
        estimated: sr.inputTokens === 0 && sr.outputTokens === 0,
        toolCallsCount: sr.toolCalls.length,
      });

      const current = useStore.getState().currentSession;
      const t = current?.id === session.id
        ? current.totals
        : session.totals;
      useStore.getState().updateSessionById(session.id, {
        totals: { ...t, inputTokens: t.inputTokens + sr.inputTokens, outputTokens: t.outputTokens + sr.outputTokens },
      });

      // Resolve mutating flags from tool specs
      const calls: ToolCall[] = sr.toolCalls.map(tc => ({
        id: tc.id,
        name: tc.name,
        arguments: tc.args,
        rawArguments: tc.rawArgs,
        workbookId: session.scope.workbookId,
        mutating: toolSpecs.find(s => s.name === tc.name)?.mutating ?? false,
      }));

      // Finalise the streaming assistant message (update toolCalls in-place)
      useStore.getState().updateMessage(sr.streamMsgId, {
        toolCalls: calls.length ? calls : undefined,
        finishReason: (sr.finishReason as AssistantMessage['finishReason']) ?? 'stop',
      } as Partial<Message>);

      if (!calls.length || sr.finishReason === 'stop') {
        useStore.getState().updateSessionById(session.id, { status: 'done', stopReason: undefined });
        return;
      }
      if (sr.finishReason === 'length') {
        this.append(session.id, msg<SystemNoticeMessage>(session.id, {
          role: 'system_notice', level: 'warn',
          text: 'Response cut off at token limit — the model may not have finished.',
        }));
        useStore.getState().updateSessionById(session.id, { status: 'done', stopReason: undefined });
        return;
      }

      // Execute tool calls — consecutive network reads run concurrently
      useStore.getState().updateSessionById(session.id, { status: 'executing_tool' });
      await this.executeCalls(calls, session, signal, parallelizable, allowedTools);
    }

    useStore.getState().updateSessionById(session.id, { status: 'done', stopReason: 'max_iterations' });
    this.append(session.id, msg<SystemNoticeMessage>(session.id, {
      role: 'system_notice', level: 'warn',
      text: `Stopped after ${MAX_ITERATIONS} iterations. The task may be incomplete.`,
    }));
  }

  // ── Streaming ──────────────────────────────────────────────────────────

  private async stream(
    session: AgentSession,
    client: LLMClient,
    req: ReturnType<ContextBuilder['build']>,
    signal: AbortSignal
  ): Promise<StreamResult> {
    // Append a streaming assistant message — update its text live as deltas arrive
    const streamMsg = msg<AssistantMessage>(session.id, { role: 'assistant', text: '' });
    this.append(session.id, streamMsg);

    let text = '';
    let finishReason = 'stop';
    let inputTokens = 0;
    let outputTokens = 0;
    const accum: Record<number, { id: string; name: string; argsBuf: string }> = {};

    useStore.getState().updateSessionById(session.id, { status: 'parsing' });

    for await (const ev of client.chat(req, signal)) {
      if (signal.aborted) break;
      switch (ev.type) {
        case 'text-delta':
          text += ev.delta;
          useStore.getState().updateMessage(streamMsg.id, { text } as Partial<Message>);
          break;
        case 'tool-call-start':
          accum[ev.index] = { id: ev.id, name: ev.name, argsBuf: '' };
          break;
        case 'tool-call-delta':
          if (accum[ev.index]) accum[ev.index].argsBuf += ev.argumentsDelta;
          break;
        case 'usage':
          inputTokens = ev.inputTokens;
          outputTokens = ev.outputTokens;
          break;
        case 'done':
          finishReason = ev.finishReason;
          break;
        case 'error':
          throw new Error(`${ev.error.code}: ${ev.error.message}`);
      }
    }

    const toolCalls = Object.values(accum).map(a => {
      let args: Record<string, unknown> = {};
      try { args = JSON.parse(a.argsBuf) as Record<string, unknown>; } catch { /* leave empty */ }
      return { id: a.id, name: a.name, args, rawArgs: a.argsBuf };
    });

    // Append tool-call UI messages
    for (const tc of toolCalls) {
      this.append(session.id, msg<ToolCallMessage>(session.id, {
        role: 'tool_call',
        toolCall: { id: tc.id, name: tc.name, arguments: tc.args, workbookId: session.scope.workbookId, mutating: false },
        status: 'pending',
      }));
    }

    return { streamMsgId: streamMsg.id, text, toolCalls, finishReason, inputTokens, outputTokens };
  }

  // ── Tool execution ─────────────────────────────────────────────────────

  private async executeCalls(
    calls: ToolCall[],
    session: AgentSession,
    signal: AbortSignal,
    parallelizable: Set<string>,
    allowedTools: Set<string>
  ): Promise<void> {
    const scope = { workbookId: session.scope.workbookId };
    let i = 0;
    while (i < calls.length) {
      if (signal.aborted) return;
      let end = i;
      while (end < calls.length && parallelizable.has(calls[end].name)) end++;
      if (end - i >= 2) {
        // Results are appended in call order so the transcript stays deterministic.
        // Batch members are drawn from parallelizable ⊆ toolSpecs, so they are
        // always allowed for this run.
        const batch = calls.slice(i, end);
        const results = await Promise.all(batch.map(c => this.executor.execute(c, scope)));
        if (signal.aborted) return;
        for (let k = 0; k < batch.length; k++) {
          this.append(session.id, msg<ToolResultMessage>(session.id, {
            role: 'tool', toolCallId: batch[k].id, result: results[k],
          }));
        }
        i = end;
      } else {
        await this.executeCall(calls[i], session, signal, allowedTools);
        i++;
      }
    }
  }

  private async executeCall(
    call: ToolCall,
    session: AgentSession,
    signal: AbortSignal,
    allowedTools: Set<string>
  ): Promise<void> {
    const scope = { workbookId: session.scope.workbookId };

    if (session.provider === 'kimi' && session.webSearchEnabled && call.name === '$web_search') {
      const result = {
        toolCallId: call.id,
        ok: true as const,
        data: call.rawArguments ?? JSON.stringify(call.arguments),
      };
      this.append(session.id, msg<ToolResultMessage>(session.id, { role: 'tool', toolCallId: call.id, result }));
      return;
    }

    if (!allowedTools.has(call.name)) {
      const isWebTool = call.name === 'web_search' || call.name === 'fetch_url';
      const result = {
        toolCallId: call.id,
        ok: false as const,
        error: {
          code: 'ValidationError' as const,
          message: isWebTool
            ? `"${call.name}" is not available in this session: web access is not configured or web search is turned off. Do not call web tools again. Complete the task with workbook data, or tell the user they can enable a search provider in Settings → Web Access.`
            : `Tool "${call.name}" is not available in this session. Use only the tools in your tool list.`,
        },
      };
      this.append(session.id, msg<ToolResultMessage>(session.id, { role: 'tool', toolCallId: call.id, result }));
      return;
    }

    if (call.name === REQUEST_USER_CHOICE.name) {
      let pendingChoice: ReturnType<typeof parsePendingChoice>;
      try {
        pendingChoice = parsePendingChoice(call.id, call.arguments);
      } catch (e) {
        const message = e instanceof Error ? e.message : String(e);
        const result = {
          toolCallId: call.id,
          ok: false as const,
          error: {
            code: 'ValidationError' as const,
            message: e instanceof ToolValidationError ? message : `Invalid choice request: ${message}`,
          },
        };
        this.append(session.id, msg<ToolResultMessage>(session.id, { role: 'tool', toolCallId: call.id, result }));
        return;
      }

      useStore.getState().updateSessionById(session.id, { status: 'awaiting_choice', pendingChoice });
      const decision = await this.waitForChoice(signal);
      useStore.getState().updateSessionById(session.id, { status: 'executing_tool', pendingChoice: undefined });

      if (decision.kind === 'dismiss') {
        const result = {
          toolCallId: call.id,
          ok: false as const,
          error: { code: 'PermissionDenied' as const, message: 'User dismissed the choice menu' },
        };
        this.append(session.id, msg<ToolResultMessage>(session.id, { role: 'tool', toolCallId: call.id, result }));
        return;
      }

      const selected = pendingChoice.options.filter(option => decision.selection.ids.includes(option.id));
      const otherText = decision.selection.otherText?.trim();
      const result = {
        toolCallId: call.id,
        ok: true as const,
        data: {
          selected_ids: selected.map(option => option.id),
          selected_options: selected,
          ...(otherText ? { other_text: otherText } : {}),
        },
      };
      this.append(session.id, msg<ToolResultMessage>(session.id, { role: 'tool', toolCallId: call.id, result }));
      return;
    }

    if (!call.mutating) {
      const result = await this.executor.execute(call, scope);
      this.append(session.id, msg<ToolResultMessage>(session.id, { role: 'tool', toolCallId: call.id, result }));
      return;
    }

    // Mutating: capture → diff → confirm → apply/cancel
    const sheet = call.arguments.sheet as string | undefined;
    const address = call.arguments.address as string | undefined;
    const targetAddress = call.arguments.target_address as string | undefined;
    const snapshotAddress = address ?? targetAddress;

    let snapshotId: string | undefined;
    try {
      if (sheet && snapshotAddress) {
        const snap = await this.snapshots.captureRange(session.id, session.scope.workbookId, sheet, snapshotAddress, this.runner);
        snapshotId = snap.id;

        const proposed = call.arguments.values as unknown[][] | undefined;
        const diff = proposed ? computeRangeDiff(snapshotAddress, snap.before.values ?? [], proposed) : [];
        const wb = this.registry.getManifest().workbooks[0];

        useStore.getState().updateSessionById(session.id, {
          status: 'awaiting_confirmation',
          pendingChange: {
            id: ulid(), toolCall: call, snapshotId: snap.id, diff,
            severity: diff.length > 50 ? 'elevated' : 'normal',
            workbookName: wb?.name ?? 'Workbook', sheet,
          },
        });
      } else {
        useStore.getState().updateSessionById(session.id, { status: 'awaiting_confirmation' });
      }
    } catch (captureErr) {
      const message = captureErr instanceof Error ? captureErr.message : String(captureErr);
      const result = { toolCallId: call.id, ok: false as const, error: { code: 'OfficeApiError' as const, message: `Snapshot failed: ${message}` } };
      this.append(session.id, msg<ToolResultMessage>(session.id, { role: 'tool', toolCallId: call.id, result }));
      return;
    }

    const decision = useStore.getState().appConfig.autoApproveSession
      ? 'apply'
      : await this.waitForConfirmation(signal);
    useStore.getState().updateSessionById(session.id, { status: 'executing_tool', pendingChange: undefined });

    if (decision === 'cancel') {
      const result = { toolCallId: call.id, ok: false as const, error: { code: 'PermissionDenied' as const, message: 'User cancelled the write.' } };
      this.append(session.id, msg<ToolResultMessage>(session.id, { role: 'tool', toolCallId: call.id, result }));
      return;
    }

    const result = await this.executor.execute(call, scope);
    if (snapshotId) result.snapshotId = snapshotId;

    // Structural snapshot: after create_chart/create_pivot we know the object name, so capture
    // a coarse snapshot that undo can use to delete the created object.
    if (result.ok && !snapshotId) {
      const data = result.data as Record<string, unknown> | undefined;
      if ((call.name === 'create_chart' || call.name === 'create_pivot') && typeof data?.name === 'string') {
        const kind: 'chart' | 'pivot' = call.name === 'create_chart' ? 'chart' : 'pivot';
        const snap = this.snapshots.captureStructural(
          session.id, session.scope.workbookId,
          (call.arguments.sheet as string) ?? '',
          kind, data.name, { action: call.name }
        );
        result.snapshotId = snap.id;
      }
    }

    this.append(session.id, msg<ToolResultMessage>(session.id, { role: 'tool', toolCallId: call.id, result }));
  }

  // ── Helpers ────────────────────────────────────────────────────────────

  private waitForConfirmation(signal: AbortSignal): Promise<'apply' | 'cancel'> {
    return new Promise((resolve, reject) => {
      this.confirmationResolve = resolve;
      signal.addEventListener('abort', () => {
        this.confirmationResolve = null;
        reject(new DOMException('Aborted', 'AbortError'));
      }, { once: true });
    });
  }

  private waitForChoice(signal: AbortSignal): Promise<{ kind: 'select'; selection: ChoiceSelection } | { kind: 'dismiss' }> {
    return new Promise((resolve, reject) => {
      this.choiceResolve = resolve;
      signal.addEventListener('abort', () => {
        this.choiceResolve = null;
        reject(new DOMException('Aborted', 'AbortError'));
      }, { once: true });
    });
  }

  private append(sessionId: string, message: Message): void {
    void sessionId;
    useStore.getState().appendMessage(message);
  }
}

// ── Message factory ────────────────────────────────────────────────────────

function msg<T extends Message>(sessionId: string, fields: Omit<T, 'id' | 'sessionId' | 'createdAt'>): T {
  return { id: ulid(), sessionId, createdAt: new Date().toISOString(), ...fields } as unknown as T;
}

function shouldForceClientWebSearch(
  webProvider: WebAccessProvider,
  providerReady: boolean
): boolean {
  return providerReady && isKeylessSearchProvider(webProvider);
}

function readerFallbackEnabled(): boolean {
  return useStore.getState().appConfig.webAccess.readerFallback;
}
