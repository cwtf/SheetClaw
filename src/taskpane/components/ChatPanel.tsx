import { useState, useRef, useEffect } from 'react';
import type { CSSProperties, ReactNode } from 'react';
import {
  Body1,
  Body1Strong,
  Button,
  Caption1,
  MessageBar,
  MessageBarActions,
  MessageBarBody,
  Menu,
  MenuItem,
  MenuList,
  MenuPopover,
  MenuTrigger,
  Spinner,
  Textarea,
  tokens,
} from '@fluentui/react-components';
import { useStore } from '../../store/index';
import type { Message, CellDiff, WorkbookSelection } from '../../types';
import { createAdapter } from '../../adapters/index';
import { getUnavailableSearchToggleHint, resolveSearchToggle } from '../../adapters/native-search';
import type { ChoiceSelection } from '../../agent/loop';
import { getTaskpaneAgentLoop, getTaskpaneWorkbookLayer } from '../workbookLayer';
import { getCurrentWorkbookSelection } from '../selection';

const STATUS_RUNNING = new Set(['building', 'calling_llm', 'parsing', 'executing_tool']);
type ToolChainMessage = Extract<Message, { role: 'tool_call' | 'tool' }>;
type ThoughtProcessMessage = Extract<Message, { role: 'assistant' }>;
type ToolChainCall = {
  id: string;
  name: string;
  arguments: Record<string, unknown>;
  status?: Extract<Message, { role: 'tool_call' }>['status'];
  result?: Extract<Message, { role: 'tool' }>['result'];
};
type ChatRenderItem =
  | { type: 'message'; message: Message }
  | { type: 'thought_process'; id: string; messages: ThoughtProcessMessage[] }
  | { type: 'tool_chain'; id: string; calls: ToolChainCall[] };

const STATUS_LABELS: Record<string, string> = {
  building: 'Preparing context',
  calling_llm: 'Calling model',
  parsing: 'Reading response',
  executing_tool: 'Running workbook tool',
  awaiting_confirmation: 'Awaiting confirmation',
  awaiting_choice: 'Awaiting selection',
};

const EXAMPLE_PROMPTS = [
  'Summarize the active sheet',
  'Sum B2:B13 into B14',
  'Make a bar chart from A1:B12',
];

const composerActionStyle = {
  width: 36,
  minWidth: 36,
  height: 32,
  padding: 0,
};

function composerPillStyle(active: boolean, unavailable = false): CSSProperties {
  return {
    width: 36,
    height: 32,
    minWidth: 36,
    borderRadius: 999,
    padding: 0,
    border: `1px solid ${active ? '#4f7fe8' : tokens.colorNeutralStroke1}`,
    background: active ? '#162033' : tokens.colorNeutralBackground1,
    color: active ? '#6ea2ff' : tokens.colorNeutralForeground2,
    fontWeight: 600,
    opacity: unavailable ? 0.55 : 1,
    cursor: unavailable ? 'not-allowed' : 'pointer',
  };
}

function PillIcon({ children }: { children: ReactNode }) {
  return (
    <span
      aria-hidden="true"
      style={{
        display: 'inline-flex',
        alignItems: 'center',
        justifyContent: 'center',
        width: 16,
        height: 16,
        lineHeight: 1,
        fontSize: 13,
      }}
    >
      {children}
    </span>
  );
}

function NewChatIcon() {
  return (
    <svg aria-hidden="true" width="17" height="17" viewBox="0 0 20 20" fill="none" stroke="currentColor" strokeWidth="1.7" strokeLinecap="round" strokeLinejoin="round">
      <path d="M10.5 3.75H7A3.25 3.25 0 0 0 3.75 7v6A3.25 3.25 0 0 0 7 16.25h6A3.25 3.25 0 0 0 16.25 13V9.5" />
      <path d="m8.25 11.75.25-2 6.65-6.65a1.2 1.2 0 0 1 1.7 1.7L10.2 11.45l-1.95.3Z" />
    </svg>
  );
}

function AskBeforeEditsIcon() {
  return <PillIcon>🛇</PillIcon>;
}

function AcceptAllEditsIcon() {
  return <PillIcon>✔</PillIcon>;
}

function SelectionGridIcon() {
  return (
    <svg aria-hidden="true" width="12" height="12" viewBox="0 0 12 12" fill="currentColor">
      <rect x="1" y="1" width="4" height="4" rx="0.75" />
      <rect x="7" y="1" width="4" height="4" rx="0.75" />
      <rect x="1" y="7" width="4" height="4" rx="0.75" />
      <rect x="7" y="7" width="4" height="4" rx="0.75" />
    </svg>
  );
}

function SelectionBadge({ selection }: { selection: WorkbookSelection }) {
  const label = `${selection.sheet} ${selection.address} selected`;
  return (
    <div
      aria-label={label}
      title={label}
      style={{
        alignSelf: 'flex-start',
        display: 'inline-flex',
        alignItems: 'center',
        gap: 4,
        maxWidth: '100%',
        minHeight: 24,
        padding: '2px 7px',
        boxSizing: 'border-box',
        border: `1px solid ${tokens.colorNeutralStroke2}`,
        borderRadius: 5,
        background: tokens.colorNeutralBackground3,
        boxShadow: tokens.shadow2,
        color: tokens.colorNeutralForeground3,
        fontSize: 12,
        lineHeight: '18px',
      }}
    >
      <span style={{ display: 'inline-flex', color: tokens.colorPaletteGreenForeground1, flexShrink: 0 }}>
        <SelectionGridIcon />
      </span>
      <span style={{
        color: tokens.colorNeutralForeground1,
        fontWeight: 600,
        overflow: 'hidden',
        textOverflow: 'ellipsis',
        whiteSpace: 'nowrap',
      }}>
        {selection.sheet}
      </span>
      <span style={{ whiteSpace: 'nowrap' }}>
        <span style={{ fontFamily: 'monospace' }}>{selection.address}</span> selected
      </span>
    </div>
  );
}

export default function ChatPanel({ onOpenSettings }: { onOpenSettings?: (target?: 'search') => void }) {
  const [input, setInput] = useState('');
  const [currentSelection, setCurrentSelection] = useState<WorkbookSelection | null>(null);
  const [initError, setInitError] = useState<string | null>(null);
  const [searchHint, setSearchHint] = useState<string | null>(null);
  const [approvalMenuOpen, setApprovalMenuOpen] = useState(false);
  const bottomRef = useRef<HTMLDivElement>(null);
  const textareaRef = useRef<HTMLTextAreaElement | HTMLSpanElement>(null);
  const [textareaHeight, setTextareaHeight] = useState(32);

  useEffect(() => {
    if (typeof Excel === 'undefined') return;

    let disposed = false;
    let refreshId = 0;
    const refreshSelection = async () => {
      const id = ++refreshId;
      try {
        const selection = await getCurrentWorkbookSelection();
        if (!disposed && id === refreshId) setCurrentSelection(selection);
      } catch {
        if (!disposed && id === refreshId) setCurrentSelection(null);
      }
    };
    const handleSelectionChanged = () => { void refreshSelection(); };

    void refreshSelection();
    const document = typeof Office === 'undefined' ? undefined : Office.context?.document;
    document?.addHandlerAsync(Office.EventType.DocumentSelectionChanged, handleSelectionChanged);

    return () => {
      disposed = true;
      refreshId++;
      document?.removeHandlerAsync(
        Office.EventType.DocumentSelectionChanged,
        { handler: handleSelectionChanged },
      );
    };
  }, []);

  useEffect(() => {
    const el = textareaRef.current?.tagName === 'TEXTAREA'
      ? (textareaRef.current as HTMLTextAreaElement)
      : (textareaRef.current as HTMLSpanElement)?.querySelector('textarea');
    if (el) {
      const prevMin = el.style.minHeight;
      el.style.height = '0px';
      el.style.minHeight = '0px';
      const scrollHeight = el.scrollHeight;

      // Add a small buffer for the wrapper's padding/border (typically ~10-12px)
      setTextareaHeight(Math.min(Math.max(scrollHeight + 12, 32), 200));

      el.style.height = '100%';
      el.style.minHeight = prevMin;
    }
  }, [input]);

  const session = useStore(s => s.currentSession);
  const messages = useStore(s => s.messages);
  const providers = useStore(s => s.providers);
  const appConfig = useStore(s => s.appConfig);
  const setAppConfig = useStore(s => s.setAppConfig);
  const setSession = useStore(s => s.setSession);
  const clearSessionTotals = useStore(s => s.clearSessionTotals);
  const webSearchEnabled = useStore(s => s.webSearchEnabled);
  const setWebSearchEnabled = useStore(s => s.setWebSearchEnabled);
  const authStates = useStore(s => s.authStates);
  const activeProviderReady = useStore(s => s.isProviderReady(s.appConfig.activeProvider));
  const byokSearchReady = useStore(s =>
    s.appConfig.webAccess.provider !== 'none' && s.isSearchProviderReady(s.appConfig.webAccess.provider)
  );

  const isRunning = session ? STATUS_RUNNING.has(session.status) : false;
  const awaitingConfirm = session?.status === 'awaiting_confirmation';
  const activeProvider = providers[appConfig.activeProvider];
  const searchToggle = resolveSearchToggle({
    provider: appConfig.activeProvider,
    model: activeProvider?.model ?? '',
    byokReady: byokSearchReady,
  });
  const modelReady = !!activeProvider?.model.trim();
  const providerReady = !!activeProvider?.enabled && activeProviderReady && modelReady;
  const providerWarning = !activeProvider?.enabled
    ? 'No provider enabled. Configure one in Settings.'
    : !activeProviderReady
      ? 'Active provider is not authenticated. Configure auth in Settings.'
      : !modelReady
        ? 'Select a model in Settings before chatting.'
        : '';
  // The pill strictly reflects keyed BYOK / native search; keyless catalogue
  // search is always on in the background and has no pill state.
  const keyedSearchEnabled = webSearchEnabled && searchToggle.available;
  useEffect(() => {
    getTaskpaneWorkbookLayer().registry.refresh().catch(e => {
      setInitError(e instanceof Error ? e.message : String(e));
    });
  }, []);

  useEffect(() => {
    bottomRef.current?.scrollIntoView({ behavior: 'smooth' });
  }, [messages.length]);

  useEffect(() => {
    if (webSearchEnabled && !searchToggle.available) {
      setWebSearchEnabled(false);
    }
  }, [searchToggle.available, setWebSearchEnabled, webSearchEnabled]);

  async function send() {
    if (!input.trim() || isRunning || !providerReady) return;
    const text = input.trim();

    const provider = appConfig.activeProvider;
    const cfg = providers[provider];
    const authState = authStates[provider];
    const client = createAdapter(cfg, authState);

    try {
      await getTaskpaneWorkbookLayer().registry.refresh();
      const selection = await getCurrentWorkbookSelection();
      setCurrentSelection(selection);
      setInitError(null);
      setInput('');
      const scope = { workbookId: getTaskpaneWorkbookLayer().registry.getActiveId() ?? 'host' };
      if (session) {
        await getTaskpaneAgentLoop().followUp(text, scope, client, cfg, selection);
      } else {
        await getTaskpaneAgentLoop().start(text, scope, client, cfg, selection);
      }
    } catch (e) {
      const loopRunning = getTaskpaneAgentLoop().isRunning();
      if (!loopRunning) {
        setInitError(e instanceof Error ? e.message : String(e));
      }
      // LLM/tool errors are captured inside the loop and written to store.
    }
  }

  function stop() { getTaskpaneAgentLoop().stop(); }
  function newChat() {
    getTaskpaneAgentLoop().stop();
    setSession(null);
    clearSessionTotals();
    setInput('');
    setSearchHint(null);
    window.setTimeout(() => {
      const ref = textareaRef.current;
      const el = ref?.tagName === 'TEXTAREA'
        ? (ref as HTMLTextAreaElement)
        : (ref as HTMLSpanElement | null)?.querySelector('textarea');
      el?.focus();
    }, 0);
  }
  function applyConfirm() { getTaskpaneAgentLoop().resolveConfirmation('apply'); }
  function cancelConfirm() { getTaskpaneAgentLoop().resolveConfirmation('cancel'); }
  function resolveChoice(selection: ChoiceSelection) { getTaskpaneAgentLoop().resolveChoice(selection); }
  function dismissChoice() { getTaskpaneAgentLoop().resolveChoice('dismiss'); }
  async function continueRun() {
    if (!session) return;
    const provider = session.provider as keyof typeof providers;
    const cfg = providers[provider] ?? providers[appConfig.activeProvider];
    const authState = authStates[provider as keyof typeof authStates] ?? authStates[appConfig.activeProvider];
    const client = createAdapter(cfg, authState);
    try {
      await getTaskpaneWorkbookLayer().registry.refresh();
      setInitError(null);
      await getTaskpaneAgentLoop().continueCurrent(client, cfg);
    } catch (e) {
      const loopRunning = getTaskpaneAgentLoop().isRunning();
      if (!loopRunning) {
        setInitError(e instanceof Error ? e.message : String(e));
      }
      // Errors are captured inside loop.continueCurrent and written to store.
    }
  }

  async function undo() {
    if (!session) return;
    const snap = getTaskpaneWorkbookLayer().snapshots.lastUndoable(session.id);
    if (!snap) return;
    try {
      await getTaskpaneWorkbookLayer().snapshots.undo(snap.id, fn => Excel.run(fn));
    } catch (e) {
      setInitError(e instanceof Error ? e.message : String(e));
    }
  }

  const visibleMessages = messages.filter(m => (m as Message & { sessionId?: string }).sessionId === session?.id);
  const chatItems = buildChatRenderItems(visibleMessages);
  const awaitingChoice = session?.status === 'awaiting_choice';
  const canContinue = session?.status === 'done' && session.stopReason === 'max_iterations';
  const showEmptyState = visibleMessages.length === 0 && !isRunning && !awaitingConfirm && !awaitingChoice;

  return (
    <div style={{
      display: 'flex',
      flexDirection: 'column',
      height: '100%',
      minHeight: 0,
      padding: 12,
      gap: 8,
      boxSizing: 'border-box',
      overflowX: 'hidden',
    }}>
      {initError && (
        <MessageBar intent="error">
          <MessageBarBody>{initError}</MessageBarBody>
        </MessageBar>
      )}

      {!providerReady && (
        <MessageBar intent="warning">
          <MessageBarBody>{providerWarning}</MessageBarBody>
          {onOpenSettings && (
            <MessageBarActions>
              <Button size="small" appearance="subtle" onClick={() => onOpenSettings()}>Settings</Button>
            </MessageBarActions>
          )}
        </MessageBar>
      )}

      {searchHint && (
        <MessageBar intent="warning">
          <MessageBarBody>{searchHint}</MessageBarBody>
          {onOpenSettings && (
            <MessageBarActions>
              <Button size="small" appearance="subtle" onClick={() => onOpenSettings('search')}>Open Settings</Button>
            </MessageBarActions>
          )}
        </MessageBar>
      )}

      <div style={{
        flex: '1 1 auto',
        minHeight: 0,
        overflowY: 'auto',
        overflowX: 'hidden',
        display: 'flex',
        flexDirection: 'column',
        gap: 8,
      }}>
        {showEmptyState && (
          <EmptyChatState
            providerReady={providerReady}
            onPickPrompt={setInput}
            onOpenSettings={onOpenSettings}
          />
        )}
        {chatItems.map(item => {
          if (item.type === 'message') {
            return <MessageBubble key={item.message.id} message={item.message} />;
          }
          if (item.type === 'thought_process') {
            return <ThoughtProcess key={item.id} messages={item.messages} />;
          }
          return <ToolCallChain key={item.id} calls={item.calls} />;
        })}
        {awaitingConfirm && session?.pendingChange && (
          <ConfirmationBlock
            diff={session.pendingChange.diff}
            sheet={session.pendingChange.sheet}
            workbookName={session.pendingChange.workbookName}
            severity={session.pendingChange.severity}
            onApply={applyConfirm}
            onCancel={cancelConfirm}
          />
        )}
        {awaitingChoice && session?.pendingChoice && (
          <ChoiceBlock
            choice={session.pendingChoice}
            onContinue={resolveChoice}
            onDismiss={dismissChoice}
          />
        )}
        {(isRunning || awaitingChoice) && (
          <div style={{ display: 'flex', alignItems: 'center', gap: 6 }}>
            <Spinner size="extra-small" />
            <Caption1 style={{ color: tokens.colorNeutralForeground3 }}>
              {STATUS_LABELS[session?.status ?? ''] ?? 'Running'}...
            </Caption1>
          </div>
        )}
        <div ref={bottomRef} />
      </div>

      {session && (
        <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', flexShrink: 0, gap: 8 }}>
          <Caption1 style={{ color: tokens.colorNeutralForeground3 }}>
            {session.model} | iter {session.iteration}/{session.maxIterations} | {session.totals.inputTokens + session.totals.outputTokens} tok
          </Caption1>
          <div style={{ display: 'flex', gap: 6, flexShrink: 0 }}>
            {canContinue && (
              <Button size="small" appearance="primary" onClick={() => void continueRun()}>
                Continue
              </Button>
            )}
            <Button size="small" appearance="subtle" onClick={() => void undo()}>Undo last write</Button>
          </div>
        </div>
      )}

      <div style={{
        display: 'flex',
        flexDirection: 'column',
        gap: 4,
        flexShrink: 0,
      }}>
        {currentSelection && <SelectionBadge selection={currentSelection} />}
        <Textarea
          ref={textareaRef as any}
          style={{ width: '100%', minHeight: 32, height: textareaHeight }}
          placeholder="Ask me anything..."
          rows={1}
          value={input}
          onChange={(_, d) => setInput(d.value)}
          onKeyDown={e => {
            if (e.key === 'Enter' && !e.shiftKey) {
              e.preventDefault();
              void send();
            }
          }}
          disabled={isRunning || awaitingConfirm || awaitingChoice || !providerReady}
          resize="none"
        />
        <div style={{
          display: 'flex',
          alignItems: 'center',
          justifyContent: 'space-between',
          gap: 8,
        }}>
          <div style={{
            display: 'flex',
            alignItems: 'center',
            gap: 6,
            minWidth: 0,
          }}>
            <Button
              size="small"
              appearance="secondary"
              style={composerActionStyle}
              onClick={newChat}
              aria-label="New chat"
              title="New chat"
            >
              <NewChatIcon />
            </Button>
            <Button
              size="small"
              appearance="secondary"
              aria-pressed={keyedSearchEnabled}
              aria-disabled={!searchToggle.available}
              aria-label="Search"
              title="Search"
              style={composerPillStyle(keyedSearchEnabled, !searchToggle.available)}
              icon={<PillIcon>🌐</PillIcon>}
              onClick={() => {
                if (!searchToggle.available) {
                  setWebSearchEnabled(false);
                  setSearchHint(getUnavailableSearchToggleHint(activeProvider?.label ?? appConfig.activeProvider, searchToggle));
                  return;
                }
                setSearchHint(null);
                setWebSearchEnabled(!webSearchEnabled);
              }}
            />
            <Menu
              open={approvalMenuOpen}
              onOpenChange={(_, data) => setApprovalMenuOpen(data.open)}
            >
              <MenuTrigger disableButtonEnhancement>
                <Button
                  size="small"
                  appearance="secondary"
                  aria-label="Edit approval mode"
                  title={appConfig.autoApproveSession ? 'Accept all edits' : 'Ask before edits'}
                  style={composerPillStyle(approvalMenuOpen)}
                  icon={appConfig.autoApproveSession ? <AcceptAllEditsIcon /> : <AskBeforeEditsIcon />}
                />
              </MenuTrigger>
              <MenuPopover style={{ minWidth: 220 }}>
                <MenuList>
                  <MenuItem
                    role="menuitemradio"
                    aria-checked={!appConfig.autoApproveSession}
                    icon={<AskBeforeEditsIcon />}
                    secondaryContent={!appConfig.autoApproveSession ? <span style={{ color: tokens.colorBrandForeground1 }}>✓</span> : undefined}
                    onClick={() => setAppConfig({ autoApproveSession: false })}
                  >
                    Ask before edits
                  </MenuItem>
                  <MenuItem
                    role="menuitemradio"
                    aria-checked={appConfig.autoApproveSession}
                    icon={<AcceptAllEditsIcon />}
                    secondaryContent={appConfig.autoApproveSession ? <span style={{ color: tokens.colorBrandForeground1 }}>✓</span> : undefined}
                    onClick={() => setAppConfig({ autoApproveSession: true })}
                  >
                    Accept all edits
                  </MenuItem>
                </MenuList>
              </MenuPopover>
            </Menu>
          </div>
          {isRunning
            ? (
              <Button
                appearance="secondary"
                onClick={stop}
                style={{ ...composerActionStyle, width: 56, minWidth: 56 }}
                aria-label="Stop"
                title="Stop"
              >
                Stop
              </Button>
            )
            : (
              <Button
                appearance="primary"
                onClick={() => void send()}
                disabled={!input.trim() || !providerReady}
                style={composerActionStyle}
                aria-label="Send"
                title="Send"
                icon={<SendIcon />}
              />
            )
          }
        </div>
      </div>
    </div>
  );
}

function isToolChainMessage(message: Message): message is ToolChainMessage {
  return message.role === 'tool_call' || message.role === 'tool';
}

export function buildChatRenderItems(messages: Message[]): ChatRenderItem[] {
  const items: ChatRenderItem[] = [];
  let segment: Message[] = [];

  function flushSegment() {
    if (!segment.length) return;
    items.push(...buildSegmentRenderItems(segment));
    segment = [];
  }

  for (const message of messages) {
    if (message.role === 'user') {
      flushSegment();
      items.push({ type: 'message', message });
    } else {
      segment.push(message);
    }
  }
  flushSegment();

  return items;
}

function buildSegmentRenderItems(segment: Message[]): ChatRenderItem[] {
  const calls = collectToolCalls(segment);
  if (!calls.length) {
    return segment.map(message => ({ type: 'message', message }));
  }

  const items: ChatRenderItem[] = [];
  const thoughts = segment.filter(isThoughtProcessAssistant);
  const toolCallIds = new Set(calls.map(call => call.id));
  let inserted = false;
  const preferredInsertIndex = findToolChainInsertIndex(segment, toolCallIds);

  segment.forEach((message, index) => {
    if (!isToolChainMessage(message) && !isToolCallingAssistant(message)) {
      items.push({ type: 'message', message });
    }
    if (!inserted && index === preferredInsertIndex) {
      if (thoughts.length) {
        items.push({
          type: 'thought_process',
          id: `thought-process-${thoughts[0].id}`,
          messages: thoughts,
        });
      }
      items.push({ type: 'tool_chain', id: `tool-chain-${calls[0].id}`, calls });
      inserted = true;
    }
  });

  if (!inserted) {
    items.push({ type: 'tool_chain', id: `tool-chain-${calls[0].id}`, calls });
  }
  return items;
}

function collectToolCalls(messages: Message[]): ToolChainCall[] {
  const calls = new Map<string, ToolChainCall>();

  function ensureCall(id: string, name = 'tool', args: Record<string, unknown> = {}) {
    const existing = calls.get(id);
    if (existing) {
      if (existing.name === 'tool' && name !== 'tool') existing.name = name;
      if (!Object.keys(existing.arguments).length && Object.keys(args).length) existing.arguments = args;
      return existing;
    }
    const call: ToolChainCall = { id, name, arguments: args };
    calls.set(id, call);
    return call;
  }

  for (const message of messages) {
    if (message.role === 'assistant') {
      for (const call of message.toolCalls ?? []) {
        ensureCall(call.id, call.name, call.arguments);
      }
    } else if (message.role === 'tool_call') {
      const call = ensureCall(message.toolCall.id, message.toolCall.name, message.toolCall.arguments);
      call.status = message.status;
    } else if (message.role === 'tool') {
      const call = ensureCall(message.toolCallId);
      call.result = message.result;
    }
  }

  return [...calls.values()];
}

function findToolChainInsertIndex(segment: Message[], toolCallIds: Set<string>): number {
  for (let i = segment.length - 1; i >= 0; i--) {
    const message = segment[i];
    if (message.role === 'assistant' && message.toolCalls?.some(call => toolCallIds.has(call.id))) return i;
  }
  const firstToolIndex = segment.findIndex(isToolChainMessage);
  return firstToolIndex > 0 ? firstToolIndex - 1 : 0;
}

function isToolCallingAssistant(message: Message): message is ThoughtProcessMessage {
  return message.role === 'assistant' && !!message.toolCalls?.length;
}

function isThoughtProcessAssistant(message: Message): message is ThoughtProcessMessage {
  return isToolCallingAssistant(message) && message.text.trim().length > 0;
}

function SendIcon() {
  return (
    <span
      aria-hidden="true"
      style={{
        width: 0,
        height: 0,
        borderTop: '5px solid transparent',
        borderBottom: '5px solid transparent',
        borderLeft: '9px solid currentColor',
        transform: 'translateX(1px)',
      }}
    />
  );
}

function ThoughtProcess({ messages }: { messages: ThoughtProcessMessage[] }) {
  const [expanded, setExpanded] = useState(false);

  return (
    <div style={{
      border: `1px solid ${tokens.colorNeutralStroke1}`,
      borderRadius: 6,
      background: tokens.colorNeutralBackground1,
      overflow: 'hidden',
      flexShrink: 0,
      boxShadow: expanded ? `0 0 0 1px ${tokens.colorBrandStroke1}` : undefined,
    }}>
      <button
        type={'button'}
        aria-expanded={expanded}
        onClick={() => setExpanded(value => !value)}
        style={{
          width: '100%',
          border: 0,
          background: 'transparent',
          color: tokens.colorNeutralForeground2,
          display: 'flex',
          alignItems: 'center',
          justifyContent: 'space-between',
          gap: 8,
          padding: '7px 9px',
          cursor: 'pointer',
          textAlign: 'left',
        }}
      >
        <span style={{
          display: 'inline-flex',
          alignItems: 'center',
          gap: 8,
          minWidth: 0,
          fontSize: 12,
          fontWeight: 600,
        }}>
          <span aria-hidden style={{ color: tokens.colorBrandForeground1, width: 10 }}>
            {expanded ? '-' : '+'}
          </span>
          <span>Thought process ({messages.length})</span>
        </span>
        <Caption1 style={{ color: tokens.colorNeutralForeground3, flexShrink: 0 }}>
          {messages.length} {messages.length === 1 ? 'step' : 'steps'}
        </Caption1>
      </button>
      {expanded && (
        <div style={{
          borderTop: `1px solid ${tokens.colorNeutralStroke1}`,
          display: 'flex',
          flexDirection: 'column',
          gap: 6,
          padding: 8,
        }}>
          {messages.map(message => (
            <div
              key={message.id}
              style={{
                borderRadius: 6,
                background: tokens.colorNeutralBackground2,
                color: tokens.colorNeutralForeground1,
                padding: '7px 9px',
                minWidth: 0,
                overflowWrap: 'anywhere',
              }}
            >
              <MarkdownMessage text={message.text} />
            </div>
          ))}
        </div>
      )}
    </div>
  );
}

function ToolCallChain({ calls }: { calls: ToolChainCall[] }) {
  const [expanded, setExpanded] = useState(false);
  const complete = calls.filter(call => call.result || call.status === 'applied' || call.status === 'failed').length;
  const failed = calls.filter(call => call.result ? !call.result.ok : call.status === 'failed').length;

  return (
    <div style={{
      border: `1px solid ${tokens.colorNeutralStroke1}`,
      borderRadius: 6,
      background: tokens.colorNeutralBackground1,
      overflow: 'hidden',
      flexShrink: 0,
      boxShadow: expanded ? `0 0 0 1px ${tokens.colorBrandStroke1}` : undefined,
    }}>
      <button
        type="button"
        aria-expanded={expanded}
        onClick={() => setExpanded(value => !value)}
        style={{
          width: '100%',
          border: 0,
          background: 'transparent',
          color: tokens.colorNeutralForeground2,
          display: 'flex',
          alignItems: 'center',
          justifyContent: 'space-between',
          gap: 8,
          padding: '7px 9px',
          cursor: 'pointer',
          textAlign: 'left',
        }}
      >
        <span style={{
          display: 'inline-flex',
          alignItems: 'center',
          gap: 8,
          minWidth: 0,
          fontSize: 12,
          fontWeight: 600,
        }}>
          <span aria-hidden="true" style={{ color: tokens.colorBrandForeground1, width: 10 }}>
            {expanded ? '-' : '+'}
          </span>
          <span>Tool calls ({calls.length})</span>
        </span>
        <Caption1 style={{
          color: failed ? tokens.colorPaletteRedForeground1 : tokens.colorNeutralForeground3,
          flexShrink: 0,
          fontFamily: 'monospace',
        }}>
          {complete}/{calls.length} complete{failed ? `, ${failed} failed` : ''}
        </Caption1>
      </button>
      {expanded && (
        <div style={{
          borderTop: `1px solid ${tokens.colorNeutralStroke1}`,
          display: 'flex',
          flexDirection: 'column',
          gap: 6,
          padding: 8,
        }}>
          {calls.map(call => (
            <ToolChainRow key={call.id} call={call} />
          ))}
        </div>
      )}
    </div>
  );
}

function ToolChainRow({ call }: { call: ToolChainCall }) {
  const state = getToolCallState(call);
  return (
    <div style={{
      border: `1px solid ${tokens.colorNeutralStroke1}`,
      borderLeft: `3px solid ${state.accent}`,
      borderRadius: 6,
      background: tokens.colorNeutralBackground2,
      color: state.color,
      fontFamily: 'monospace',
      fontSize: 11,
      lineHeight: 1.35,
      padding: '6px 8px',
      minWidth: 0,
      overflowWrap: 'anywhere',
    }}>
      <strong>{state.label}:</strong> {call.name}{state.outcome ? ` -> ${state.outcome}` : ''}
    </div>
  );
}

function getToolCallState(call: ToolChainCall) {
  if (!call.result && call.status !== 'applied' && call.status !== 'failed') {
    return {
      label: 'running',
      outcome: '',
      color: '#25b7d3',
      accent: '#25b7d3',
    };
  }
  if (call.result && !call.result.ok || call.status === 'failed') {
    const message = call.result?.error?.message?.trim();
    return {
      label: 'error',
      outcome: message ? `failed — ${truncateErrorMessage(message)}` : 'failed',
      color: tokens.colorPaletteRedForeground1,
      accent: tokens.colorPaletteRedForeground1,
    };
  }
  return {
    label: 'ok',
    outcome: formatToolOutcome(call.result?.data),
    color: tokens.colorPaletteGreenForeground1,
    accent: tokens.colorPaletteGreenForeground1,
  };
}

function truncateErrorMessage(message: string): string {
  const MAX = 180;
  return message.length <= MAX ? message : `${message.slice(0, MAX)}…`;
}

function formatToolOutcome(data: unknown): string {
  if (data && typeof data === 'object') {
    const status = (data as { status?: unknown }).status;
    if (typeof status === 'string' && status.trim()) return status;
  }
  return 'ok';
}

function EmptyChatState({
  providerReady,
  onPickPrompt,
  onOpenSettings,
}: {
  providerReady: boolean;
  onPickPrompt: (prompt: string) => void;
  onOpenSettings?: (target?: 'search') => void;
}) {
  return (
    <div style={{
      margin: 'auto 0',
      display: 'flex',
      flexDirection: 'column',
      gap: 10,
      alignItems: 'stretch',
    }}>
      <div>
        <Body1Strong>{providerReady ? 'Ready for this workbook' : 'Set up a provider to start'}</Body1Strong>
        <Caption1 style={{
          display: 'block',
          marginTop: 2,
          color: tokens.colorNeutralForeground3,
        }}>
          {providerReady
            ? 'Pick a starter prompt or ask your own question.'
            : 'Choose a provider, model, and authentication method in Settings.'}
        </Caption1>
      </div>
      {providerReady ? (
        <div style={{ display: 'flex', flexDirection: 'column', gap: 6 }}>
          {EXAMPLE_PROMPTS.map(prompt => (
            <Button
              key={prompt}
              size="small"
              appearance="secondary"
              style={{ justifyContent: 'flex-start' }}
              onClick={() => onPickPrompt(prompt)}
            >
              {prompt}
            </Button>
          ))}
        </div>
      ) : onOpenSettings ? (
        <Button appearance="primary" size="small" onClick={() => onOpenSettings()}>
          Open Settings
        </Button>
      ) : null}
    </div>
  );
}

function MessageBubble({ message }: { message: Message }) {
  const isUser = message.role === 'user';
  const isSystem = message.role === 'system_notice';
  const text = message.role === 'assistant' || message.role === 'user' ? message.text : '';

  if (message.role === 'tool_call') {
    return (
      <Caption1 style={{
        display: 'block',
        minWidth: 0,
        maxWidth: '100%',
        color: tokens.colorNeutralForeground3,
        fontFamily: 'monospace',
        whiteSpace: 'pre-wrap',
        overflowWrap: 'anywhere',
      }}>
        Tool: {message.toolCall.name}({JSON.stringify(message.toolCall.arguments).slice(0, 80)})
      </Caption1>
    );
  }
  if (message.role === 'tool') {
    const ok = message.result.ok;
    return (
      <Caption1 style={{
        display: 'block',
        minWidth: 0,
        maxWidth: '100%',
        color: ok ? tokens.colorPaletteGreenForeground1 : tokens.colorPaletteRedForeground1,
        fontFamily: 'monospace',
        whiteSpace: 'pre-wrap',
        overflowWrap: 'anywhere',
      }}>
        {ok ? 'OK' : 'ERR'} {message.toolCallId.slice(0, 12)}... {ok ? JSON.stringify(message.result.data).slice(0, 80) : message.result.error?.message}
      </Caption1>
    );
  }
  if (message.role === 'confirmation') return null;
  if (isSystem) {
    const intent = message.level === 'error' ? 'error' : message.level === 'warn' ? 'warning' : 'info';
    return (
      <MessageBar intent={intent}>
        <MessageBarBody><Caption1>{message.text}</Caption1></MessageBarBody>
      </MessageBar>
    );
  }

  return (
    <div style={{
      alignSelf: isUser ? 'flex-end' : 'flex-start',
      maxWidth: '85%',
      background: isUser ? tokens.colorBrandBackground : tokens.colorNeutralBackground2,
      color: isUser ? tokens.colorNeutralForegroundOnBrand : tokens.colorNeutralForeground1,
      borderRadius: 8,
      padding: '8px 12px',
      minWidth: 0,
      overflowWrap: 'anywhere',
      wordBreak: 'break-word',
    }}>
      {message.role === 'assistant'
        ? <MarkdownMessage text={text} />
        : <Body1 style={{ whiteSpace: 'pre-wrap' }}>{text}</Body1>}
    </div>
  );
}

type MarkdownPart =
  | { type: 'text'; text: string }
  | { type: 'table'; headers: string[]; aligns: Array<'left' | 'right' | 'center'>; rows: string[][] };

function MarkdownMessage({ text }: { text: string }) {
  const parts = parseMarkdownTables(text);
  if (parts.length === 1 && parts[0].type === 'text') {
    return <Body1 style={{ whiteSpace: 'pre-wrap' }}>{renderInlineMarkdown(text)}</Body1>;
  }

  return (
    <div style={{ display: 'flex', flexDirection: 'column', gap: 8 }}>
      {parts.map((part, index) => part.type === 'text' ? (
        part.text.trim() ? (
          <Body1 key={index} style={{ whiteSpace: 'pre-wrap' }}>
            {renderInlineMarkdown(part.text.trim())}
          </Body1>
        ) : null
      ) : (
        <MarkdownTable key={index} part={part} />
      ))}
    </div>
  );
}

function MarkdownTable({ part }: { part: Extract<MarkdownPart, { type: 'table' }> }) {
  return (
    <div style={{
      maxWidth: '100%',
      overflowX: 'auto',
      border: `1px solid ${tokens.colorNeutralStroke1}`,
      borderRadius: 6,
      background: tokens.colorNeutralBackground1,
    }}>
      <table style={{
        width: '100%',
        borderCollapse: 'collapse',
        fontSize: 12,
        lineHeight: 1.35,
        minWidth: Math.min(720, Math.max(360, part.headers.length * 120)),
      }}>
        <thead>
          <tr>
            {part.headers.map((header, index) => (
              <th
                key={index}
                style={{
                  textAlign: part.aligns[index] ?? 'left',
                  padding: '6px 8px',
                  borderBottom: `1px solid ${tokens.colorNeutralStroke1}`,
                  background: tokens.colorNeutralBackground2,
                  fontWeight: 600,
                  verticalAlign: 'top',
                }}
              >
                {renderInlineMarkdown(header)}
              </th>
            ))}
          </tr>
        </thead>
        <tbody>
          {part.rows.map((row, rowIndex) => (
            <tr key={rowIndex}>
              {part.headers.map((_, cellIndex) => (
                <td
                  key={cellIndex}
                  style={{
                    textAlign: part.aligns[cellIndex] ?? 'left',
                    padding: '6px 8px',
                    borderBottom: rowIndex === part.rows.length - 1 ? undefined : `1px solid ${tokens.colorNeutralStroke2}`,
                    verticalAlign: 'top',
                    overflowWrap: 'anywhere',
                  }}
                >
                  {renderInlineMarkdown(row[cellIndex] ?? '')}
                </td>
              ))}
            </tr>
          ))}
        </tbody>
      </table>
    </div>
  );
}

function parseMarkdownTables(text: string): MarkdownPart[] {
  const lines = text.split(/\r?\n/);
  const parts: MarkdownPart[] = [];
  const textBuffer: string[] = [];

  function flushText() {
    if (textBuffer.length) {
      parts.push({ type: 'text', text: textBuffer.join('\n') });
      textBuffer.length = 0;
    }
  }

  for (let i = 0; i < lines.length; i++) {
    const header = splitMarkdownRow(lines[i]);
    const separator = i + 1 < lines.length ? parseSeparator(lines[i + 1]) : null;
    if (header.length >= 2 && separator && separator.length === header.length) {
      flushText();
      const rows: string[][] = [];
      i += 2;
      while (i < lines.length) {
        const row = splitMarkdownRow(lines[i]);
        if (row.length === 0) break;
        rows.push(normalizeTableRow(row, header.length));
        i++;
      }
      i--;
      parts.push({ type: 'table', headers: header, aligns: separator, rows });
      continue;
    }
    textBuffer.push(lines[i]);
  }

  flushText();
  return parts.length ? parts : [{ type: 'text', text }];
}

function splitMarkdownRow(line: string): string[] {
  const trimmed = line.trim();
  if (!trimmed.includes('|')) return [];
  const withoutEdges = trimmed.replace(/^\|/, '').replace(/\|$/, '');
  const cells = withoutEdges.split('|').map(cell => cell.trim());
  return cells.some(Boolean) ? cells : [];
}

function parseSeparator(line: string): Array<'left' | 'right' | 'center'> | null {
  const cells = splitMarkdownRow(line);
  if (!cells.length) return null;
  const aligns: Array<'left' | 'right' | 'center'> = [];
  for (const cell of cells) {
    const marker = cell.replace(/\s/g, '');
    if (!/^:?-{3,}:?$/.test(marker)) return null;
    aligns.push(marker.startsWith(':') && marker.endsWith(':') ? 'center' : marker.endsWith(':') ? 'right' : 'left');
  }
  return aligns;
}

function normalizeTableRow(row: string[], width: number): string[] {
  const normalized = [...row];
  while (normalized.length > width && normalized[normalized.length - 1] === '') normalized.pop();
  return normalized.slice(0, width);
}

function renderInlineMarkdown(value: string): ReactNode[] {
  const nodes: ReactNode[] = [];
  const pattern = /\*\*([^*]+)\*\*/g;
  let lastIndex = 0;
  let match: RegExpExecArray | null;
  while ((match = pattern.exec(value))) {
    if (match.index > lastIndex) nodes.push(value.slice(lastIndex, match.index));
    nodes.push(<strong key={match.index}>{match[1]}</strong>);
    lastIndex = match.index + match[0].length;
  }
  if (lastIndex < value.length) nodes.push(value.slice(lastIndex));
  return nodes.length ? nodes : [value];
}

function ConfirmationBlock({
  diff, sheet, workbookName, severity, onApply, onCancel,
}: {
  diff: CellDiff[];
  sheet: string;
  workbookName: string;
  severity: 'normal' | 'elevated';
  onApply: () => void;
  onCancel: () => void;
}) {
  const MAX_SHOWN = 10;
  const shown = diff.slice(0, MAX_SHOWN);

  return (
    <div style={{
      border: `1px solid ${severity === 'elevated' ? tokens.colorPaletteRedBorder2 : tokens.colorNeutralStroke1}`,
      borderRadius: 6,
      padding: 12,
      display: 'flex',
      flexDirection: 'column',
      gap: 8,
      background: tokens.colorNeutralBackground1,
    }}>
      <Body1Strong>Confirm change - {workbookName} / {sheet}</Body1Strong>
      {severity === 'elevated' && (
        <Caption1 style={{ color: tokens.colorPaletteRedForeground1 }}>Large change - review carefully</Caption1>
      )}
      <div style={{ fontFamily: 'monospace', fontSize: 11, display: 'flex', flexDirection: 'column', gap: 2 }}>
        {shown.map((d, i) => (
          <div key={i}>
            <span style={{ color: tokens.colorNeutralForeground3 }}>{d.address}: </span>
            <span style={{ color: tokens.colorPaletteRedForeground1 }}>{fmt(d.before)}</span>
            <span> to </span>
            <span style={{ color: tokens.colorPaletteGreenForeground1 }}>{fmt(d.after)}</span>
          </div>
        ))}
        {diff.length > MAX_SHOWN && (
          <Caption1 style={{ color: tokens.colorNeutralForeground3 }}>...and {diff.length - MAX_SHOWN} more cells</Caption1>
        )}
        {diff.length === 0 && <Caption1 style={{ color: tokens.colorNeutralForeground3 }}>(no cell values change)</Caption1>}
      </div>
      <div style={{ display: 'flex', gap: 8 }}>
        <Button appearance="primary" onClick={onApply}>Apply</Button>
        <Button appearance="secondary" onClick={onCancel}>Cancel</Button>
      </div>
    </div>
  );
}

function ChoiceBlock({
  choice,
  onContinue,
  onDismiss,
}: {
  choice: NonNullable<import('../../types').AgentSession['pendingChoice']>;
  onContinue: (selection: ChoiceSelection) => void;
  onDismiss: () => void;
}) {
  const [selected, setSelected] = useState<string[]>([]);
  const [otherText, setOtherText] = useState('');
  const selectedSet = new Set(selected);
  const selectedOther = choice.options.find(option => option.requiresText && selectedSet.has(option.id));
  const canContinue = selected.length > 0 && (!selectedOther || otherText.trim().length > 0);

  function toggle(id: string) {
    if (choice.allowMultiple) {
      setSelected(prev => prev.includes(id) ? prev.filter(x => x !== id) : [...prev, id]);
    } else {
      setSelected(prev => prev[0] === id ? [] : [id]);
    }
  }

  return (
    <div style={{
      border: `1px solid ${tokens.colorNeutralStroke1}`,
      borderRadius: 6,
      padding: 12,
      display: 'flex',
      flexDirection: 'column',
      gap: 10,
      background: tokens.colorNeutralBackground1,
      minWidth: 0,
    }}>
      <Body1Strong>{choice.question}</Body1Strong>
      <div style={{ display: 'flex', flexDirection: 'column', gap: 6 }}>
        {choice.options.map((option, index) => {
          const active = selectedSet.has(option.id);
          return (
            <div key={option.id} style={{ display: 'flex', flexDirection: 'column', gap: 6 }}>
              <button
                type="button"
                onClick={() => toggle(option.id)}
                aria-pressed={active}
                style={{
                  textAlign: 'left',
                  border: `1px solid ${active ? tokens.colorBrandStroke1 : tokens.colorNeutralStroke1}`,
                  background: active ? tokens.colorBrandBackground2 : tokens.colorNeutralBackground2,
                  color: tokens.colorNeutralForeground1,
                  borderRadius: 6,
                  padding: 8,
                  cursor: 'pointer',
                  display: 'flex',
                  gap: 8,
                  minWidth: 0,
                }}
              >
                <Caption1 style={{ color: tokens.colorNeutralForeground3, width: 18, flexShrink: 0 }}>
                  {index + 1}.
                </Caption1>
                <span style={{ display: 'flex', flexDirection: 'column', gap: 2, minWidth: 0 }}>
                  <Body1Strong style={{ overflowWrap: 'anywhere' }}>{option.label}</Body1Strong>
                  {option.description && (
                    <Caption1 style={{ color: tokens.colorNeutralForeground3, overflowWrap: 'anywhere' }}>
                      {option.description}
                    </Caption1>
                  )}
                </span>
              </button>
              {option.requiresText && active && (
                <Textarea
                  value={otherText}
                  onChange={(_, d) => setOtherText(d.value)}
                  placeholder="Specify your requirements..."
                  resize="vertical"
                  rows={3}
                  style={{ width: '100%' }}
                />
              )}
            </div>
          );
        })}
      </div>
      <div style={{ display: 'flex', gap: 8 }}>
        <Button
          appearance="primary"
          disabled={!canContinue}
          onClick={() => onContinue({ ids: selected, otherText: otherText.trim() || undefined })}
        >
          Continue
        </Button>
        <Button appearance="secondary" onClick={onDismiss}>Dismiss</Button>
      </div>
    </div>
  );
}

function fmt(v: unknown): string {
  if (v === null || v === undefined || v === '') return '(empty)';
  return String(v).slice(0, 40);
}
