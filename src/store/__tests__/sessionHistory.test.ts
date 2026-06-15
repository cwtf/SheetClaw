import { beforeEach, describe, expect, it, vi } from 'vitest';
import { useStore } from '../index';
import type { AgentSession, Message } from '../../types';

function makeLocalStorageStub() {
  const map = new Map<string, string>();
  return {
    getItem: (k: string) => (map.has(k) ? map.get(k)! : null),
    setItem: (k: string, v: string) => {
      map.set(k, String(v));
    },
    removeItem: (k: string) => {
      map.delete(k);
    },
    clear: () => map.clear(),
    key: (i: number) => Array.from(map.keys())[i] ?? null,
    get length() {
      return map.size;
    },
  };
}

function session(id: string, status: AgentSession['status'] = 'done'): AgentSession {
  return {
    id,
    createdAt: '2026-06-12T01:00:00.000Z',
    scope: { workbookId: 'wb-1' },
    status,
    iteration: 1,
    maxIterations: 25,
    provider: 'openai',
    model: 'gpt-4o',
    messageIds: [],
    webSearchEnabled: false,
    tokenBudget: { used: 0, window: 128000 },
    totals: { inputTokens: 12, outputTokens: 34, costUsd: 0.001 },
  };
}

function userMessage(sessionId: string, text: string): Message {
  return {
    id: `${sessionId}-m1`,
    sessionId,
    createdAt: '2026-06-12T01:01:00.000Z',
    role: 'user',
    text,
  };
}

beforeEach(() => {
  vi.stubGlobal('localStorage', makeLocalStorageStub());
  useStore.setState({
    currentSession: null,
    messages: [],
    chatHistory: [],
    webSearchEnabled: false,
    sessionTotals: null,
  });
});

describe('session history', () => {
  it('persists chat summaries and restores messages for resume', () => {
    useStore.getState().setSession(session('s1'));
    useStore.getState().appendMessage(userMessage('s1', 'Summarize the active sheet'));
    useStore.getState().updateSession({ status: 'done' });

    expect(useStore.getState().chatHistory[0]).toMatchObject({
      id: 's1',
      title: 'Summarize the active sheet',
      messageCount: 1,
    });

    useStore.setState({ currentSession: null, messages: [], chatHistory: [] });
    useStore.getState().loadChatHistory();
    expect(useStore.getState().chatHistory).toHaveLength(1);

    expect(useStore.getState().resumeChat('s1')).toBe(true);
    expect(useStore.getState().currentSession?.id).toBe('s1');
    expect(useStore.getState().messages).toEqual([
      expect.objectContaining({ role: 'user', text: 'Summarize the active sheet' }),
    ]);
  });

  it('restores interrupted active sessions as stopped', () => {
    useStore.getState().setSession({
      ...session('s2', 'awaiting_choice'),
      pendingChoice: {
        id: 'choice-1',
        toolCallId: 'call-1',
        question: 'Pick one',
        options: [],
        allowMultiple: false,
      },
    });
    useStore.getState().appendMessage(userMessage('s2', 'Find revenue'));

    useStore.setState({ currentSession: null, messages: [] });
    expect(useStore.getState().resumeChat('s2')).toBe(true);

    const restored = useStore.getState().currentSession;
    expect(restored?.status).toBe('stopped');
    expect(restored?.pendingChoice).toBeUndefined();
    expect(restored?.lastError?.code).toBe('SessionInterrupted');
  });

  it('deletes an individual chat transcript', () => {
    useStore.getState().setSession(session('s1'));
    useStore.getState().appendMessage(userMessage('s1', 'First chat'));
    useStore.getState().setSession(session('s2'));
    useStore.getState().appendMessage(userMessage('s2', 'Second chat'));

    useStore.getState().deleteChat('s1');

    expect(useStore.getState().chatHistory.map(item => item.id)).toEqual(['s2']);
    expect(localStorage.getItem('xl.chat.history.s1')).toBeNull();
    expect(useStore.getState().resumeChat('s1')).toBe(false);
    expect(useStore.getState().resumeChat('s2')).toBe(true);
  });

  it('deletes all chat history and clears the active chat', () => {
    useStore.getState().setSession(session('s1'));
    useStore.getState().appendMessage(userMessage('s1', 'First chat'));
    useStore.getState().setSession(session('s2'));
    useStore.getState().appendMessage(userMessage('s2', 'Second chat'));
    localStorage.setItem('xl.chat.history.orphan', 'stale');

    useStore.getState().deleteAllChatHistory();

    expect(useStore.getState().currentSession).toBeNull();
    expect(useStore.getState().messages).toEqual([]);
    expect(useStore.getState().chatHistory).toEqual([]);
    expect(localStorage.getItem('xl.chat.history.index')).toBeNull();
    expect(localStorage.getItem('xl.chat.history.s1')).toBeNull();
    expect(localStorage.getItem('xl.chat.history.s2')).toBeNull();
    expect(localStorage.getItem('xl.chat.history.orphan')).toBeNull();
  });
});
