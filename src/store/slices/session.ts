import type { StateCreator } from 'zustand';
import type { AgentSession, Message } from '../../types';
import { storage } from '../storage';

const HISTORY_INDEX_KEY = 'xl.chat.history.index';
const HISTORY_SESSION_KEY = (id: string) => `xl.chat.history.${id}`;

const ACTIVE_STATUSES = new Set<AgentSession['status']>([
  'building',
  'calling_llm',
  'parsing',
  'awaiting_confirmation',
  'awaiting_choice',
  'executing_tool',
]);

export interface ChatHistoryItem {
  id: string;
  title: string;
  preview: string;
  createdAt: string;
  updatedAt: string;
  provider: string;
  model: string;
  status: AgentSession['status'];
  messageCount: number;
}

interface ChatTranscript {
  session: AgentSession;
  messages: Message[];
}

function messageText(message: Message): string {
  if (message.role === 'user' || message.role === 'assistant' || message.role === 'system_notice') {
    return message.text;
  }
  if (message.role === 'tool_call') return `Tool: ${message.toolCall.name}`;
  if (message.role === 'tool') return message.result.ok ? 'Tool result: OK' : `Tool error: ${message.result.error?.message ?? 'failed'}`;
  return 'Confirmation requested';
}

function compactText(text: string, fallback: string): string {
  const compact = text.replace(/\s+/g, ' ').trim();
  if (!compact) return fallback;
  return compact.length > 96 ? `${compact.slice(0, 93)}...` : compact;
}

function summarizeTranscript(session: AgentSession, messages: Message[]): ChatHistoryItem {
  const firstUser = messages.find(m => m.role === 'user');
  const latest = [...messages].reverse().find(m => m.role !== 'confirmation');
  const updatedAt = latest?.createdAt ?? session.createdAt;

  return {
    id: session.id,
    title: compactText(firstUser ? messageText(firstUser) : '', 'Untitled chat'),
    preview: compactText(latest ? messageText(latest) : '', 'No messages yet'),
    createdAt: session.createdAt,
    updatedAt,
    provider: session.provider,
    model: session.model,
    status: session.status,
    messageCount: messages.length,
  };
}

function loadIndex(): ChatHistoryItem[] {
  return storage.get<ChatHistoryItem[]>(HISTORY_INDEX_KEY) ?? [];
}

function persistTranscript(session: AgentSession, messages: Message[]): ChatHistoryItem[] {
  const transcript = { session, messages };
  storage.put(HISTORY_SESSION_KEY(session.id), transcript);

  const nextItem = summarizeTranscript(session, messages);
  const index = loadIndex()
    .filter(item => item.id !== session.id)
    .concat(nextItem)
    .sort((a, b) => Date.parse(b.updatedAt) - Date.parse(a.updatedAt));
  storage.put(HISTORY_INDEX_KEY, index);
  return index;
}

function loadTranscript(sessionId: string): ChatTranscript | null {
  const transcript = storage.get<ChatTranscript>(HISTORY_SESSION_KEY(sessionId));
  if (!transcript?.session || !Array.isArray(transcript.messages)) return null;
  return transcript;
}

function sessionForResume(session: AgentSession): AgentSession {
  if (!ACTIVE_STATUSES.has(session.status)) return session;
  return {
    ...session,
    status: 'stopped',
    pendingChange: undefined,
    pendingChoice: undefined,
    lastError: {
      code: 'SessionInterrupted',
      message: 'This chat was restored after an interrupted run.',
    },
  };
}

export interface SessionSlice {
  currentSession: AgentSession | null;
  messages: Message[];
  chatHistory: ChatHistoryItem[];
  webSearchEnabled: boolean;
  setWebSearchEnabled(enabled: boolean): void;
  setSession(session: AgentSession | null): void;
  updateSession(patch: Partial<AgentSession>): void;
  updateSessionById(sessionId: string, patch: Partial<AgentSession>): void;
  appendMessage(msg: Message): void;
  updateMessage(id: string, patch: Partial<Message>): void;
  clearMessages(): void;
  loadChatHistory(): void;
  resumeChat(sessionId: string): boolean;
  deleteChat(sessionId: string): void;
}

export const createSessionSlice: StateCreator<SessionSlice> = set => ({
  currentSession: null,
  messages: [],
  chatHistory: [],
  webSearchEnabled: false,

  setWebSearchEnabled(enabled) {
    set({ webSearchEnabled: enabled });
  },

  setSession(session) {
    if (!session) {
      set({ currentSession: null, messages: [] });
      return;
    }
    const chatHistory = persistTranscript(session, []);
    set({ currentSession: session, messages: [], chatHistory, webSearchEnabled: session.webSearchEnabled });
  },

  updateSession(patch) {
    set(state => {
      if (!state.currentSession) return { currentSession: null };
      const currentSession = { ...state.currentSession, ...patch };
      const chatHistory = persistTranscript(currentSession, state.messages);
      return { currentSession, chatHistory };
    });
  },

  updateSessionById(sessionId, patch) {
    set(state => {
      if (state.currentSession?.id === sessionId) {
        const currentSession = { ...state.currentSession, ...patch };
        const chatHistory = persistTranscript(currentSession, state.messages);
        return { currentSession, chatHistory };
      }

      const transcript = loadTranscript(sessionId);
      if (!transcript) return {};
      const session = { ...transcript.session, ...patch };
      const chatHistory = persistTranscript(session, transcript.messages);
      return { chatHistory };
    });
  },

  appendMessage(msg) {
    set(state => {
      if (state.currentSession?.id === msg.sessionId) {
        const messages = [...state.messages, msg];
        const currentSession = {
          ...state.currentSession,
          messageIds: state.currentSession.messageIds.includes(msg.id)
            ? state.currentSession.messageIds
            : [...state.currentSession.messageIds, msg.id],
        };
        const chatHistory = persistTranscript(currentSession, messages);
        return { currentSession, messages, chatHistory };
      }

      const transcript = loadTranscript(msg.sessionId);
      if (!transcript) return {};
      const messages = [...transcript.messages, msg];
      const session = {
        ...transcript.session,
        messageIds: transcript.session.messageIds.includes(msg.id)
          ? transcript.session.messageIds
          : [...transcript.session.messageIds, msg.id],
      };
      const chatHistory = persistTranscript(session, messages);
      return { chatHistory };
    });
  },

  updateMessage(id, patch) {
    set(state => {
      if (state.messages.some(m => m.id === id) && state.currentSession) {
        const messages = state.messages.map(m =>
          m.id === id ? ({ ...m, ...patch } as Message) : m
        );
        const chatHistory = persistTranscript(state.currentSession, messages);
        return { messages, chatHistory };
      }

      for (const item of state.chatHistory) {
        const transcript = loadTranscript(item.id);
        if (!transcript?.messages.some(m => m.id === id)) continue;
        const messages = transcript.messages.map(m =>
          m.id === id ? ({ ...m, ...patch } as Message) : m
        );
        const chatHistory = persistTranscript(transcript.session, messages);
        return { chatHistory };
      }
      return {};
    });
  },

  clearMessages() {
    set(state => {
      if (!state.currentSession) return { messages: [] };
      const chatHistory = persistTranscript(state.currentSession, []);
      return { messages: [], chatHistory };
    });
  },

  loadChatHistory() {
    set({ chatHistory: loadIndex() });
  },

  resumeChat(sessionId) {
    const transcript = loadTranscript(sessionId);
    if (!transcript) return false;

    const session = sessionForResume(transcript.session);
    const chatHistory = persistTranscript(session, transcript.messages);
    set({
      currentSession: session,
      messages: transcript.messages,
      chatHistory,
      webSearchEnabled: session.webSearchEnabled,
    });
    return true;
  },

  deleteChat(sessionId) {
    storage.remove(HISTORY_SESSION_KEY(sessionId));
    const chatHistory = loadIndex().filter(item => item.id !== sessionId);
    storage.put(HISTORY_INDEX_KEY, chatHistory);
    set(state => ({
      chatHistory,
      ...(state.currentSession?.id === sessionId ? { currentSession: null, messages: [] } : {}),
    }));
  },
});
