import { useEffect, useState } from 'react';
import {
  Body1,
  Body1Strong,
  Button,
  Caption1,
  tokens,
} from '@fluentui/react-components';
import { useStore } from '../../store/index';
import type { ChatHistoryItem } from '../../store/slices/session';
import { getTaskpaneAgentLoop } from '../workbookLayer';

const RUNNING_STATUSES = new Set(['building', 'calling_llm', 'parsing', 'executing_tool']);

function formatDate(value: string): string {
  const date = new Date(value);
  if (Number.isNaN(date.getTime())) return '';
  return new Intl.DateTimeFormat(undefined, {
    month: 'short',
    day: 'numeric',
    hour: 'numeric',
    minute: '2-digit',
  }).format(date);
}

function statusLabel(status: ChatHistoryItem['status']): string {
  if (status === 'done') return 'Done';
  if (status === 'error') return 'Error';
  if (status === 'stopped') return 'Stopped';
  if (status === 'awaiting_confirmation') return 'Confirming';
  if (status === 'awaiting_choice') return 'Choosing';
  if (RUNNING_STATUSES.has(status)) return 'Running';
  return 'Idle';
}

export default function HistoryPanel({ onOpenChat }: { onOpenChat: () => void }) {
  const history = useStore(s => s.chatHistory);
  const currentSessionId = useStore(s => s.currentSession?.id);
  const loadChatHistory = useStore(s => s.loadChatHistory);
  const resumeChat = useStore(s => s.resumeChat);
  const deleteChat = useStore(s => s.deleteChat);
  const deleteAllChatHistory = useStore(s => s.deleteAllChatHistory);
  const resetSessionTotals = useStore(s => s.resetSessionTotals);
  const [pendingDeleteId, setPendingDeleteId] = useState<string | 'all' | null>(null);

  useEffect(() => {
    loadChatHistory();
  }, [loadChatHistory]);

  function openHistoryItem(id: string) {
    getTaskpaneAgentLoop().stop();
    if (!resumeChat(id)) return;
    const restored = useStore.getState().currentSession;
    if (restored) {
      resetSessionTotals(restored.id, {
        inputTokens: restored.totals.inputTokens,
        outputTokens: restored.totals.outputTokens,
        costUsd: restored.totals.costUsd,
      });
    }
    onOpenChat();
  }

  function deleteHistoryItem(item: ChatHistoryItem) {
    if (item.id === currentSessionId) getTaskpaneAgentLoop().stop();
    deleteChat(item.id);
    setPendingDeleteId(null);
  }

  function deleteAllHistory() {
    getTaskpaneAgentLoop().stop();
    deleteAllChatHistory();
    setPendingDeleteId(null);
  }

  return (
    <div style={{
      display: 'flex',
      flexDirection: 'column',
      height: '100%',
      minHeight: 0,
      padding: 12,
      gap: 10,
      boxSizing: 'border-box',
      overflowX: 'hidden',
    }}>
      <div style={{
        flexShrink: 0,
        display: 'flex',
        alignItems: 'flex-start',
        justifyContent: 'space-between',
        gap: 8,
      }}>
        <div style={{ minWidth: 0 }}>
          <Body1Strong>History</Body1Strong>
          <Caption1 style={{
            display: 'block',
            marginTop: 2,
            color: tokens.colorNeutralForeground3,
          }}>
            Resume a previous workbook chat.
          </Caption1>
        </div>
        {history.length > 0 && pendingDeleteId !== 'all' && (
          <Button
            appearance="subtle"
            size="small"
            type="button"
            onClick={() => setPendingDeleteId('all')}
            style={{
              flexShrink: 0,
              color: tokens.colorPaletteRedForeground1,
            }}
          >
            Delete all
          </Button>
        )}
        {history.length > 0 && pendingDeleteId === 'all' && (
          <div style={{
            display: 'flex',
            gap: 4,
            flexShrink: 0,
          }}>
            <Button
              appearance="primary"
              size="small"
              type="button"
              onClick={deleteAllHistory}
            >
              Delete
            </Button>
            <Button
              appearance="secondary"
              size="small"
              type="button"
              onClick={() => setPendingDeleteId(null)}
            >
              Cancel
            </Button>
          </div>
        )}
      </div>

      {history.length === 0 ? (
        <div style={{
          margin: 'auto 0',
          display: 'flex',
          flexDirection: 'column',
          gap: 4,
        }}>
          <Body1>No chat history yet</Body1>
          <Caption1 style={{ color: tokens.colorNeutralForeground3 }}>
            Chats appear here after you start a conversation.
          </Caption1>
        </div>
      ) : (
        <div style={{
          display: 'flex',
          flexDirection: 'column',
          gap: 8,
          minHeight: 0,
          overflowY: 'auto',
          overflowX: 'hidden',
        }}>
          {history.map(item => {
            const selected = item.id === currentSessionId;
            const pendingDelete = pendingDeleteId === item.id;
            return (
              <div
                key={item.id}
                style={{
                  display: 'flex',
                  alignItems: 'flex-start',
                  gap: 4,
                  width: '100%',
                  minWidth: 0,
                  padding: 4,
                  borderRadius: 6,
                  border: `1px solid ${selected ? tokens.colorBrandStroke1 : tokens.colorNeutralStroke1}`,
                  background: selected ? tokens.colorBrandBackground2 : tokens.colorNeutralBackground1,
                  color: tokens.colorNeutralForeground1,
                  boxSizing: 'border-box',
                }}
              >
                <button
                  type="button"
                  onClick={() => openHistoryItem(item.id)}
                  style={{
                    display: 'flex',
                    flexDirection: 'column',
                    alignItems: 'stretch',
                    gap: 6,
                    flex: 1,
                    minWidth: 0,
                    padding: 6,
                    textAlign: 'left',
                    border: 0,
                    background: 'transparent',
                    color: 'inherit',
                    cursor: 'pointer',
                    font: 'inherit',
                  }}
                  aria-current={selected ? 'true' : undefined}
                >
                  <span style={{
                    display: 'flex',
                    alignItems: 'center',
                    justifyContent: 'space-between',
                    gap: 8,
                    minWidth: 0,
                  }}>
                    <Body1Strong style={{
                      minWidth: 0,
                      overflow: 'hidden',
                      textOverflow: 'ellipsis',
                      whiteSpace: 'nowrap',
                    }}>
                      {item.title}
                    </Body1Strong>
                    <Caption1 style={{
                      flexShrink: 0,
                      padding: '1px 6px',
                      borderRadius: 999,
                      border: `1px solid ${tokens.colorNeutralStroke1}`,
                      color: tokens.colorNeutralForeground3,
                      background: tokens.colorNeutralBackground2,
                    }}>
                      {statusLabel(item.status)}
                    </Caption1>
                  </span>

                  <Caption1 style={{
                    color: tokens.colorNeutralForeground2,
                    overflow: 'hidden',
                    textOverflow: 'ellipsis',
                    whiteSpace: 'nowrap',
                  }}>
                    {item.preview}
                  </Caption1>

                  <span style={{
                    display: 'flex',
                    justifyContent: 'space-between',
                    gap: 8,
                    minWidth: 0,
                  }}>
                    <Caption1 style={{
                      color: tokens.colorNeutralForeground3,
                      overflow: 'hidden',
                      textOverflow: 'ellipsis',
                      whiteSpace: 'nowrap',
                    }}>
                      {item.model || item.provider}
                    </Caption1>
                    <Caption1 style={{ color: tokens.colorNeutralForeground3, flexShrink: 0 }}>
                      {formatDate(item.updatedAt)} | {item.messageCount}
                    </Caption1>
                  </span>
                </button>
                {pendingDelete ? (
                  <div style={{
                    display: 'flex',
                    flexDirection: 'column',
                    gap: 4,
                    flexShrink: 0,
                  }}>
                    <Button
                      appearance="primary"
                      size="small"
                      type="button"
                      onClick={() => deleteHistoryItem(item)}
                      style={{ minWidth: 56 }}
                    >
                      Delete
                    </Button>
                    <Button
                      appearance="secondary"
                      size="small"
                      type="button"
                      onClick={() => setPendingDeleteId(null)}
                      style={{ minWidth: 56 }}
                    >
                      Cancel
                    </Button>
                  </div>
                ) : (
                  <Button
                    appearance="subtle"
                    size="small"
                    type="button"
                    aria-label={`Delete ${item.title}`}
                    title="Delete"
                    onClick={() => setPendingDeleteId(item.id)}
                    style={{
                      flexShrink: 0,
                      minWidth: 32,
                      color: tokens.colorPaletteRedForeground1,
                    }}
                  >
                    {'\u{1F5D1}\uFE0F'}
                  </Button>
                )}
              </div>
            );
          })}
        </div>
      )}

      {history.length > 0 && (
        <Button appearance="secondary" size="small" onClick={onOpenChat} style={{ flexShrink: 0 }}>
          Back to chat
        </Button>
      )}
    </div>
  );
}
