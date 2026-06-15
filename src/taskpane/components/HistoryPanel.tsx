import { useEffect } from 'react';
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
  const resetSessionTotals = useStore(s => s.resetSessionTotals);

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
      <div style={{ flexShrink: 0 }}>
        <Body1Strong>History</Body1Strong>
        <Caption1 style={{
          display: 'block',
          marginTop: 2,
          color: tokens.colorNeutralForeground3,
        }}>
          Resume a previous workbook chat.
        </Caption1>
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
            return (
              <button
                key={item.id}
                type="button"
                onClick={() => openHistoryItem(item.id)}
                style={{
                  display: 'flex',
                  flexDirection: 'column',
                  alignItems: 'stretch',
                  gap: 6,
                  width: '100%',
                  minWidth: 0,
                  padding: 10,
                  textAlign: 'left',
                  borderRadius: 6,
                  border: `1px solid ${selected ? tokens.colorBrandStroke1 : tokens.colorNeutralStroke1}`,
                  background: selected ? tokens.colorBrandBackground2 : tokens.colorNeutralBackground1,
                  color: tokens.colorNeutralForeground1,
                  cursor: 'pointer',
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
