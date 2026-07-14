import { describe, expect, it } from 'vitest';
import type { Message, ToolCall } from '../../types';
import { buildChatRenderItems } from './ChatPanel';

const createdAt = '2026-07-15T00:00:00.000Z';
const sessionId = 'session-1';

function toolCall(id: string): ToolCall {
  return {
    id,
    name: 'web_search',
    arguments: {},
    workbookId: 'workbook-1',
    mutating: false,
  };
}

describe('buildChatRenderItems', () => {
  it('groups tool-calling narration while leaving the final answer visible', () => {
    const firstCall = toolCall('call-1');
    const secondCall = toolCall('call-2');
    const messages: Message[] = [
      { id: 'user', sessionId, createdAt, role: 'user', text: 'Find the data' },
      {
        id: 'thought-1',
        sessionId,
        createdAt,
        role: 'assistant',
        text: 'I will search the catalogue.',
        toolCalls: [firstCall],
      },
      { id: 'call-message-1', sessionId, createdAt, role: 'tool_call', toolCall: firstCall, status: 'applied' },
      { id: 'result-1', sessionId, createdAt, role: 'tool', toolCallId: firstCall.id, result: { toolCallId: firstCall.id, ok: true } },
      {
        id: 'thought-2',
        sessionId,
        createdAt,
        role: 'assistant',
        text: 'I found a more specific dataset.',
        toolCalls: [secondCall],
      },
      { id: 'call-message-2', sessionId, createdAt, role: 'tool_call', toolCall: secondCall, status: 'applied' },
      { id: 'result-2', sessionId, createdAt, role: 'tool', toolCallId: secondCall.id, result: { toolCallId: secondCall.id, ok: true } },
      { id: 'final', sessionId, createdAt, role: 'assistant', text: 'Here is the requested data.' },
    ];

    const items = buildChatRenderItems(messages);

    expect(items.map(item => item.type)).toEqual([
      'message',
      'thought_process',
      'tool_chain',
      'message',
    ]);
    const thoughts = items.find(item => item.type === 'thought_process');
    expect(thoughts?.messages.map(message => message.id)).toEqual(['thought-1', 'thought-2']);
    const finalItem = items[items.length - 1];
    expect(finalItem?.type).toBe('message');
    if (finalItem?.type === 'message') expect(finalItem.message.id).toBe('final');
  });
});
