import type { ToolCall, ToolResult } from './tool';

export type Message =
  | UserMessage
  | AssistantMessage
  | ToolCallMessage
  | ToolResultMessage
  | ConfirmationMessage
  | SystemNoticeMessage;

interface BaseMessage {
  id: string;
  sessionId: string;
  createdAt: string;
}

export interface UserMessage extends BaseMessage {
  role: 'user';
  text: string;
  selection?: WorkbookSelection;
}

export interface WorkbookSelection {
  sheet: string;
  address: string;
}

export interface AssistantMessage extends BaseMessage {
  role: 'assistant';
  text: string;
  toolCalls?: ToolCall[];
  usageRef?: string;
  finishReason?: 'stop' | 'tool_calls' | 'length' | 'error';
}

export interface ToolCallMessage extends BaseMessage {
  role: 'tool_call';
  toolCall: ToolCall;
  status: 'pending' | 'awaiting_confirmation' | 'applied' | 'failed';
}

export interface ToolResultMessage extends BaseMessage {
  role: 'tool';
  toolCallId: string;
  result: ToolResult;
}

export interface ConfirmationMessage extends BaseMessage {
  role: 'confirmation';
  pendingChangeId: string;
  decision?: 'apply' | 'cancel' | 'apply_all';
}

export interface SystemNoticeMessage extends BaseMessage {
  role: 'system_notice';
  level: 'info' | 'warn' | 'error';
  text: string;
}
