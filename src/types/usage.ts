export type ProviderKey =
  | 'ollama'
  | 'openai'
  | 'anthropic'
  | 'generic'
  | 'deepseek'
  | 'groq'
  | 'mistral'
  | 'together'
  | 'kimi'
  | 'glm'
  | 'qwen'
  | 'llama'
  | 'gemini'
  | 'cerebras'
  | 'cloudflare'
  | 'huggingface';

export interface UsageRecord {
  id: string;
  sessionId: string;
  turnIndex: number;
  timestamp: string;
  provider: ProviderKey;
  model: string;
  inputTokens: number;
  outputTokens: number;
  cacheReadTokens?: number;
  cacheWriteTokens?: number;
  totalTokens: number;
  estimatedCostUsd: number;
  pricingVersion?: string;
  estimated: boolean;
  toolCallsCount: number;
}

export interface PricingEntry {
  provider: string;
  modelMatch: string;
  inputPerMTok: number;
  outputPerMTok: number;
  cacheReadPerMTok?: number;
  cacheWritePerMTok?: number;
  currency: 'USD';
}

export interface PricingTable {
  version: string;
  updatedAt: string;
  entries: PricingEntry[];
  defaults: { inputPerMTok: number; outputPerMTok: number };
}
