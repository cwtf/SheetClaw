export { OpenAIAdapter } from './openai';
export type { OpenAIAdapterConfig } from './openai';
export { AnthropicAdapter } from './anthropic';
export type { AnthropicAdapterConfig } from './anthropic';
export { OllamaAdapter } from './ollama';
export type { OllamaAdapterConfig } from './ollama';
export { parseLenientToolCall } from './ollama';

import type { AuthState, LLMClient, ProviderConfig } from '../types';
import { OpenAIAdapter } from './openai';
import { AnthropicAdapter } from './anthropic';
import { OllamaAdapter } from './ollama';
import { getAuthCredential } from '../auth/credentials';

export function createAdapter(cfg: ProviderConfig, auth: string | AuthState = ''): LLMClient {
  const apiKey = typeof auth === 'string' ? auth : getAuthCredential(auth);
  switch (cfg.provider) {
    case 'anthropic':
      return new AnthropicAdapter({ apiKey, baseUrl: cfg.baseUrl, provider: cfg.provider });
    case 'ollama':
      return new OllamaAdapter({ baseUrl: cfg.baseUrl });
    case 'openai':
    case 'generic':
    case 'deepseek':
    case 'groq':
    case 'mistral':
    case 'together':
    case 'kimi':
    case 'glm':
    case 'qwen':
    case 'llama':
    case 'gemini':
    case 'cerebras':
    case 'cloudflare':
    case 'huggingface':
      return new OpenAIAdapter({ apiKey, baseUrl: cfg.baseUrl, provider: cfg.provider, extraHeaders: cfg.headers });
  }
}
