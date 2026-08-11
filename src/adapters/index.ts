export { OpenAIAdapter } from './openai';
export type { OpenAIAdapterConfig } from './openai';
export { AnthropicAdapter } from './anthropic';
export type { AnthropicAdapterConfig } from './anthropic';
export { OllamaAdapter } from './ollama';
export type { OllamaAdapterConfig } from './ollama';
export { parseLenientToolCall } from './ollama';
export {
  OMNIROUTE_DEFAULT_BASE_URL,
  OMNIROUTE_NOT_RUNNING,
  diagnoseOmniRouteFailure,
  getOmniRouteBrowserAccessHint,
  isOmniRouteBrowserAccessError,
} from './omniroute';

import type { AuthState, LLMClient, ProviderConfig, ProviderKey } from '../types';
import { OpenAIAdapter } from './openai';
import { AnthropicAdapter } from './anthropic';
import { OllamaAdapter } from './ollama';
import { getAuthCredential } from '../auth/credentials';

/**
 * Providers that reach a server on the user's own machine rather than a vendor
 * API, so a credential is optional: Ollama has no auth at all, and OmniRoute's
 * gateway key only matters if the user turned it on. The UI uses this to keep
 * the key field optional and to treat an unauthenticated state as ready.
 */
const KEY_OPTIONAL_PROVIDERS: ReadonlySet<ProviderKey> = new Set(['ollama', 'omniroute']);

export function isKeyOptionalProvider(provider: ProviderKey): boolean {
  return KEY_OPTIONAL_PROVIDERS.has(provider);
}

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
    case 'omniroute':
      return new OpenAIAdapter({ apiKey, baseUrl: cfg.baseUrl, provider: cfg.provider, extraHeaders: cfg.headers });
  }
}
