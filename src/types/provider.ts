import type { ProviderKey } from './usage';
import type { SearchProviderId, WebAccessProvider } from '../web/providers';

export interface ModelInfo {
  id: string;
  name?: string;
  contextWindow?: number;
  supportsTools?: boolean;
}

export interface ProviderCapabilities {
  supportsTools: boolean;
  supportsStreaming: boolean;
  supportsOAuth: boolean;
  nativeUsage: boolean;
  toolFormat: 'openai' | 'anthropic';
}

export interface ProviderConfig {
  provider: ProviderKey;
  label?: string;
  enabled: boolean;
  baseUrl: string;
  model: string;
  knownModels?: ModelInfo[];
  authMode: 'apikey' | 'oauth' | 'none';
  authStateRef: string;
  headers?: Record<string, string>;
  contextLimits: {
    maxContextTokens: number;
    historyTokenCap: number;
    maxInlineSheetCells: number;
  };
  temperature?: number;
  maxOutputTokens?: number;
}

export interface AuthState {
  provider: string;
  state:
    | 'unauthenticated'
    | 'authenticating'
    | 'authenticated'
    | 'token-expired'
    | 'validating'
    | 'error';
  apiKeyMasked?: string;
  authMode?: 'apikey' | 'oauth' | 'none';
  oauthProvider?: 'openrouter';
  accessToken?: string;
  refreshToken?: string;
  tokenType?: string;
  expiresAt?: string;
  userId?: string;
  error?: string;
  /** Raw key in memory only; persisted AES-GCM-sealed via auth/secureStore. */
  _key?: string;
}

export interface WebAccessConfig {
  provider: WebAccessProvider;
  baseUrl?: string;
  /** Google CSE Programmable Search Engine id (cx); unused by other providers. */
  engineId?: string;
  /**
   * Base URL of the local Agent-Reach bridge. Kept separate from `baseUrl`
   * because it governs the fetch_url platform backend too, which must work
   * whichever search provider is selected. Absent disables both.
   */
  agentReachBaseUrl?: string;
  readerFallback: boolean;
}

export type { SearchProviderId, WebAccessProvider };
