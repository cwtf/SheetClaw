import type { StateCreator } from 'zustand';
import type { AuthState, ProviderKey, SearchProviderId } from '../../types';
import { SEARCH_PROVIDERS as SEARCH_PROVIDER_REGISTRY, SEARCH_PROVIDER_IDS } from '../../web/providers';
import { storage } from '../storage';
import { encryptSecret, decryptSecret, isEncryptedSecret } from '../../auth/secureStore';

const AUTH_KEY = (p: ProviderKey) => `xl.auth.${p}`;
const SEARCH_AUTH_KEY = (p: SearchProviderId) => `xl.auth.search:${p}`;

function maskKey(key: string): string {
  if (key.length <= 8) return '••••••••';
  return key.slice(0, 4) + '••••' + key.slice(-4);
}

// ── Encryption at rest ─────────────────────────────────────────────────────
// In-memory AuthState keeps plaintext secrets (adapters need them per call);
// the persisted copy carries them AES-GCM-sealed via auth/secureStore.

const SENSITIVE_FIELDS = ['_key', 'accessToken', 'refreshToken'] as const;

async function sealAuthState(state: AuthState): Promise<AuthState> {
  const sealed = { ...state };
  for (const field of SENSITIVE_FIELDS) {
    const value = sealed[field];
    if (value) sealed[field] = await encryptSecret(value);
  }
  return sealed;
}

async function openAuthState(saved: AuthState): Promise<{ state: AuthState; migrated: boolean }> {
  const opened = { ...saved };
  let migrated = false;
  for (const field of SENSITIVE_FIELDS) {
    const value = opened[field];
    if (!value) continue;
    if (!isEncryptedSecret(value)) migrated = true;
    opened[field] = await decryptSecret(value);
  }
  return { state: opened, migrated };
}

// All persistence goes through one queue so writes land in call order —
// otherwise a clear could be overtaken by a still-encrypting earlier save.
let pendingWrites: Promise<void> = Promise.resolve();

function persistSealed(storageKey: string, state: AuthState): void {
  pendingWrites = pendingWrites
    .then(async () => {
      storage.put(storageKey, await sealAuthState(state));
    })
    .catch(() => {
      // Keep the queue alive; a failed write degrades to session-only auth.
    });
}

/** Resolves once every queued credential write has been flushed to storage. */
export function flushAuthPersistence(): Promise<void> {
  return pendingWrites;
}

async function openSavedAuth(
  storageKey: string,
  saved: AuthState,
  providerLabel: string
): Promise<AuthState> {
  try {
    const { state, migrated } = await openAuthState(saved);
    // Legacy plaintext entry: re-persist it sealed.
    if (migrated) persistSealed(storageKey, state);
    return state;
  } catch {
    // Key lost or ciphertext tampered — drop the secret and ask again.
    return {
      provider: providerLabel,
      state: 'unauthenticated',
      error: 'Saved credential could not be unlocked. Enter it again in Settings.',
    };
  }
}

export interface AuthSlice {
  authStates: Record<ProviderKey, AuthState>;
  searchAuthStates: Record<SearchProviderId, AuthState>;
  setAuthState(provider: ProviderKey, state: Partial<AuthState>): void;
  saveApiKey(provider: ProviderKey, key: string): void;
  saveOAuthCredential(provider: ProviderKey, credential: {
    accessToken: string;
    oauthProvider?: AuthState['oauthProvider'];
    userId?: string;
    refreshToken?: string;
    tokenType?: string;
    expiresAt?: string;
  }): void;
  clearApiKey(provider: ProviderKey): void;
  saveSearchApiKey(provider: SearchProviderId, key: string): void;
  clearSearchApiKey(provider: SearchProviderId): void;
  loadAuthFromStorage(): Promise<void>;
  isProviderReady(provider: ProviderKey): boolean;
  isSearchProviderReady(provider: SearchProviderId): boolean;
}

const DEFAULT_AUTH_STATE = (provider: string): AuthState => ({
  provider,
  state: 'unauthenticated',
});

const ALL_PROVIDERS: ProviderKey[] = [
  'ollama',
  'openai',
  'anthropic',
  'generic',
  'deepseek',
  'groq',
  'mistral',
  'together',
  'kimi',
  'glm',
  'qwen',
  'llama',
  'gemini',
  'cerebras',
  'cloudflare',
  'huggingface',
];

const SEARCH_PROVIDERS: SearchProviderId[] = SEARCH_PROVIDER_IDS;

export const createAuthSlice: StateCreator<AuthSlice> = (set, get) => ({
  authStates: Object.fromEntries(
    ALL_PROVIDERS.map(p => [p, DEFAULT_AUTH_STATE(p)])
  ) as Record<ProviderKey, AuthState>,
  searchAuthStates: Object.fromEntries(
    SEARCH_PROVIDERS.map(p => [p, DEFAULT_AUTH_STATE(`search:${p}`)])
  ) as Record<SearchProviderId, AuthState>,

  setAuthState(provider, patch) {
    set(state => ({
      authStates: {
        ...state.authStates,
        [provider]: { ...state.authStates[provider], ...patch },
      },
    }));
  },

  saveApiKey(provider, key) {
    const trimmed = key.trim();
    const next: AuthState = {
      provider,
      state: trimmed ? 'authenticated' : 'unauthenticated',
      authMode: trimmed ? 'apikey' : undefined,
      apiKeyMasked: trimmed ? maskKey(trimmed) : undefined,
      _key: trimmed || undefined,
    };
    persistSealed(AUTH_KEY(provider), next);
    set(state => ({ authStates: { ...state.authStates, [provider]: next } }));
  },

  saveOAuthCredential(provider, credential) {
    const trimmed = credential.accessToken.trim();
    const next: AuthState = {
      provider,
      state: trimmed ? 'authenticated' : 'unauthenticated',
      authMode: trimmed ? 'oauth' : undefined,
      oauthProvider: credential.oauthProvider,
      apiKeyMasked: trimmed ? maskKey(trimmed) : undefined,
      accessToken: trimmed || undefined,
      refreshToken: credential.refreshToken,
      tokenType: credential.tokenType ?? 'Bearer',
      expiresAt: credential.expiresAt,
      userId: credential.userId,
      _key: trimmed || undefined,
    };
    persistSealed(AUTH_KEY(provider), next);
    set(state => ({ authStates: { ...state.authStates, [provider]: next } }));
  },

  clearApiKey(provider) {
    const next: AuthState = { provider, state: 'unauthenticated' };
    persistSealed(AUTH_KEY(provider), next);
    set(state => ({ authStates: { ...state.authStates, [provider]: next } }));
  },

  saveSearchApiKey(provider, key) {
    const trimmed = key.trim();
    const next: AuthState = {
      provider: `search:${provider}`,
      state: trimmed ? 'authenticated' : 'unauthenticated',
      authMode: trimmed ? 'apikey' : undefined,
      apiKeyMasked: trimmed ? maskKey(trimmed) : undefined,
      _key: trimmed || undefined,
    };
    persistSealed(SEARCH_AUTH_KEY(provider), next);
    set(state => ({ searchAuthStates: { ...state.searchAuthStates, [provider]: next } }));
  },

  clearSearchApiKey(provider) {
    const next: AuthState = { provider: `search:${provider}`, state: 'unauthenticated' };
    persistSealed(SEARCH_AUTH_KEY(provider), next);
    set(state => ({ searchAuthStates: { ...state.searchAuthStates, [provider]: next } }));
  },

  async loadAuthFromStorage() {
    const loaded: Partial<Record<ProviderKey, AuthState>> = {};
    for (const p of ALL_PROVIDERS) {
      const saved = storage.get<AuthState>(AUTH_KEY(p));
      if (saved) loaded[p] = await openSavedAuth(AUTH_KEY(p), saved, p);
    }
    const loadedSearch: Partial<Record<SearchProviderId, AuthState>> = {};
    for (const p of SEARCH_PROVIDERS) {
      const saved = storage.get<AuthState>(SEARCH_AUTH_KEY(p));
      if (saved) loadedSearch[p] = await openSavedAuth(SEARCH_AUTH_KEY(p), saved, `search:${p}`);
    }
    if (Object.keys(loaded).length > 0) {
      set(state => ({
        authStates: { ...state.authStates, ...loaded },
      }));
    }
    if (Object.keys(loadedSearch).length > 0) {
      set(state => ({
        searchAuthStates: { ...state.searchAuthStates, ...loadedSearch },
      }));
    }
  },

  isProviderReady(provider) {
    const auth = get().authStates[provider];
    const s = auth?.state;
    if (provider === 'ollama') return s === 'unauthenticated' || s === 'authenticated';
    if (auth?.expiresAt && Date.parse(auth.expiresAt) <= Date.now() + 60_000) return false;
    return s === 'authenticated';
  },

  isSearchProviderReady(provider) {
    if (SEARCH_PROVIDER_REGISTRY[provider]?.requiresKey === false) return true;
    const auth = get().searchAuthStates[provider];
    return auth?.state === 'authenticated';
  },
});
