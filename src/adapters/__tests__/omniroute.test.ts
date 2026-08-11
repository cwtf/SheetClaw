import { afterEach, describe, expect, it } from 'vitest';
import { createAdapter, isKeyOptionalProvider, OpenAIAdapter } from '../index';
import {
  OMNIROUTE_DEFAULT_BASE_URL,
  OMNIROUTE_NOT_RUNNING,
  diagnoseOmniRouteFailure,
  getOmniRouteBrowserAccessHint,
  isOmniRouteBrowserAccessError,
} from '../omniroute';
import type { ProviderConfig } from '../../types';

const originalFetch = globalThis.fetch;

afterEach(() => {
  globalThis.fetch = originalFetch;
});

const OMNIROUTE_CONFIG: ProviderConfig = {
  provider: 'omniroute',
  enabled: true,
  baseUrl: OMNIROUTE_DEFAULT_BASE_URL,
  model: 'openai/gpt-4o-mini',
  authMode: 'none',
  authStateRef: 'xl.auth.omniroute',
  contextLimits: { maxContextTokens: 128000, historyTokenCap: 100000, maxInlineSheetCells: 5000 },
};

describe('OmniRoute provider wiring', () => {
  it('routes through the OpenAI-compatible adapter', () => {
    expect(createAdapter(OMNIROUTE_CONFIG, '')).toBeInstanceOf(OpenAIAdapter);
  });

  it('treats the gateway key as optional, unlike a vendor API', () => {
    expect(isKeyOptionalProvider('omniroute')).toBe(true);
    expect(isKeyOptionalProvider('ollama')).toBe(true);
    expect(isKeyOptionalProvider('openai')).toBe(false);
  });

  it('defaults to the gateway port OmniRoute documents', () => {
    expect(OMNIROUTE_DEFAULT_BASE_URL).toBe('http://localhost:20128/v1');
  });
});

describe('OmniRoute keyless requests', () => {
  it('sends no Authorization header when no gateway key is set', async () => {
    let sent: Headers | undefined;
    globalThis.fetch = async (_url: RequestInfo | URL, init?: RequestInit) => {
      sent = new Headers(init?.headers);
      return new Response(JSON.stringify({ data: [{ id: 'openai/gpt-4o-mini' }] }), {
        status: 200,
        headers: { 'Content-Type': 'application/json' },
      });
    };

    await createAdapter(OMNIROUTE_CONFIG, '').listModels();
    expect(sent?.has('authorization')).toBe(false);
  });

  it('still sends the gateway key when the user enabled auth on it', async () => {
    let sent: Headers | undefined;
    globalThis.fetch = async (_url: RequestInfo | URL, init?: RequestInit) => {
      sent = new Headers(init?.headers);
      return new Response(JSON.stringify({ data: [] }), {
        status: 200,
        headers: { 'Content-Type': 'application/json' },
      });
    };

    await createAdapter(OMNIROUTE_CONFIG, 'gw-secret').listModels();
    expect(sent?.get('authorization')).toBe('Bearer gw-secret');
  });
});

describe('OmniRoute failure diagnosis', () => {
  it('reports a CORS block when the gateway answers an opaque probe', async () => {
    globalThis.fetch = async () => new Response(null, { status: 200 });

    const hint = await diagnoseOmniRouteFailure(OMNIROUTE_DEFAULT_BASE_URL);
    expect(isOmniRouteBrowserAccessError(hint)).toBe(true);
    expect(hint).toMatch(/Access-Control-Allow-Origin/);
  });

  it('reports a dead gateway when nothing answers at all', async () => {
    globalThis.fetch = async () => {
      throw new TypeError('Failed to fetch');
    };

    const hint = await diagnoseOmniRouteFailure(OMNIROUTE_DEFAULT_BASE_URL);
    expect(hint).toBe(OMNIROUTE_NOT_RUNNING);
    expect(isOmniRouteBrowserAccessError(hint)).toBe(false);
  });

  it('names the add-in origin in the CORS hint when one is known', () => {
    expect(getOmniRouteBrowserAccessHint('https://localhost:3000')).toContain('https://localhost:3000');
  });
});
