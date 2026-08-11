import { beforeEach, describe, expect, it, vi } from 'vitest';
import { useStore } from '../index';

function makeLocalStorageStub() {
  const map = new Map<string, string>();
  return {
    getItem: (key: string) => map.get(key) ?? null,
    setItem: (key: string, value: string) => map.set(key, String(value)),
    removeItem: (key: string) => map.delete(key),
    clear: () => map.clear(),
    key: (index: number) => Array.from(map.keys())[index] ?? null,
    get length() {
      return map.size;
    },
  };
}

beforeEach(() => {
  vi.stubGlobal('localStorage', makeLocalStorageStub());
  useStore.setState(state => ({
    appConfig: { ...state.appConfig, theme: 'system' },
  }));
});

describe('theme persistence', () => {
  it('defaults legacy configs to the system theme', () => {
    localStorage.setItem(
      'xl.config.app',
      JSON.stringify({
        _v: 1,
        activeProvider: 'ollama',
        autoApproveSession: false,
        pricingMode: 'bundled',
        webAccess: { provider: 'none', readerFallback: false },
      }),
    );
    useStore.setState(state => ({
      appConfig: { ...state.appConfig, theme: 'dark' },
    }));
    useStore.getState().loadConfigFromStorage();

    expect(useStore.getState().appConfig.theme).toBe('system');
  });

  it('saves and restores the dark theme preference', () => {
    useStore.getState().setAppConfig({ theme: 'dark' });
    expect(useStore.getState().appConfig.theme).toBe('dark');

    useStore.setState(state => ({
      appConfig: { ...state.appConfig, theme: 'light' },
    }));
    useStore.getState().loadConfigFromStorage();

    expect(useStore.getState().appConfig.theme).toBe('dark');
  });
});

describe('OmniRoute provider defaults', () => {
  it('ships a keyless local-gateway config pointing at OmniRoute', () => {
    const cfg = useStore.getState().providers.omniroute;

    expect(cfg.baseUrl).toBe('http://localhost:20128/v1');
    expect(cfg.authMode).toBe('none');
    expect(cfg.enabled).toBe(false);
    // The model list is whatever the user wired into their own gateway, so
    // there is nothing sensible to preselect.
    expect(cfg.model).toBe('');
  });

  it('reports readiness without a stored credential', () => {
    expect(useStore.getState().isProviderReady('omniroute')).toBe(true);
  });

  it('survives a round trip through storage alongside older stored providers', () => {
    localStorage.setItem(
      'xl.config.providers',
      JSON.stringify({ _v: 1, ollama: useStore.getState().providers.ollama }),
    );
    useStore.getState().loadConfigFromStorage();

    expect(useStore.getState().providers.omniroute.baseUrl).toBe('http://localhost:20128/v1');
  });
});
