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
