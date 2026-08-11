import { getAgentLoop } from '../agent/index';
import { createWorkbookLayer } from '../workbook/index';
import { getAuthCredential } from '../auth/credentials';
import { useStore } from '../store/index';
import { FETCH_URL, createFetchUrlHandler } from '../web/fetch';
import { WEB_SEARCH, createWebSearchHandler } from '../web/search';
import { KEYLESS_BUNDLE_ID } from '../web/providers';
import type { SearchProviderId } from '../types';

let layer: ReturnType<typeof createWorkbookLayer> | null = null;

export function getTaskpaneWorkbookLayer(): ReturnType<typeof createWorkbookLayer> {
  if (!layer) {
    layer = createWorkbookLayer();
    layer.executor.register(FETCH_URL, createFetchUrlHandler({
      readerFallback: () => useStore.getState().appConfig.webAccess.readerFallback,
      agentReachBaseUrl: () => useStore.getState().appConfig.webAccess.agentReachBaseUrl,
    }));
    layer.executor.register(WEB_SEARCH, createWebSearchHandler({
      // Effective backend: the configured keyed provider only when the Search
      // toggle is on and the provider is ready; otherwise the keyless bundle.
      getProvider: () => {
        const s = useStore.getState();
        const p = s.appConfig.webAccess.provider;
        return p !== 'none' && s.webSearchEnabled && s.isSearchProviderReady(p)
          ? p
          : KEYLESS_BUNDLE_ID;
      },
      getApiKey: (provider: SearchProviderId) =>
        getAuthCredential(useStore.getState().searchAuthStates[provider]),
      // The bridge URL is configured once and drives both the search provider
      // and the fetch_url backend, so agent-reach reads its own field.
      getBaseUrl: (provider: SearchProviderId) => {
        const web = useStore.getState().appConfig.webAccess;
        return provider === 'agent-reach' ? (web.agentReachBaseUrl ?? web.baseUrl) : web.baseUrl;
      },
      getEngineId: () => useStore.getState().appConfig.webAccess.engineId,
    }));
  }
  return layer;
}

export function getTaskpaneAgentLoop(): ReturnType<typeof getAgentLoop> {
  const { registry, executor, snapshots } = getTaskpaneWorkbookLayer();
  return getAgentLoop(registry, executor, snapshots);
}
