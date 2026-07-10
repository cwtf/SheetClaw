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
      getBaseUrl: () => useStore.getState().appConfig.webAccess.baseUrl,
      getEngineId: () => useStore.getState().appConfig.webAccess.engineId,
    }));
  }
  return layer;
}

export function getTaskpaneAgentLoop(): ReturnType<typeof getAgentLoop> {
  const { registry, executor, snapshots } = getTaskpaneWorkbookLayer();
  return getAgentLoop(registry, executor, snapshots);
}
