import { ToolNetworkError } from '../../workbook/executor';
import { SEARCH_PROVIDERS, SEARCH_PROVIDER_IDS, type SearchProviderAdapter, type SearchProviderId, type SearchResult } from './index';

/** Id of the meta-provider that bundles every keyless source behind one selection. */
export const KEYLESS_BUNDLE_ID = 'keyless';
export type KeylessBundleId = typeof KEYLESS_BUNDLE_ID;

/**
 * One-line domain hints per keyless source, surfaced in the web_search tool
 * schema so the model can route each query to the best catalogue.
 */
export const KEYLESS_SOURCE_HINTS: Partial<Record<SearchProviderId, string>> = {
  wikipedia: 'encyclopedic articles',
  wikidata: 'structured entities and identifiers',
  worldbank: 'World Bank development indicators, country-level statistics',
  ckan: 'Australian open-data catalogue (data.gov.au)',
  'data-gov-my': 'Malaysia official open-data APIs including OpenDOSM statistics',
  'data-gov-sg': 'Singapore official datasets',
  ecb: 'European Central Bank euro-area finance and statistics',
  eurostat: 'EU official statistics',
  imf: 'IMF macroeconomic indicators by country',
  'open-meteo': 'weather forecasts for coordinates',
  'un-sdg': 'UN Sustainable Development Goal indicators',
};

/** General-purpose sources queried when no explicit source is requested. */
const AUTO_SOURCES: SearchProviderId[] = ['wikipedia', 'wikidata', 'worldbank'];

/** Computed lazily: SEARCH_PROVIDERS is not yet initialised while this module evaluates (import cycle with index.ts). */
export function keylessSourceIds(): SearchProviderId[] {
  return SEARCH_PROVIDER_IDS.filter(id => {
    const p = SEARCH_PROVIDERS[id];
    return p.requiresKey === false && !p.selfHosted;
  });
}

export const keylessBundleProvider: SearchProviderAdapter = {
  id: KEYLESS_BUNDLE_ID,
  label: 'All keyless sources (bundled)',
  requiresKey: false,
  // The bundle has no single endpoint or signup page; the UI hides both.
  endpoint: '',
  signupUrl: '',

  async search(query, opts): Promise<SearchResult[]> {
    const available = keylessSourceIds();
    const targeted = opts.source && (available as string[]).includes(opts.source)
      ? [opts.source as SearchProviderId]
      : AUTO_SOURCES.filter(id => available.includes(id));

    // Per-source base URL overrides make no sense across heterogeneous
    // endpoints, so the bundle never forwards one.
    const subOpts = { ...opts, baseUrl: undefined, source: undefined };
    const settled = await Promise.allSettled(
      targeted.map(async id => ({ id, results: await SEARCH_PROVIDERS[id].search(query, subOpts) }))
    );

    const succeeded = settled.filter(
      (s): s is PromiseFulfilledResult<{ id: SearchProviderId; results: SearchResult[] }> => s.status === 'fulfilled'
    );
    if (succeeded.length === 0) {
      // A single targeted source keeps its original error semantics.
      if (settled.length === 1 && settled[0].status === 'rejected') throw settled[0].reason;
      const details = settled
        .map((s, i) => `${targeted[i]}: ${s.status === 'rejected' ? messageOf(s.reason) : 'no results'}`)
        .join('; ');
      throw new ToolNetworkError(`All keyless sources failed - ${details}`);
    }

    return interleave(succeeded.map(({ value }) =>
      value.results.map(result => ({ ...result, source: value.id }))
    ));
  },
};

function messageOf(reason: unknown): string {
  return reason instanceof Error ? reason.message : String(reason);
}

/** Round-robin merge so every source is represented near the top of a capped result list. */
function interleave(buckets: SearchResult[][]): SearchResult[] {
  const merged: SearchResult[] = [];
  const longest = Math.max(0, ...buckets.map(b => b.length));
  for (let i = 0; i < longest; i++) {
    for (const bucket of buckets) {
      if (i < bucket.length) merged.push(bucket[i]);
    }
  }
  return merged;
}
