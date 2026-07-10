import { describe, expect, it } from 'vitest';
import { SEARCH_PROVIDERS, type SearchProviderId } from '../providers';

/**
 * Live-network CORS regression check for the keyless search providers.
 *
 * The Vitest suite runs in Node, which never enforces CORS, so a provider whose
 * endpoint stops sending Access-Control-Allow-Origin keeps passing unit tests
 * while every fetch from the Excel task pane throws TypeError. This check runs
 * each keyless provider against its real endpoint through a fetch wrapper that
 * enforces CORS the way a browser at the task-pane origin would, so it fails
 * exactly when the add-in would fail at runtime.
 *
 * Run with: npm run test:providers (excluded from the default `npm test`).
 */

const TASK_PANE_ORIGIN = 'https://localhost:3000';

/** Queries that make each provider issue a realistic request; open-meteo only fetches when coordinates are present. */
const PROVIDER_QUERIES: Partial<Record<SearchProviderId, string>> = {
  wikipedia: 'water quality',
  wikidata: 'gross domestic product',
  worldbank: 'gdp',
  ckan: 'water',
  'data-gov-sg': 'rainfall',
  'data-gov-my': 'weather forecast',
  ecb: 'exchange rates',
  eurostat: 'gdp',
  imf: 'real gdp growth',
  'open-meteo': '1.3521, 103.8198 rainfall',
  'un-sdg': 'poverty',
};

/**
 * Performs a real fetch but rejects with TypeError when the response lacks an
 * Access-Control-Allow-Origin header covering the task-pane origin — the same
 * observable behavior the WebView produces, so provider fallback paths (e.g.
 * IMF's reader proxy) are exercised exactly as they would be in Excel.
 */
const browserLikeFetch: typeof fetch = async (input, init) => {
  const headers = new Headers(init?.headers);
  headers.set('Origin', TASK_PANE_ORIGIN);
  const response = await fetch(input, { ...init, headers });
  const allowOrigin = response.headers.get('access-control-allow-origin');
  if (allowOrigin !== '*' && allowOrigin !== TASK_PANE_ORIGIN) {
    const url = typeof input === 'string' ? input : input instanceof URL ? input.href : input.url;
    throw new TypeError(`Failed to fetch: ${new URL(url).hostname} sent no usable Access-Control-Allow-Origin header (got ${JSON.stringify(allowOrigin)})`);
  }
  return response;
};

describe('keyless provider endpoints allow task-pane CORS', () => {
  // searxng is keyless but self-hosted (localhost endpoint), so there is no public instance to probe.
  const keylessProviders = Object.values(SEARCH_PROVIDERS)
    .filter(provider => !provider.requiresKey && provider.id !== 'searxng');

  for (const provider of keylessProviders) {
    it(`${provider.id} search succeeds under browser-style CORS enforcement`, async () => {
      const results = await provider.search(PROVIDER_QUERIES[provider.id] ?? 'water', {
        maxResults: 3,
        apiKey: '',
        signal: new AbortController().signal,
        fetchImpl: browserLikeFetch,
      });
      expect(Array.isArray(results)).toBe(true);
    });
  }
});
