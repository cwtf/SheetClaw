import { describe, expect, it, vi } from 'vitest';
import { createWebSearchHandler, WEB_SEARCH } from '../search';
import { keylessBundleProvider, keylessSourceIds, KEYLESS_BUNDLE_ID } from '../providers';
import { ToolNetworkError, ToolValidationError } from '../../workbook/executor';

function jsonResponse(data: unknown, status = 200): Response {
  return new Response(JSON.stringify(data), {
    status,
    headers: { 'content-type': 'application/json' },
  });
}

/**
 * Serves shape-correct payloads for the bundle's general sources, routed by
 * endpoint URL, so fan-out behaviour is exercised through the real adapters.
 */
function makeRoutedFetch(overrides: Partial<Record<'wikipedia' | 'wikidata' | 'worldbank', () => Response | Promise<Response>>> = {}) {
  return vi.fn(async (input: RequestInfo | URL) => {
    const url = String(input);
    if (url.includes('en.wikipedia.org')) {
      return overrides.wikipedia?.() ?? jsonResponse({
        query: { search: [{ title: 'GDP of Malaysia', snippet: 'economy article' }] },
      });
    }
    if (url.includes('wikidata.org')) {
      return overrides.wikidata?.() ?? jsonResponse({
        search: [{ id: 'Q833', title: 'Q833', label: 'Malaysia', description: 'country', concepturi: 'https://www.wikidata.org/wiki/Q833' }],
      });
    }
    if (url.includes('worldbank.org')) {
      return overrides.worldbank?.() ?? jsonResponse([
        { page: 1, pages: 1, per_page: '20000', total: 1 },
        [{ id: 'NY.GDP.PCAP.CD', name: 'GDP per capita (current US$)', sourceNote: 'gdp per person', topics: [] }],
      ]);
    }
    throw new Error(`unexpected fetch in test: ${url}`);
  });
}

describe('keyless bundle provider', () => {
  it('lists every keyless, non-self-hosted source and excludes searxng', () => {
    const ids = keylessSourceIds();
    expect(ids).toContain('wikipedia');
    expect(ids).toContain('worldbank');
    expect(ids).toContain('imf');
    expect(ids).not.toContain('searxng');
    expect(ids).not.toContain('tavily');
  });

  it('auto mode fans out to the general sources and tags results with their source', async () => {
    const fetchImpl = makeRoutedFetch();
    const results = await keylessBundleProvider.search('malaysia gdp', {
      maxResults: 6,
      apiKey: '',
      signal: new AbortController().signal,
      fetchImpl,
    });

    const sources = new Set(results.map(r => r.source));
    expect(sources).toEqual(new Set(['wikipedia', 'wikidata', 'worldbank']));
    // Round-robin interleave puts one result from each source first.
    expect(results.slice(0, 3).map(r => r.source).sort()).toEqual(['wikidata', 'wikipedia', 'worldbank']);
  });

  it('a targeted source queries only that source', async () => {
    const fetchImpl = makeRoutedFetch();
    const results = await keylessBundleProvider.search('gdp per capita', {
      maxResults: 5,
      apiKey: '',
      source: 'worldbank',
      signal: new AbortController().signal,
      fetchImpl,
    });

    expect(results.every(r => r.source === 'worldbank')).toBe(true);
    expect(fetchImpl.mock.calls.every(([url]) => String(url).includes('worldbank.org'))).toBe(true);
  });

  it('ignores a per-provider base URL override when fanning out', async () => {
    const fetchImpl = makeRoutedFetch();
    await keylessBundleProvider.search('malaysia', {
      maxResults: 3,
      apiKey: '',
      baseUrl: 'https://public.example/api',
      signal: new AbortController().signal,
      fetchImpl,
    });

    expect(fetchImpl.mock.calls.some(([url]) => String(url).includes('public.example'))).toBe(false);
  });

  it('tolerates individual source failures and returns the remaining results', async () => {
    const fetchImpl = makeRoutedFetch({
      worldbank: () => { throw new TypeError('Failed to fetch'); },
    });
    const results = await keylessBundleProvider.search('malaysia gdp', {
      maxResults: 6,
      apiKey: '',
      signal: new AbortController().signal,
      fetchImpl,
    });

    const sources = new Set(results.map(r => r.source));
    expect(sources).toEqual(new Set(['wikipedia', 'wikidata']));
  });

  it('throws a combined ToolNetworkError when every source fails', async () => {
    const boom = () => { throw new TypeError('Failed to fetch'); };
    const fetchImpl = makeRoutedFetch({ wikipedia: boom, wikidata: boom, worldbank: boom });

    await expect(keylessBundleProvider.search('malaysia gdp', {
      maxResults: 3,
      apiKey: '',
      signal: new AbortController().signal,
      fetchImpl,
    })).rejects.toThrow(/All keyless sources failed.*wikipedia.*wikidata.*worldbank/s);
  });

  it('preserves the original error when a single targeted source fails', async () => {
    const fetchImpl = makeRoutedFetch({
      wikipedia: () => { throw new TypeError('Failed to fetch'); },
    });

    await expect(keylessBundleProvider.search('malaysia gdp', {
      maxResults: 3,
      apiKey: '',
      source: 'wikipedia',
      signal: new AbortController().signal,
      fetchImpl,
    })).rejects.toThrow(/wikipedia request failed/);
  });
});

describe('web_search handler with the keyless bundle', () => {
  function makeHandler(fetchImpl: typeof fetch) {
    const handler = createWebSearchHandler({
      getProvider: () => KEYLESS_BUNDLE_ID,
      getApiKey: () => '',
      fetchImpl,
    });
    // runtime:'none' handlers never touch the Excel context or registry.
    return (args: Record<string, unknown>) =>
      handler(args, undefined as unknown as Excel.RequestContext, undefined as unknown as never);
  }

  it('routes the source argument through to the bundle', async () => {
    const fetchImpl = makeRoutedFetch();
    const handler = makeHandler(fetchImpl as unknown as typeof fetch);

    const result = await handler({ query: 'gdp per capita', source: 'worldbank' }) as {
      provider: string;
      results: Array<{ source?: string }>;
    };

    expect(result.provider).toBe(KEYLESS_BUNDLE_ID);
    expect(result.results.every(r => r.source === 'worldbank')).toBe(true);
  });

  it("treats source 'auto' as the fan-out default", async () => {
    const fetchImpl = makeRoutedFetch();
    const handler = makeHandler(fetchImpl as unknown as typeof fetch);

    const result = await handler({ query: 'malaysia gdp', source: 'auto', max_results: 6 }) as {
      results: Array<{ source?: string }>;
    };

    expect(new Set(result.results.map(r => r.source))).toEqual(new Set(['wikipedia', 'wikidata', 'worldbank']));
  });

  it('rejects an unknown source with a correctable ValidationError', async () => {
    const fetchImpl = vi.fn();
    const handler = makeHandler(fetchImpl as unknown as typeof fetch);

    await expect(handler({ query: 'gdp', source: 'bing' })).rejects.toThrow(ToolValidationError);
    await expect(handler({ query: 'gdp', source: 'bing' })).rejects.toThrow(/must be 'auto' or one of/);
    expect(fetchImpl).not.toHaveBeenCalled();
  });

  it('surfaces an all-sources failure as a ToolNetworkError', async () => {
    const boom = () => { throw new TypeError('Failed to fetch'); };
    const fetchImpl = makeRoutedFetch({ wikipedia: boom, wikidata: boom, worldbank: boom });
    const handler = makeHandler(fetchImpl as unknown as typeof fetch);

    await expect(handler({ query: 'malaysia gdp' })).rejects.toThrow(ToolNetworkError);
  });
});

describe('web_search tool spec', () => {
  it('documents every bundle source in the source enum', () => {
    const sourceParam = (WEB_SEARCH.parameters as unknown as {
      properties: { source: { enum: string[]; description: string } };
    }).properties.source;

    expect(sourceParam.enum).toEqual(['auto', ...keylessSourceIds()]);
    expect(sourceParam.description).toContain('wikipedia');
    expect(sourceParam.description).toContain('open-meteo');
  });
});
