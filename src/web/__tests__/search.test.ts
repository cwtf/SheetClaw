import { describe, expect, it, vi } from 'vitest';
import { filterToolsForRun } from '../../agent/tool-filter';
import { resolveSearchToggle } from '../../adapters/native-search';
import type { ToolSpec } from '../../types';
import { createWebSearchHandler, WEB_SEARCH } from '../search';
import { MAX_RESULT_CONTENT_CHARS, resolveBaseUrl } from '../providers';
import { tavilyProvider } from '../providers/tavily';
import { googleCseProvider } from '../providers/google-cse';
import { jinaProvider } from '../providers/jina';
import { searxngProvider } from '../providers/searxng';
import { wikipediaProvider } from '../providers/wikipedia';
import { ToolExecutor, ToolValidationError } from '../../workbook/executor';
import { WorkbookRegistry } from '../../workbook/registry';
import type { ToolCall } from '../../types';

function jsonResponse(data: unknown, status = 200): Response {
  return new Response(JSON.stringify(data), {
    status,
    headers: { 'content-type': 'application/json' },
  });
}

function textResponse(body: string, contentType = 'text/html'): Response {
  return new Response(body, {
    status: 200,
    headers: { 'content-type': contentType },
  });
}

const DOMAIN_FIXTURES = [
  {
    title: 'Quarterly filing',
    url: 'https://finance.example/reports/q1',
    content: 'Revenue table and notes.',
  },
  {
    title: 'Hourly forecast',
    url: 'https://weather.example/api/hourly',
    content: 'Temperature and precipitation fields.',
  },
  {
    title: 'Match schedule',
    url: 'https://sports.example/schedule',
    content: 'Fixtures and scores.',
  },
];

describe('Tavily adapter', () => {
  it('parses documented JSON results from unrelated domains', async () => {
    const fetchImpl = vi.fn(async (_url: RequestInfo | URL, _init?: RequestInit) => jsonResponse({ results: DOMAIN_FIXTURES }));

    const results = await tavilyProvider.search('public data', {
      maxResults: 3,
      apiKey: 'key',
      signal: new AbortController().signal,
      fetchImpl,
    });

    expect(results).toEqual([
      { title: 'Quarterly filing', url: 'https://finance.example/reports/q1', snippet: 'Revenue table and notes.' },
      { title: 'Hourly forecast', url: 'https://weather.example/api/hourly', snippet: 'Temperature and precipitation fields.' },
      { title: 'Match schedule', url: 'https://sports.example/schedule', snippet: 'Fixtures and scores.' },
    ]);
  });

  it('preserves zero results as an empty array', async () => {
    const fetchImpl = vi.fn(async (_url: RequestInfo | URL, _init?: RequestInit) => jsonResponse({ results: [] }));

    const results = await tavilyProvider.search('no matches', {
      maxResults: 3,
      apiKey: 'key',
      signal: new AbortController().signal,
      fetchImpl,
    });

    expect(results).toEqual([]);
  });

  it('throws a clear config error when Tavily returns HTML instead of JSON', async () => {
    const fetchImpl = vi.fn(async (_url: RequestInfo | URL, _init?: RequestInit) =>
      textResponse('<html><title>Service unavailable</title></html>')
    );

    await expect(tavilyProvider.search('public data', {
      maxResults: 3,
      apiKey: 'key',
      signal: new AbortController().signal,
      fetchImpl,
    })).rejects.toThrow(/web-access base URL looks misconfigured.*HTML/);
  });

  it('keeps raw content disabled and omits the content field by default', async () => {
    const fetchImpl = vi.fn(async (_url: RequestInfo | URL, _init?: RequestInit) => jsonResponse({ results: DOMAIN_FIXTURES }));

    const results = await tavilyProvider.search('public data', {
      maxResults: 3,
      apiKey: 'key',
      signal: new AbortController().signal,
      fetchImpl,
    });

    const [, init] = fetchImpl.mock.calls[0];
    expect(JSON.parse((init as RequestInit).body as string)).toMatchObject({ include_raw_content: false });
    expect(results.every(r => r.content === undefined)).toBe(true);
  });

  it('requests text extraction and returns capped per-result content when includeContent is set', async () => {
    const longText = 'row data '.repeat(1000); // 9000 chars, well past the cap
    const fetchImpl = vi.fn(async (_url: RequestInfo | URL, _init?: RequestInit) => jsonResponse({
      results: [
        {
          title: 'Quarterly filing',
          url: 'https://finance.example/reports/q1',
          content: 'Revenue table and notes.',
          raw_content: longText,
        },
        {
          title: 'Match schedule',
          url: 'https://sports.example/schedule',
          content: 'Fixtures and scores.',
          raw_content: null,
        },
      ],
    }));

    const results = await tavilyProvider.search('public data', {
      maxResults: 2,
      apiKey: 'key',
      includeContent: true,
      signal: new AbortController().signal,
      fetchImpl,
    });

    const [, init] = fetchImpl.mock.calls[0];
    expect(JSON.parse((init as RequestInit).body as string)).toMatchObject({ include_raw_content: 'text' });
    expect(results[0].content!.startsWith(longText.slice(0, MAX_RESULT_CONTENT_CHARS))).toBe(true);
    expect(results[0].content!.endsWith('[truncated: page continues beyond this point]')).toBe(true);
    expect(results[0].snippet).toBe('Revenue table and notes.');
    expect(results[1].content).toBeUndefined();
    expect(results[1].snippet).toBe('Fixtures and scores.');
  });
});

describe('Google CSE adapter', () => {
  it('parses documented items, sends the key as a header, and caps num at 10', async () => {
    const fetchImpl = vi.fn(async (_url: RequestInfo | URL, _init?: RequestInit) => jsonResponse({
      items: DOMAIN_FIXTURES.map(f => ({ title: f.title, link: f.url, snippet: f.content })),
    }));

    const results = await googleCseProvider.search('public data', {
      maxResults: 10,
      apiKey: 'gkey',
      engineId: 'engine123',
      signal: new AbortController().signal,
      fetchImpl,
    });

    expect(results).toHaveLength(3);
    expect(results[0]).toEqual({
      title: 'Quarterly filing',
      url: 'https://finance.example/reports/q1',
      snippet: 'Revenue table and notes.',
    });
    const [url, init] = fetchImpl.mock.calls[0];
    expect(String(url)).toContain('cx=engine123');
    expect(String(url)).toContain('num=10');
    expect(String(url)).not.toContain('gkey');
    expect((init as RequestInit).headers).toMatchObject({ 'X-Goog-Api-Key': 'gkey' });
  });

  it('requires an engine id before issuing a request', async () => {
    const fetchImpl = vi.fn();
    await expect(googleCseProvider.search('public data', {
      maxResults: 5,
      apiKey: 'gkey',
      signal: new AbortController().signal,
      fetchImpl,
    })).rejects.toBeInstanceOf(ToolValidationError);
    expect(fetchImpl).not.toHaveBeenCalled();
  });
});

describe('Jina adapter', () => {
  it('parses documented data items and authenticates with a bearer header', async () => {
    const fetchImpl = vi.fn(async (_url: RequestInfo | URL, _init?: RequestInit) => jsonResponse({
      data: DOMAIN_FIXTURES.map(f => ({ title: f.title, url: f.url, description: f.content })),
    }));

    const results = await jinaProvider.search('public data', {
      maxResults: 3,
      apiKey: 'jkey',
      signal: new AbortController().signal,
      fetchImpl,
    });

    expect(results.map(r => r.url)).toEqual(DOMAIN_FIXTURES.map(f => f.url));
    const [, init] = fetchImpl.mock.calls[0];
    expect((init as RequestInit).headers).toMatchObject({ 'Authorization': 'Bearer jkey' });
  });
});

describe('SearXNG adapter', () => {
  it('parses documented results keylessly from a user-supplied instance URL', async () => {
    const fetchImpl = vi.fn(async (_url: RequestInfo | URL, _init?: RequestInit) => jsonResponse({
      results: DOMAIN_FIXTURES.map(f => ({ title: f.title, url: f.url, content: f.content })),
    }));

    const results = await searxngProvider.search('public data', {
      maxResults: 3,
      apiKey: '',
      baseUrl: 'http://localhost:8888/search',
      signal: new AbortController().signal,
      fetchImpl,
    });

    expect(results).toHaveLength(3);
    const [url] = fetchImpl.mock.calls[0];
    expect(String(url)).toContain('http://localhost:8888/search');
    expect(String(url)).toContain('format=json');
  });
});

describe('Wikipedia adapter', () => {
  it('parses documented search hits keylessly, building article URLs and stripping snippet HTML', async () => {
    const fetchImpl = vi.fn(async (_url: RequestInfo | URL, _init?: RequestInit) => jsonResponse({
      query: {
        search: [
          { title: 'Compound interest', snippet: 'the addition of <span class="searchmatch">interest</span> to principal' },
          { title: 'Gross domestic product', snippet: 'monetary measure of market value' },
        ],
      },
    }));

    const results = await wikipediaProvider.search('public data', {
      maxResults: 2,
      apiKey: '',
      signal: new AbortController().signal,
      fetchImpl,
    });

    expect(results).toEqual([
      {
        title: 'Compound interest',
        url: 'https://en.wikipedia.org/wiki/Compound_interest',
        snippet: 'the addition of interest to principal',
      },
      {
        title: 'Gross domestic product',
        url: 'https://en.wikipedia.org/wiki/Gross_domestic_product',
        snippet: 'monetary measure of market value',
      },
    ]);
    const [url] = fetchImpl.mock.calls[0];
    expect(String(url)).toContain('origin=*');
  });

  it('fetches intro extracts for all results when includeContent is set', async () => {
    const fetchImpl = vi.fn(async (url: RequestInfo | URL, _init?: RequestInit) => {
      if (String(url).includes('prop=extracts')) {
        return jsonResponse({
          query: {
            pages: {
              '1234': { title: 'Compound interest', extract: 'Compound interest is the addition of interest to principal.' },
            },
          },
        });
      }
      return jsonResponse({
        query: {
          search: [{ title: 'Compound interest', snippet: 'the addition of <span>interest</span> to principal' }],
        },
      });
    });

    const results = await wikipediaProvider.search('compound interest', {
      maxResults: 1,
      apiKey: '',
      includeContent: true,
      signal: new AbortController().signal,
      fetchImpl,
    });

    expect(results).toHaveLength(1);
    expect(results[0].content).toBe('Compound interest is the addition of interest to principal.');
    expect(results[0].snippet).toBe('the addition of interest to principal');
    expect(fetchImpl).toHaveBeenCalledTimes(2);
    const [extractsUrl] = fetchImpl.mock.calls[1];
    expect(String(extractsUrl)).toContain('prop=extracts');
    expect(String(extractsUrl)).toContain('explaintext=1');
    expect(String(extractsUrl)).toContain('exintro=1');
    expect(String(extractsUrl)).toContain('origin=*');
    expect(String(extractsUrl)).toContain('Compound+interest');
  });

  it('caps extract content at MAX_RESULT_CONTENT_CHARS', async () => {
    const longExtract = 'word '.repeat(2000);
    const fetchImpl = vi.fn(async (url: RequestInfo | URL, _init?: RequestInit) => {
      if (String(url).includes('prop=extracts')) {
        return jsonResponse({
          query: { pages: { '1': { title: 'Compound interest', extract: longExtract } } },
        });
      }
      return jsonResponse({
        query: { search: [{ title: 'Compound interest', snippet: 'snippet' }] },
      });
    });

    const results = await wikipediaProvider.search('compound interest', {
      maxResults: 1,
      apiKey: '',
      includeContent: true,
      signal: new AbortController().signal,
      fetchImpl,
    });

    expect(results[0].content!.startsWith(longExtract.slice(0, MAX_RESULT_CONTENT_CHARS))).toBe(true);
    expect(results[0].content!.endsWith('[truncated: article continues beyond this point]')).toBe(true);
  });

  it('falls back to snippets only if the extracts call fails', async () => {
    let callCount = 0;
    const fetchImpl = vi.fn(async (_url: RequestInfo | URL, _init?: RequestInit) => {
      callCount++;
      if (callCount === 2) throw new Error('network error');
      return jsonResponse({
        query: { search: [{ title: 'Compound interest', snippet: 'the addition of interest to principal' }] },
      });
    });

    const results = await wikipediaProvider.search('compound interest', {
      maxResults: 1,
      apiKey: '',
      includeContent: true,
      signal: new AbortController().signal,
      fetchImpl,
    });

    expect(results).toHaveLength(1);
    expect(results[0].content).toBeUndefined();
    expect(results[0].snippet).toBe('the addition of interest to principal');
  });

  it('skips the extracts call when includeContent is not set', async () => {
    const fetchImpl = vi.fn(async (_url: RequestInfo | URL, _init?: RequestInit) => jsonResponse({
      query: { search: [{ title: 'Compound interest', snippet: 'the addition of interest to principal' }] },
    }));

    const results = await wikipediaProvider.search('compound interest', {
      maxResults: 1,
      apiKey: '',
      signal: new AbortController().signal,
      fetchImpl,
    });

    expect(results[0].content).toBeUndefined();
    expect(fetchImpl).toHaveBeenCalledTimes(1);
  });
});

describe('resolveBaseUrl', () => {
  const FALLBACK = 'https://finance.example';

  it('returns fallback for undefined', () => {
    expect(resolveBaseUrl(undefined, FALLBACK)).toBe(FALLBACK);
  });

  it('returns fallback for empty string', () => {
    expect(resolveBaseUrl('', FALLBACK)).toBe(FALLBACK);
  });

  it('returns fallback for a relative path', () => {
    expect(resolveBaseUrl('/api/search', FALLBACK)).toBe(FALLBACK);
  });

  it('returns fallback for a protocol-relative URL', () => {
    expect(resolveBaseUrl('//host/path', FALLBACK)).toBe(FALLBACK);
  });

  it('returns fallback for a non-http scheme', () => {
    expect(resolveBaseUrl('data:text/html,hi', FALLBACK)).toBe(FALLBACK);
  });

  it('returns the provided URL when it is a valid http URL', () => {
    expect(resolveBaseUrl('http://public.example/search', FALLBACK))
      .toBe('http://public.example/search');
  });

  it('returns the provided URL when it is a valid https URL', () => {
    expect(resolveBaseUrl('https://weather.example/api', FALLBACK))
      .toBe('https://weather.example/api');
  });
});

describe('base URL guard — empty/relative baseUrl falls back to provider endpoint', () => {
  it('Tavily: empty baseUrl sends request to the real Tavily endpoint', async () => {
    const fetchImpl = vi.fn(async (_url: RequestInfo | URL, _init?: RequestInit) =>
      jsonResponse({ results: [] })
    );
    await tavilyProvider.search('test', {
      maxResults: 1, apiKey: 'key', baseUrl: '', signal: new AbortController().signal, fetchImpl,
    });
    const [url] = fetchImpl.mock.calls[0];
    expect(String(url)).toBe('https://api.tavily.com/search');
  });

  it('SearXNG: relative baseUrl falls back to endpoint; valid custom URL is preserved', async () => {
    const fetchImpl = vi.fn(async (_url: RequestInfo | URL, _init?: RequestInit) =>
      jsonResponse({ results: [] })
    );
    // Relative URL → falls back to default endpoint
    await searxngProvider.search('test', {
      maxResults: 1, apiKey: '', baseUrl: '/search', signal: new AbortController().signal, fetchImpl,
    });
    expect(String(fetchImpl.mock.calls[0][0])).toContain('localhost:8080');

    fetchImpl.mockClear();
    // Valid absolute custom URL → preserved
    await searxngProvider.search('test', {
      maxResults: 1, apiKey: '', baseUrl: 'http://localhost:9090/search',
      signal: new AbortController().signal, fetchImpl,
    });
    expect(String(fetchImpl.mock.calls[0][0])).toContain('localhost:9090');
  });

  it('providers throw a clear config error when the endpoint returns HTML', async () => {
    const htmlResp = textResponse('<html><body>login required</body></html>');
    for (const [provider, opts] of [
      [tavilyProvider, { apiKey: 'k' }] as const,
      [searxngProvider, { apiKey: '' }] as const,
      [jinaProvider, { apiKey: 'k' }] as const,
    ] as const) {
      const fetchImpl = vi.fn(async () => htmlResp.clone());
      await expect(
        provider.search('q', { maxResults: 1, ...opts, signal: new AbortController().signal, fetchImpl })
      ).rejects.toThrow(/web-access base URL looks misconfigured.*HTML/);
    }
  });
});

describe('web_search handler', () => {
  it('returns NetworkError through the executor path without fabricated result URLs', async () => {
    const executor = new ToolExecutor(new WorkbookRegistry(), async () => {
      throw new Error('runtime none should not use Excel runner');
    });
    executor.register(WEB_SEARCH, createWebSearchHandler({
      getProvider: () => 'tavily',
      getApiKey: () => 'key',
      fetchImpl: vi.fn(async () => { throw new TypeError('network down'); }),
    }));

    const call: ToolCall = {
      id: 'search_call',
      name: 'web_search',
      arguments: { query: 'anything' },
      workbookId: 'host',
      mutating: false,
    };
    const result = await executor.execute(call, { workbookId: 'host' });

    expect(result.ok).toBe(false);
    expect(result.error?.code).toBe('NetworkError');
    expect(JSON.stringify(result)).not.toMatch(/https?:\/\//);
  });

  it('plumbs include_content through to the provider and surfaces content in the result', async () => {
    const fetchImpl = vi.fn(async (_url: RequestInfo | URL, _init?: RequestInit) => jsonResponse({
      results: [
        {
          title: 'Quarterly filing',
          url: 'https://finance.example/reports/q1',
          content: 'Revenue table and notes.',
          raw_content: 'Full extracted report text.',
        },
      ],
    }));
    const executor = new ToolExecutor(new WorkbookRegistry(), async () => {
      throw new Error('runtime none should not use Excel runner');
    });
    executor.register(WEB_SEARCH, createWebSearchHandler({
      getProvider: () => 'tavily',
      getApiKey: () => 'key',
      fetchImpl,
    }));

    const call: ToolCall = {
      id: 'search_call',
      name: 'web_search',
      arguments: { query: 'public data', include_content: true },
      workbookId: 'host',
      mutating: false,
    };
    const result = await executor.execute(call, { workbookId: 'host' });

    expect(result.ok).toBe(true);
    const [, init] = fetchImpl.mock.calls[0];
    expect(JSON.parse((init as RequestInit).body as string)).toMatchObject({ include_raw_content: 'text' });
    expect((result.data as { results: Array<{ content?: string }> }).results[0].content)
      .toBe('Full extracted report text.');
  });

  it('rejects a non-boolean include_content with a ValidationError', async () => {
    const fetchImpl = vi.fn();
    const executor = new ToolExecutor(new WorkbookRegistry(), async () => {
      throw new Error('runtime none should not use Excel runner');
    });
    executor.register(WEB_SEARCH, createWebSearchHandler({
      getProvider: () => 'tavily',
      getApiKey: () => 'key',
      fetchImpl,
    }));

    const call: ToolCall = {
      id: 'search_call',
      name: 'web_search',
      arguments: { query: 'public data', include_content: 'yes' },
      workbookId: 'host',
      mutating: false,
    };
    const result = await executor.execute(call, { workbookId: 'host' });

    expect(result.ok).toBe(false);
    expect(result.error?.code).toBe('ValidationError');
    expect(fetchImpl).not.toHaveBeenCalled();
  });
});

describe('web tool exposure gating', () => {
  const tools: ToolSpec[] = [
    { name: 'read_range', description: '', parameters: { type: 'object', properties: {} }, mutating: false },
    WEB_SEARCH,
    { name: 'fetch_url', description: '', parameters: { type: 'object', properties: {} }, mutating: false, runtime: 'none' },
  ];

  it('toggle off removes web tools from the LLM request', () => {
    const search = resolveSearchToggle({ provider: 'ollama', model: 'llama3.2', byokReady: true });
    expect(filterToolsForRun(tools, false, search).map(t => t.name)).toEqual(['read_range']);
  });

  it('missing provider key removes web tools even if toggle is on', () => {
    const search = resolveSearchToggle({ provider: 'ollama', model: 'llama3.2', byokReady: false });
    expect(filterToolsForRun(tools, true, search).map(t => t.name)).toEqual(['read_range']);
  });

  it('BYOK tier with configured provider and toggle on exposes both web tools', () => {
    const search = resolveSearchToggle({ provider: 'ollama', model: 'llama3.2', byokReady: true });
    expect(filterToolsForRun(tools, true, search).map(t => t.name)).toEqual(['read_range', 'web_search', 'fetch_url']);
  });

  it('native tier with toggle on suppresses client web_search but exposes fetch_url', () => {
    const search = resolveSearchToggle({ provider: 'generic', model: 'openai/gpt-4o-mini', byokReady: true });
    expect(filterToolsForRun(tools, true, search).map(t => t.name)).toEqual(['read_range', 'fetch_url']);
  });

  it('Qwen unsupported models fall back to BYOK gating', () => {
    const unavailable = resolveSearchToggle({ provider: 'qwen', model: 'qwen-plus', byokReady: false });
    const configured = resolveSearchToggle({ provider: 'qwen', model: 'qwen-plus', byokReady: true });

    expect(filterToolsForRun(tools, true, unavailable).map(t => t.name)).toEqual(['read_range']);
    expect(filterToolsForRun(tools, true, configured).map(t => t.name)).toEqual(['read_range', 'web_search', 'fetch_url']);
  });
});
