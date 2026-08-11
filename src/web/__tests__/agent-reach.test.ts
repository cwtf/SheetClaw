import { describe, expect, it, vi } from 'vitest';
import { handleFetchUrlWithOptions } from '../fetch';
import { agentReachPlatformFor, readViaAgentReach } from '../agent-reach';
import { agentReachProvider, splitPlatformPrefix, getSearchProvider, keylessSourceIds } from '../providers';

const BRIDGE = 'http://localhost:8788';

function jsonResponse(body: unknown, status = 200): Response {
  return new Response(JSON.stringify(body), {
    status,
    headers: { 'content-type': 'application/json' },
  });
}

describe('platform routing', () => {
  it('maps the hosts the bridge can read, ignoring www', () => {
    expect(agentReachPlatformFor('https://x.com/someone/status/1')).toBe('twitter');
    expect(agentReachPlatformFor('https://www.reddit.com/r/excel/comments/abc')).toBe('reddit');
    expect(agentReachPlatformFor('https://youtu.be/dQw4w9WgXcQ')).toBe('youtube');
    expect(agentReachPlatformFor('https://github.com/vitest-dev/vitest')).toBe('github');
  });

  it('returns null for ordinary sites and for junk input', () => {
    expect(agentReachPlatformFor('https://public.example/page')).toBeNull();
    expect(agentReachPlatformFor('not a url')).toBeNull();
  });
});

describe('platform prefix parsing', () => {
  it('pulls a known platform off the front of the query', () => {
    expect(splitPlatformPrefix('youtube: excel pivot tables')).toEqual({
      platform: 'youtube',
      query: 'excel pivot tables',
    });
  });

  it('leaves an unknown prefix alone so ordinary queries with colons survive', () => {
    expect(splitPlatformPrefix('excel: how to sum')).toEqual({ query: 'excel: how to sum' });
    expect(splitPlatformPrefix('  plain query  ')).toEqual({ query: 'plain query' });
  });
});

describe('agent-reach search provider', () => {
  it('is registered, keyless, and excluded from the keyless bundle', () => {
    expect(getSearchProvider('agent-reach')).toBe(agentReachProvider);
    expect(agentReachProvider.requiresKey).toBe(false);
    // selfHosted keeps a possibly-stopped localhost service out of the default fan-out.
    expect(agentReachProvider.selfHosted).toBe(true);
    expect(keylessSourceIds()).not.toContain('agent-reach');
  });

  it('forwards the platform prefix and normalizes results', async () => {
    const seen: string[] = [];
    const fetchImpl = vi.fn(async (url: RequestInfo | URL) => {
      seen.push(String(url));
      return jsonResponse({
        query: 'pivot tables',
        platform: 'youtube',
        results: [
          { title: 'Pivot tables', url: 'https://youtu.be/abc', snippet: 'a guide' },
          { title: 'no url', url: '' },
        ],
      });
    });

    const results = await agentReachProvider.search('youtube: pivot tables', {
      maxResults: 5,
      apiKey: '',
      baseUrl: BRIDGE,
      signal: new AbortController().signal,
      fetchImpl: fetchImpl as unknown as typeof fetch,
    });

    expect(seen[0]).toContain('/search');
    expect(seen[0]).toContain('platform=youtube');
    expect(seen[0]).toContain('q=pivot+tables');
    // The entry with no URL is dropped rather than handed to the model.
    expect(results).toEqual([
      { title: 'Pivot tables', url: 'https://youtu.be/abc', snippet: 'a guide' },
    ]);
  });

  it('turns a 501 from an undispatched platform into actionable guidance', async () => {
    const fetchImpl = vi.fn(async () =>
      jsonResponse({ error: 'Searching reddit is routed but not dispatched by this bridge.' }, 501)
    );

    await expect(agentReachProvider.search('reddit: excel', {
      maxResults: 5,
      apiKey: '',
      baseUrl: BRIDGE,
      signal: new AbortController().signal,
      fetchImpl: fetchImpl as unknown as typeof fetch,
    })).rejects.toThrow(/not dispatched.*Prefix the query/s);
  });

  it('names the npm script when the bridge is not running', async () => {
    const fetchImpl = vi.fn(async () => {
      throw new TypeError('Failed to fetch');
    });

    await expect(agentReachProvider.search('youtube: excel', {
      maxResults: 5,
      apiKey: '',
      baseUrl: BRIDGE,
      signal: new AbortController().signal,
      fetchImpl: fetchImpl as unknown as typeof fetch,
    })).rejects.toThrow(/npm run agent-reach-bridge/);
  });
});

describe('readViaAgentReach', () => {
  it('requests /read and returns the normalized payload', async () => {
    const seen: string[] = [];
    const fetchImpl = vi.fn(async (url: RequestInfo | URL) => {
      seen.push(String(url));
      return jsonResponse({
        url: 'https://x.com/a/status/1',
        platform: 'twitter',
        tool: 'twitter-backend',
        title: 'a post',
        text: 'post body',
        truncated: false,
      });
    });

    const read = await readViaAgentReach('https://x.com/a/status/1', {
      baseUrl: BRIDGE,
      fetchImpl: fetchImpl as unknown as typeof fetch,
    });

    expect(seen[0]).toContain('/read');
    expect(seen[0]).toContain(encodeURIComponent('https://x.com/a/status/1'));
    expect(read).toMatchObject({ platform: 'twitter', tool: 'twitter-backend', text: 'post body' });
  });

  it('surfaces the bridge error body on a non-OK status', async () => {
    const fetchImpl = vi.fn(async () => jsonResponse({ error: 'rdt is not on PATH.' }, 501));

    await expect(readViaAgentReach('https://reddit.com/r/x', {
      baseUrl: BRIDGE,
      fetchImpl: fetchImpl as unknown as typeof fetch,
    })).rejects.toThrow(/rdt is not on PATH/);
  });
});

describe('fetch_url platform backend', () => {
  it('routes a platform URL through the bridge instead of fetching it directly', async () => {
    const fetchImpl = vi.fn(async (url: RequestInfo | URL) => {
      if (String(url).includes('/read')) {
        return jsonResponse({ platform: 'youtube', tool: 'yt-dlp', title: 'v', text: 'transcript', truncated: false });
      }
      throw new Error('direct fetch should not be attempted for a platform URL');
    });

    const result = await handleFetchUrlWithOptions(
      { url: 'https://www.youtube.com/watch?v=abc' },
      { fetchImpl: fetchImpl as unknown as typeof fetch, agentReachBaseUrl: BRIDGE }
    );

    expect(result).toMatchObject({
      source: 'agent-reach',
      platform: 'youtube',
      tool: 'yt-dlp',
      text: 'transcript',
    });
  });

  it('falls back to the normal path when the bridge is down', async () => {
    const fetchImpl = vi.fn(async (url: RequestInfo | URL) => {
      if (String(url).includes('/read')) throw new TypeError('Failed to fetch');
      const response = new Response('<html><body>direct body</body></html>', {
        status: 200,
        headers: { 'content-type': 'text/html' },
      });
      Object.defineProperty(response, 'url', { value: 'https://www.youtube.com/watch?v=abc' });
      return response;
    });

    const result = await handleFetchUrlWithOptions(
      { url: 'https://www.youtube.com/watch?v=abc' },
      { fetchImpl: fetchImpl as unknown as typeof fetch, agentReachBaseUrl: BRIDGE }
    );

    expect(result).toMatchObject({ source: 'direct' });
    expect((result as { text: string }).text).toContain('direct body');
  });

  it('never calls the bridge when no bridge URL is configured', async () => {
    const seen: string[] = [];
    const fetchImpl = vi.fn(async (url: RequestInfo | URL) => {
      seen.push(String(url));
      const response = new Response('plain', { status: 200, headers: { 'content-type': 'text/plain' } });
      Object.defineProperty(response, 'url', { value: 'https://x.com/a/status/1' });
      return response;
    });

    await handleFetchUrlWithOptions(
      { url: 'https://x.com/a/status/1' },
      { fetchImpl: fetchImpl as unknown as typeof fetch }
    );

    expect(seen.some(u => u.includes('/read'))).toBe(false);
  });

  it('leaves non-platform URLs on the direct path even with the bridge configured', async () => {
    const seen: string[] = [];
    const fetchImpl = vi.fn(async (url: RequestInfo | URL) => {
      seen.push(String(url));
      const response = new Response('plain', { status: 200, headers: { 'content-type': 'text/plain' } });
      Object.defineProperty(response, 'url', { value: 'https://public.example/page' });
      return response;
    });

    const result = await handleFetchUrlWithOptions(
      { url: 'https://public.example/page' },
      { fetchImpl: fetchImpl as unknown as typeof fetch, agentReachBaseUrl: BRIDGE }
    );

    expect(seen.some(u => u.includes('/read'))).toBe(false);
    expect(result).toMatchObject({ source: 'direct' });
  });
});
