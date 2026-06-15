import { ToolNetworkError } from '../../workbook/executor';
import { resolveBaseUrl, type SearchProviderAdapter, type SearchResult } from './index';

interface SearxngResponse {
  results?: Array<{
    title?: unknown;
    url?: unknown;
    content?: unknown;
    publishedDate?: unknown;
  }>;
}

export const searxngProvider: SearchProviderAdapter = {
  id: 'searxng',
  label: 'SearXNG (self-hosted)',
  requiresKey: false,
  endpoint: 'http://localhost:8080/search',
  signupUrl: 'https://docs.searxng.org/',

  async search(query, opts): Promise<SearchResult[]> {
    const fetchImpl = opts.fetchImpl ?? fetch;
    const url = new URL(resolveBaseUrl(opts.baseUrl, this.endpoint));
    url.searchParams.set('q', query);
    url.searchParams.set('format', 'json');

    let response: Response;
    try {
      response = await fetchImpl(url.toString(), {
        method: 'GET',
        headers: { 'Accept': 'application/json' },
        signal: opts.signal,
      });
    } catch (e) {
      const message = e instanceof Error ? e.message : String(e);
      throw new ToolNetworkError(`searxng request failed: ${message}. Confirm the instance URL in Settings and that its JSON format and CORS settings are enabled.`);
    }

    if (!response.ok) {
      throw new ToolNetworkError(`searxng request failed with HTTP ${response.status}. The instance may have format=json disabled.`);
    }

    const contentType = response.headers.get('content-type') ?? '';
    if (contentType.includes('text/html')) {
      throw new ToolNetworkError(
        'SearXNG: web-access base URL looks misconfigured — response was HTML instead of JSON. Check your search provider settings.'
      );
    }

    let json: SearxngResponse;
    try {
      json = await response.json() as SearxngResponse;
    } catch {
      throw new ToolNetworkError('searxng response was not valid JSON');
    }

    return (json.results ?? [])
      .map(item => normalizeResult(item))
      .filter((result): result is SearchResult => !!result);
  },
};

function normalizeResult(item: NonNullable<SearxngResponse['results']>[number]): SearchResult | null {
  if (typeof item.title !== 'string' || typeof item.url !== 'string') return null;
  return {
    title: item.title,
    url: item.url,
    ...(typeof item.content === 'string' ? { snippet: item.content } : {}),
    ...(typeof item.publishedDate === 'string' ? { publishedAt: item.publishedDate } : {}),
  };
}
