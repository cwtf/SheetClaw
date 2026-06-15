import { ToolNetworkError } from '../../workbook/executor';
import { resolveBaseUrl, type SearchProviderAdapter, type SearchResult } from './index';

interface JinaSearchResponse {
  data?: Array<{
    title?: unknown;
    url?: unknown;
    description?: unknown;
    date?: unknown;
  }>;
}

export const jinaProvider: SearchProviderAdapter = {
  id: 'jina',
  label: 'Jina Search',
  requiresKey: true,
  endpoint: 'https://s.jina.ai/',
  signupUrl: 'https://jina.ai/api-dashboard/',

  async search(query, opts): Promise<SearchResult[]> {
    const fetchImpl = opts.fetchImpl ?? fetch;
    const url = new URL(resolveBaseUrl(opts.baseUrl, this.endpoint));
    url.searchParams.set('q', query);

    let response: Response;
    try {
      response = await fetchImpl(url.toString(), {
        method: 'GET',
        headers: {
          'Authorization': `Bearer ${opts.apiKey}`,
          'Accept': 'application/json',
          // Skip full page content: links and descriptions are enough for search results.
          'X-Respond-With': 'no-content',
        },
        signal: opts.signal,
      });
    } catch (e) {
      const message = e instanceof Error ? e.message : String(e);
      throw new ToolNetworkError(`jina request failed: ${message}`);
    }

    if (!response.ok) {
      throw new ToolNetworkError(`jina request failed with HTTP ${response.status}`);
    }

    const contentType = response.headers.get('content-type') ?? '';
    if (contentType.includes('text/html')) {
      throw new ToolNetworkError(
        'Jina: web-access base URL looks misconfigured — response was HTML instead of JSON. Check your search provider settings.'
      );
    }

    let json: JinaSearchResponse;
    try {
      json = await response.json() as JinaSearchResponse;
    } catch {
      throw new ToolNetworkError('jina response was not valid JSON');
    }

    return (json.data ?? [])
      .map(item => normalizeResult(item))
      .filter((result): result is SearchResult => !!result);
  },
};

function normalizeResult(item: NonNullable<JinaSearchResponse['data']>[number]): SearchResult | null {
  if (typeof item.title !== 'string' || typeof item.url !== 'string') return null;
  return {
    title: item.title,
    url: item.url,
    ...(typeof item.description === 'string' ? { snippet: item.description } : {}),
    ...(typeof item.date === 'string' ? { publishedAt: item.date } : {}),
  };
}
