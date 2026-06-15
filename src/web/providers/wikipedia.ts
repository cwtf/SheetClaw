import { ToolNetworkError } from '../../workbook/executor';
import { MAX_RESULT_CONTENT_CHARS, resolveBaseUrl, type SearchProviderAdapter, type SearchResult } from './index';

interface WikipediaSearchResponse {
  query?: {
    search?: Array<{
      title?: unknown;
      snippet?: unknown;
    }>;
  };
}

interface WikipediaExtractsResponse {
  query?: {
    pages?: Record<string, {
      title?: string;
      extract?: string;
    }>;
  };
}

export const wikipediaProvider: SearchProviderAdapter = {
  id: 'wikipedia',
  label: 'Wikipedia (keyless, encyclopedic only)',
  requiresKey: false,
  endpoint: 'https://en.wikipedia.org/w/api.php',
  signupUrl: 'https://www.mediawiki.org/wiki/API:Search',

  async search(query, opts): Promise<SearchResult[]> {
    const fetchImpl = opts.fetchImpl ?? fetch;
    const baseEndpoint = resolveBaseUrl(opts.baseUrl, this.endpoint);

    const searchUrl = new URL(baseEndpoint);
    searchUrl.searchParams.set('action', 'query');
    searchUrl.searchParams.set('list', 'search');
    searchUrl.searchParams.set('format', 'json');
    searchUrl.searchParams.set('origin', '*');
    searchUrl.searchParams.set('srlimit', String(Math.min(opts.maxResults, 10)));
    searchUrl.searchParams.set('srsearch', query);

    let response: Response;
    try {
      response = await fetchImpl(searchUrl.toString(), {
        method: 'GET',
        signal: opts.signal,
      });
    } catch (e) {
      const message = e instanceof Error ? e.message : String(e);
      throw new ToolNetworkError(`wikipedia request failed: ${message}`);
    }

    if (!response.ok) {
      throw new ToolNetworkError(`wikipedia request failed with HTTP ${response.status}`);
    }

    const contentType = response.headers.get('content-type') ?? '';
    if (contentType.includes('text/html')) {
      throw new ToolNetworkError(
        'Wikipedia: web-access base URL looks misconfigured — response was HTML instead of JSON. Check your search provider settings.'
      );
    }

    let json: WikipediaSearchResponse;
    try {
      json = await response.json() as WikipediaSearchResponse;
    } catch {
      throw new ToolNetworkError('wikipedia response was not valid JSON');
    }

    const articleBase = new URL(baseEndpoint).origin;
    const results = (json.query?.search ?? [])
      .map(item => normalizeResult(item, articleBase))
      .filter((result): result is SearchResult => !!result);

    if (!opts.includeContent || results.length === 0) return results;

    const extractMap = await fetchExtracts(
      results.map(r => r.title),
      baseEndpoint,
      fetchImpl,
      opts.signal
    );

    return results.map(r => ({
      ...r,
      ...(extractMap.has(r.title) ? { content: extractMap.get(r.title) } : {}),
    }));
  },
};

async function fetchExtracts(
  titles: string[],
  baseEndpoint: string,
  fetchImpl: typeof fetch,
  signal: AbortSignal
): Promise<Map<string, string>> {
  const url = new URL(baseEndpoint);
  url.searchParams.set('action', 'query');
  url.searchParams.set('prop', 'extracts');
  url.searchParams.set('format', 'json');
  url.searchParams.set('origin', '*');
  url.searchParams.set('explaintext', '1');
  url.searchParams.set('exintro', '1');
  url.searchParams.set('titles', titles.join('|'));

  let response: Response;
  try {
    response = await fetchImpl(url.toString(), { method: 'GET', signal });
  } catch {
    return new Map();
  }

  if (!response.ok) return new Map();

  let json: WikipediaExtractsResponse;
  try {
    json = await response.json() as WikipediaExtractsResponse;
  } catch {
    return new Map();
  }

  const map = new Map<string, string>();
  for (const page of Object.values(json.query?.pages ?? {})) {
    if (typeof page.title === 'string' && typeof page.extract === 'string' && page.extract) {
      map.set(page.title, capContent(page.extract));
    }
  }
  return map;
}

function capContent(text: string): string {
  if (text.length <= MAX_RESULT_CONTENT_CHARS) return text;
  return `${text.slice(0, MAX_RESULT_CONTENT_CHARS)}… [truncated: article continues beyond this point]`;
}

function normalizeResult(
  item: NonNullable<NonNullable<WikipediaSearchResponse['query']>['search']>[number],
  articleBase: string
): SearchResult | null {
  if (typeof item.title !== 'string' || !item.title) return null;
  return {
    title: item.title,
    url: `${articleBase}/wiki/${encodeURIComponent(item.title.replace(/ /g, '_'))}`,
    ...(typeof item.snippet === 'string' ? { snippet: stripHtml(item.snippet) } : {}),
  };
}

function stripHtml(text: string): string {
  return text.replace(/<[^>]+>/g, '').replace(/\s+/g, ' ').trim();
}
