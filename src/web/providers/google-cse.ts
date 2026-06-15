import { ToolNetworkError, ToolValidationError } from '../../workbook/executor';
import { resolveBaseUrl, type SearchProviderAdapter, type SearchResult } from './index';

interface GoogleCseResponse {
  items?: Array<{
    title?: unknown;
    link?: unknown;
    snippet?: unknown;
  }>;
}

export const googleCseProvider: SearchProviderAdapter = {
  id: 'google-cse',
  label: 'Google Programmable Search',
  requiresKey: true,
  requiresEngineId: true,
  endpoint: 'https://www.googleapis.com/customsearch/v1',
  signupUrl: 'https://programmablesearchengine.google.com/',

  async search(query, opts): Promise<SearchResult[]> {
    const engineId = opts.engineId?.trim();
    if (!engineId) {
      throw new ToolValidationError('Google Programmable Search requires an engine ID (cx). Add it in Settings > Search.');
    }

    const fetchImpl = opts.fetchImpl ?? fetch;
    const url = new URL(resolveBaseUrl(opts.baseUrl, this.endpoint));
    url.searchParams.set('cx', engineId);
    url.searchParams.set('q', query);
    url.searchParams.set('num', String(Math.min(opts.maxResults, 10)));

    let response: Response;
    try {
      response = await fetchImpl(url.toString(), {
        method: 'GET',
        // API key goes in a header, not the query string, so it stays out of URL logs.
        headers: { 'X-Goog-Api-Key': opts.apiKey },
        signal: opts.signal,
      });
    } catch (e) {
      const message = e instanceof Error ? e.message : String(e);
      throw new ToolNetworkError(`google-cse request failed: ${message}`);
    }

    if (!response.ok) {
      throw new ToolNetworkError(`google-cse request failed with HTTP ${response.status}`);
    }

    const contentType = response.headers.get('content-type') ?? '';
    if (contentType.includes('text/html')) {
      throw new ToolNetworkError(
        'Google CSE: web-access base URL looks misconfigured — response was HTML instead of JSON. Check your search provider settings.'
      );
    }

    let json: GoogleCseResponse;
    try {
      json = await response.json() as GoogleCseResponse;
    } catch {
      throw new ToolNetworkError('google-cse response was not valid JSON');
    }

    return (json.items ?? [])
      .map(item => normalizeResult(item))
      .filter((result): result is SearchResult => !!result);
  },
};

function normalizeResult(item: NonNullable<GoogleCseResponse['items']>[number]): SearchResult | null {
  if (typeof item.title !== 'string' || typeof item.link !== 'string') return null;
  return {
    title: item.title,
    url: item.link,
    ...(typeof item.snippet === 'string' ? { snippet: item.snippet } : {}),
  };
}
