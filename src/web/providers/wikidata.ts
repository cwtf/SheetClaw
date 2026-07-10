import { ToolNetworkError } from '../../workbook/executor';
import { MAX_RESULT_CONTENT_CHARS, resolveBaseUrl, type SearchProviderAdapter, type SearchResult } from './index';

interface WikidataSearchResponse {
  search?: Array<{
    id?: unknown;
    title?: unknown;
    label?: unknown;
    description?: unknown;
    concepturi?: unknown;
    url?: unknown;
    aliases?: unknown;
  }>;
}

export const wikidataProvider: SearchProviderAdapter = {
  id: 'wikidata',
  label: 'Wikidata (keyless entities)',
  requiresKey: false,
  endpoint: 'https://www.wikidata.org/w/api.php',
  signupUrl: 'https://www.wikidata.org/wiki/Special:ApiHelp/wbsearchentities',

  async search(query, opts): Promise<SearchResult[]> {
    const fetchImpl = opts.fetchImpl ?? fetch;
    const baseEndpoint = resolveBaseUrl(opts.baseUrl, this.endpoint);
    const url = new URL(baseEndpoint);
    url.searchParams.set('action', 'wbsearchentities');
    url.searchParams.set('format', 'json');
    url.searchParams.set('origin', '*');
    url.searchParams.set('language', 'en');
    url.searchParams.set('uselang', 'en');
    url.searchParams.set('type', 'item');
    url.searchParams.set('limit', String(Math.min(opts.maxResults, 10)));
    url.searchParams.set('search', query);

    let response: Response;
    try {
      response = await fetchImpl(url.toString(), {
        method: 'GET',
        headers: { 'Accept': 'application/json' },
        signal: opts.signal,
      });
    } catch (e) {
      const message = e instanceof Error ? e.message : String(e);
      throw new ToolNetworkError(`wikidata request failed: ${message}`);
    }

    if (!response.ok) {
      throw new ToolNetworkError(`wikidata request failed with HTTP ${response.status}`);
    }

    const contentType = response.headers.get('content-type') ?? '';
    if (contentType.includes('text/html')) {
      throw new ToolNetworkError(
        'Wikidata: web-access base URL looks misconfigured - response was HTML instead of JSON. Check your search provider settings.'
      );
    }

    let json: WikidataSearchResponse;
    try {
      json = await response.json() as WikidataSearchResponse;
    } catch {
      throw new ToolNetworkError('wikidata response was not valid JSON');
    }

    return (json.search ?? [])
      .map(item => normalizeResult(item, opts.includeContent))
      .filter((result): result is SearchResult => !!result);
  },
};

function normalizeResult(
  item: NonNullable<WikidataSearchResponse['search']>[number],
  includeContent?: boolean
): SearchResult | null {
  const id = typeof item.id === 'string' && item.id ? item.id : undefined;
  const title = typeof item.title === 'string' && item.title ? item.title : id;
  const label = typeof item.label === 'string' && item.label ? item.label : title;
  if (!label || !title) return null;

  const url = typeof item.concepturi === 'string' && item.concepturi
    ? item.concepturi
    : `https://www.wikidata.org/wiki/${encodeURIComponent(title)}`;
  const description = typeof item.description === 'string' ? item.description : undefined;
  const aliases = Array.isArray(item.aliases)
    ? item.aliases.filter((alias): alias is string => typeof alias === 'string' && !!alias)
    : [];
  const snippet = [id, description].filter(Boolean).join(' - ');

  return {
    title: id ? `${label} (${id})` : label,
    url,
    ...(snippet ? { snippet } : {}),
    ...(includeContent ? { content: capContent([
      `Entity: ${label}`,
      id ? `ID: ${id}` : '',
      description ? `Description: ${description}` : '',
      aliases.length ? `Aliases: ${aliases.join(', ')}` : '',
      `Wikidata URL: ${url}`,
    ].filter(Boolean).join('\n')) } : {}),
  };
}

function capContent(raw: string): string {
  if (raw.length <= MAX_RESULT_CONTENT_CHARS) return raw;
  return `${raw.slice(0, MAX_RESULT_CONTENT_CHARS)}... [truncated: entity summary continues beyond this point]`;
}
