import { ToolNetworkError } from '../../workbook/executor';
import { MAX_RESULT_CONTENT_CHARS, resolveBaseUrl, type SearchProviderAdapter, type SearchResult } from './index';

interface EurostatDataflow {
  id?: unknown;
  name?: unknown;
  description?: unknown;
  agencyID?: unknown;
}

interface EurostatResponse {
  dataflows?: EurostatDataflow[];
  data?: {
    dataflows?: EurostatDataflow[];
  };
}

export const eurostatProvider: SearchProviderAdapter = {
  id: 'eurostat',
  label: 'Eurostat datasets (keyless)',
  requiresKey: false,
  endpoint: 'https://ec.europa.eu/eurostat/api/dissemination/sdmx/2.1/dataflow/ESTAT/all/latest',
  signupUrl: 'https://ec.europa.eu/eurostat/web/user-guides/data-browser/api-data-access/api-detailed-guidelines/api-statistics',

  async search(query, opts): Promise<SearchResult[]> {
    const fetchImpl = opts.fetchImpl ?? fetch;
    const url = new URL(resolveBaseUrl(opts.baseUrl, this.endpoint));
    url.searchParams.set('format', 'JSON');

    let response: Response;
    try {
      response = await fetchImpl(url.toString(), {
        method: 'GET',
        headers: { 'Accept': 'application/json' },
        signal: opts.signal,
      });
    } catch (e) {
      const message = e instanceof Error ? e.message : String(e);
      throw new ToolNetworkError(`eurostat request failed: ${message}`);
    }

    if (!response.ok) throw new ToolNetworkError(`eurostat request failed with HTTP ${response.status}`);
    const contentType = response.headers.get('content-type') ?? '';
    if (contentType.includes('text/html')) {
      throw new ToolNetworkError('Eurostat: web-access base URL looks misconfigured - response was HTML instead of JSON. Check your search provider settings.');
    }

    let json: EurostatResponse;
    try {
      json = await response.json() as EurostatResponse;
    } catch {
      throw new ToolNetworkError('eurostat response was not valid JSON');
    }

    return rankDataflows(json.dataflows ?? json.data?.dataflows ?? [], query)
      .slice(0, opts.maxResults)
      .map(item => toResult(item, opts.includeContent))
      .filter((result): result is SearchResult => !!result);
  },
};

function rankDataflows(items: EurostatDataflow[], query: string): EurostatDataflow[] {
  const terms = tokenize(query);
  return items
    .map(item => ({ item, score: scoreDataflow(item, terms, query) }))
    .filter(hit => hit.score > 0)
    .sort((a, b) => b.score - a.score || stringValue(a.item.name).localeCompare(stringValue(b.item.name)))
    .map(hit => hit.item);
}

function scoreDataflow(item: EurostatDataflow, terms: string[], rawQuery: string): number {
  const id = stringValue(item.id).toLowerCase();
  const name = stringValue(item.name).toLowerCase();
  const description = stringValue(item.description).toLowerCase();
  const query = rawQuery.trim().toLowerCase();
  let score = 0;
  if (id === query) score += 200;
  if (id.includes(query)) score += 80;
  if (name.includes(query)) score += 60;
  for (const term of terms) {
    if (id.includes(term)) score += 18;
    if (name.includes(term)) score += 12;
    if (description.includes(term)) score += 4;
  }
  return score;
}

function toResult(item: EurostatDataflow, includeContent?: boolean): SearchResult | null {
  const id = stringValue(item.id);
  const name = stringValue(item.name) || id;
  if (!id || !name) return null;
  const description = stringValue(item.description);
  const url = `https://ec.europa.eu/eurostat/api/dissemination/statistics/1.0/data/${encodeURIComponent(id)}?lang=en`;
  return {
    title: `${name} (${id})`,
    url,
    ...(description ? { snippet: description } : {}),
    ...(includeContent ? { content: capContent([
      `Dataset: ${name}`,
      `ID: ${id}`,
      description ? `Description: ${description}` : '',
      `Data API URL: ${url}`,
    ].filter(Boolean).join('\n')) } : {}),
  };
}

function tokenize(query: string): string[] {
  return query.toLowerCase().split(/[^a-z0-9_]+/).map(t => t.trim()).filter(t => t.length >= 2);
}

function stringValue(value: unknown): string {
  return typeof value === 'string' ? value.trim() : '';
}

function capContent(raw: string): string {
  return raw.length <= MAX_RESULT_CONTENT_CHARS ? raw : `${raw.slice(0, MAX_RESULT_CONTENT_CHARS)}... [truncated: dataset metadata continues beyond this point]`;
}
