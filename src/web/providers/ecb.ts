import { ToolNetworkError } from '../../workbook/executor';
import { MAX_RESULT_CONTENT_CHARS, resolveBaseUrl, type SearchProviderAdapter, type SearchResult } from './index';

interface EcbDataflow {
  id?: unknown;
  name?: unknown;
  description?: unknown;
}

interface EcbResponse {
  dataflows?: EcbDataflow[];
  data?: {
    dataflows?: EcbDataflow[];
  };
}

export const ecbProvider: SearchProviderAdapter = {
  id: 'ecb',
  label: 'ECB Data Portal (keyless)',
  requiresKey: false,
  endpoint: 'https://data-api.ecb.europa.eu/service/dataflow/ECB/ALL/latest',
  signupUrl: 'https://data.ecb.europa.eu/help/api/overview',

  async search(query, opts): Promise<SearchResult[]> {
    const fetchImpl = opts.fetchImpl ?? fetch;
    const url = new URL(resolveBaseUrl(opts.baseUrl, this.endpoint));
    url.searchParams.set('format', 'jsondata');

    let response: Response;
    try {
      response = await fetchImpl(url.toString(), {
        method: 'GET',
        headers: { 'Accept': 'application/json' },
        signal: opts.signal,
      });
    } catch (e) {
      const message = e instanceof Error ? e.message : String(e);
      throw new ToolNetworkError(`ecb request failed: ${message}`);
    }

    if (!response.ok) throw new ToolNetworkError(`ecb request failed with HTTP ${response.status}`);
    const contentType = response.headers.get('content-type') ?? '';
    if (contentType.includes('text/html')) {
      throw new ToolNetworkError('ECB: web-access base URL looks misconfigured - response was HTML instead of JSON. Check your search provider settings.');
    }

    let json: EcbResponse;
    try {
      json = await response.json() as EcbResponse;
    } catch {
      throw new ToolNetworkError('ecb response was not valid JSON');
    }

    return rankDataflows(json.dataflows ?? json.data?.dataflows ?? [], query)
      .slice(0, opts.maxResults)
      .map(item => toResult(item, opts.includeContent))
      .filter((result): result is SearchResult => !!result);
  },
};

function rankDataflows(items: EcbDataflow[], query: string): EcbDataflow[] {
  const terms = tokenize(query);
  return items
    .map(item => ({ item, score: scoreDataflow(item, terms, query) }))
    .filter(hit => hit.score > 0)
    .sort((a, b) => b.score - a.score || stringValue(a.item.name).localeCompare(stringValue(b.item.name)))
    .map(hit => hit.item);
}

function scoreDataflow(item: EcbDataflow, terms: string[], rawQuery: string): number {
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

function toResult(item: EcbDataflow, includeContent?: boolean): SearchResult | null {
  const id = stringValue(item.id);
  const name = stringValue(item.name) || id;
  if (!id || !name) return null;
  const description = stringValue(item.description);
  const url = `https://data-api.ecb.europa.eu/service/data/${encodeURIComponent(id)}?format=jsondata`;
  return {
    title: `${name} (${id})`,
    url,
    ...(description ? { snippet: description } : {}),
    ...(includeContent ? { content: capContent([
      `Dataflow: ${name}`,
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
  return raw.length <= MAX_RESULT_CONTENT_CHARS ? raw : `${raw.slice(0, MAX_RESULT_CONTENT_CHARS)}... [truncated: dataflow metadata continues beyond this point]`;
}
