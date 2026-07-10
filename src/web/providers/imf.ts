import { ToolNetworkError } from '../../workbook/executor';
import { MAX_RESULT_CONTENT_CHARS, READER_PROVIDER_ENDPOINT, resolveBaseUrl, type SearchProviderAdapter, type SearchResult } from './index';

type ImfIndicatorEntry = {
  label?: unknown;
  description?: unknown;
  unit?: unknown;
};

interface ImfIndicatorsResponse {
  indicators?: Record<string, ImfIndicatorEntry> | Array<ImfIndicatorEntry & { id?: unknown; code?: unknown }>;
}

export const imfProvider: SearchProviderAdapter = {
  id: 'imf',
  label: 'IMF DataMapper (keyless)',
  requiresKey: false,
  endpoint: 'https://www.imf.org/external/datamapper/api/v1/indicators',
  signupUrl: 'https://www.imf.org/external/datamapper/api/help',

  async search(query, opts): Promise<SearchResult[]> {
    const fetchImpl = opts.fetchImpl ?? fetch;
    const url = resolveBaseUrl(opts.baseUrl, this.endpoint);
    const json = await fetchIndicators(url, fetchImpl, opts.signal);

    return rankIndicators(normalizeIndicators(json.indicators), query)
      .slice(0, opts.maxResults)
      .map(item => toResult(item, opts.includeContent));
  },
};

async function fetchIndicators(url: string, fetchImpl: typeof fetch, signal: AbortSignal): Promise<ImfIndicatorsResponse> {
  let response: Response;
  try {
    response = await fetchImpl(url, {
      method: 'GET',
      headers: { 'Accept': 'application/json' },
      signal,
    });
  } catch (e) {
    if (signal.aborted) throw e;
    // The DataMapper API sends no Access-Control-Allow-Origin header, so a direct fetch
    // always fails inside the task pane; retry through the CORS-enabled reader proxy.
    return fetchIndicatorsViaReader(url, fetchImpl, signal, e);
  }

  if (!response.ok) throw new ToolNetworkError(`imf request failed with HTTP ${response.status}`);
  const contentType = response.headers.get('content-type') ?? '';
  if (contentType.includes('text/html')) {
    throw new ToolNetworkError('IMF: web-access base URL looks misconfigured - response was HTML instead of JSON. Check your search provider settings.');
  }

  try {
    return await response.json() as ImfIndicatorsResponse;
  } catch {
    throw new ToolNetworkError('imf response was not valid JSON');
  }
}

/** The reader proxy wraps the fetched body in a JSON envelope: { data: { content: "<raw body>" } }. */
async function fetchIndicatorsViaReader(url: string, fetchImpl: typeof fetch, signal: AbortSignal, directError: unknown): Promise<ImfIndicatorsResponse> {
  const directMessage = directError instanceof Error ? directError.message : String(directError);

  let response: Response;
  try {
    response = await fetchImpl(`${READER_PROVIDER_ENDPOINT}${url}`, {
      method: 'GET',
      headers: { 'Accept': 'application/json' },
      signal,
    });
  } catch (e) {
    const message = e instanceof Error ? e.message : String(e);
    throw new ToolNetworkError(`imf request failed directly (${directMessage}) and via the reader proxy (${message})`);
  }

  if (!response.ok) {
    throw new ToolNetworkError(`imf request failed directly (${directMessage}) and the reader proxy returned HTTP ${response.status}`);
  }

  let envelope: { data?: { content?: unknown } };
  try {
    envelope = await response.json() as { data?: { content?: unknown } };
  } catch {
    throw new ToolNetworkError('imf reader-proxy response was not valid JSON');
  }

  const content = envelope?.data?.content;
  if (typeof content !== 'string') {
    throw new ToolNetworkError('imf reader-proxy response was missing extracted content');
  }
  try {
    return JSON.parse(content) as ImfIndicatorsResponse;
  } catch {
    throw new ToolNetworkError('imf reader-proxy content was not valid JSON');
  }
}

interface Indicator {
  id: string;
  label: string;
  description: string;
  unit: string;
}

function normalizeIndicators(value: ImfIndicatorsResponse['indicators']): Indicator[] {
  if (Array.isArray(value)) {
    return value.map(item => ({
      id: stringValue(item.id) || stringValue(item.code),
      label: stringValue(item.label),
      description: stringValue(item.description),
      unit: stringValue(item.unit),
    })).filter(item => item.id && item.label);
  }
  return Object.entries(value ?? {}).map(([id, item]) => ({
    id,
    label: stringValue(item.label),
    description: stringValue(item.description),
    unit: stringValue(item.unit),
  })).filter(item => item.id && item.label);
}

function rankIndicators(items: Indicator[], query: string): Indicator[] {
  const terms = tokenize(query);
  return items
    .map(item => ({ item, score: scoreIndicator(item, terms, query) }))
    .filter(hit => hit.score > 0)
    .sort((a, b) => b.score - a.score || a.item.label.localeCompare(b.item.label))
    .map(hit => hit.item);
}

function scoreIndicator(item: Indicator, terms: string[], rawQuery: string): number {
  const id = item.id.toLowerCase();
  const label = item.label.toLowerCase();
  const description = item.description.toLowerCase();
  const query = rawQuery.trim().toLowerCase();
  let score = 0;
  if (id === query) score += 200;
  if (label === query) score += 120;
  if (label.includes(query)) score += 70;
  for (const term of terms) {
    if (id.includes(term)) score += 20;
    if (label.includes(term)) score += 12;
    if (description.includes(term)) score += 4;
  }
  return score;
}

function toResult(item: Indicator, includeContent?: boolean): SearchResult {
  const url = `https://www.imf.org/external/datamapper/api/v1/${encodeURIComponent(item.id)}`;
  return {
    title: `${item.label} (${item.id})`,
    url,
    snippet: [item.unit ? `Unit: ${item.unit}` : '', item.description].filter(Boolean).join(' - '),
    ...(includeContent ? { content: capContent([
      `Indicator: ${item.label}`,
      `ID: ${item.id}`,
      item.unit ? `Unit: ${item.unit}` : '',
      item.description ? `Description: ${item.description}` : '',
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
  return raw.length <= MAX_RESULT_CONTENT_CHARS ? raw : `${raw.slice(0, MAX_RESULT_CONTENT_CHARS)}... [truncated: indicator metadata continues beyond this point]`;
}
