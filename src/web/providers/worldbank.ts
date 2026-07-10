import { ToolNetworkError } from '../../workbook/executor';
import { MAX_RESULT_CONTENT_CHARS, resolveBaseUrl, type SearchProviderAdapter, type SearchResult } from './index';

interface WorldBankIndicator {
  id?: unknown;
  name?: unknown;
  sourceNote?: unknown;
  sourceOrganization?: unknown;
  topics?: Array<{ value?: unknown }>;
}

type WorldBankIndicatorResponse = [
  {
    page?: number;
    pages?: number;
    per_page?: unknown;
    total?: number;
  },
  WorldBankIndicator[]
];

const WORLD_BANK_DATA_BASE = 'https://api.worldbank.org/v2/country/all/indicator/';
const DEFAULT_PER_PAGE = 20000;

export const worldBankProvider: SearchProviderAdapter = {
  id: 'worldbank',
  label: 'World Bank Indicators (keyless)',
  requiresKey: false,
  endpoint: 'https://api.worldbank.org/v2/indicator',
  signupUrl: 'https://datahelpdesk.worldbank.org/knowledgebase/articles/889392-api-documentation',

  async search(query, opts): Promise<SearchResult[]> {
    const fetchImpl = opts.fetchImpl ?? fetch;
    const url = new URL(resolveBaseUrl(opts.baseUrl, this.endpoint));
    url.searchParams.set('format', 'json');
    url.searchParams.set('per_page', String(DEFAULT_PER_PAGE));

    let response: Response;
    try {
      response = await fetchImpl(url.toString(), {
        method: 'GET',
        headers: { 'Accept': 'application/json' },
        signal: opts.signal,
      });
    } catch (e) {
      const message = e instanceof Error ? e.message : String(e);
      throw new ToolNetworkError(`worldbank request failed: ${message}`);
    }

    if (!response.ok) {
      throw new ToolNetworkError(`worldbank request failed with HTTP ${response.status}`);
    }

    const contentType = response.headers.get('content-type') ?? '';
    if (contentType.includes('text/html')) {
      throw new ToolNetworkError(
        'World Bank: web-access base URL looks misconfigured - response was HTML instead of JSON. Check your search provider settings.'
      );
    }

    let json: WorldBankIndicatorResponse;
    try {
      json = await response.json() as WorldBankIndicatorResponse;
    } catch {
      throw new ToolNetworkError('worldbank response was not valid JSON');
    }

    const indicators = Array.isArray(json?.[1]) ? json[1] : [];
    return rankIndicators(indicators, query)
      .slice(0, opts.maxResults)
      .map(item => normalizeResult(item, opts.includeContent))
      .filter((result): result is SearchResult => !!result);
  },
};

function rankIndicators(indicators: WorldBankIndicator[], query: string): WorldBankIndicator[] {
  const terms = tokenize(query);
  if (terms.length === 0) return [];

  return indicators
    .map(item => ({ item, score: scoreIndicator(item, terms, query) }))
    .filter(({ score }) => score > 0)
    .sort((a, b) => b.score - a.score || compareIndicatorNames(a.item, b.item))
    .map(({ item }) => item);
}

function scoreIndicator(item: WorldBankIndicator, terms: string[], rawQuery: string): number {
  const id = valueToString(item.id).toLowerCase();
  const name = valueToString(item.name).toLowerCase();
  const note = valueToString(item.sourceNote).toLowerCase();
  const topics = topicValues(item).join(' ').toLowerCase();
  const query = rawQuery.trim().toLowerCase();
  let score = 0;

  if (id === query) score += 200;
  if (id.includes(query)) score += 80;
  if (name === query) score += 120;
  if (name.includes(query)) score += 60;

  for (const term of terms) {
    if (id.includes(term)) score += 18;
    if (name.includes(term)) score += 12;
    if (topics.includes(term)) score += 5;
    if (note.includes(term)) score += 2;
  }
  return score;
}

function normalizeResult(item: WorldBankIndicator, includeContent?: boolean): SearchResult | null {
  const id = valueToString(item.id);
  const name = valueToString(item.name);
  if (!id || !name) return null;

  const topics = topicValues(item);
  const sourceNote = valueToString(item.sourceNote);
  const sourceOrganization = valueToString(item.sourceOrganization);
  const dataUrl = `${WORLD_BANK_DATA_BASE}${encodeURIComponent(id)}?format=json&per_page=20000`;
  const snippetParts = [
    topics.length ? `Topics: ${topics.join(', ')}` : '',
    sourceNote,
  ].filter(Boolean);

  return {
    title: `${name} (${id})`,
    url: dataUrl,
    ...(snippetParts.length ? { snippet: capSnippet(snippetParts.join(' - ')) } : {}),
    ...(includeContent ? { content: capContent([
      `Indicator: ${name}`,
      `ID: ${id}`,
      topics.length ? `Topics: ${topics.join(', ')}` : '',
      sourceNote ? `Source note: ${sourceNote}` : '',
      sourceOrganization ? `Source organization: ${sourceOrganization}` : '',
      `Data API URL: ${dataUrl}`,
    ].filter(Boolean).join('\n')) } : {}),
  };
}

function tokenize(query: string): string[] {
  return query
    .toLowerCase()
    .split(/[^a-z0-9.]+/)
    .map(term => term.trim())
    .filter(term => term.length >= 2);
}

function topicValues(item: WorldBankIndicator): string[] {
  return Array.isArray(item.topics)
    ? item.topics.map(topic => valueToString(topic.value)).filter(Boolean)
    : [];
}

function valueToString(value: unknown): string {
  return typeof value === 'string' ? value.trim() : '';
}

function compareIndicatorNames(a: WorldBankIndicator, b: WorldBankIndicator): number {
  return valueToString(a.name).localeCompare(valueToString(b.name));
}

function capSnippet(raw: string): string {
  if (raw.length <= 500) return raw;
  return `${raw.slice(0, 500)}...`;
}

function capContent(raw: string): string {
  if (raw.length <= MAX_RESULT_CONTENT_CHARS) return raw;
  return `${raw.slice(0, MAX_RESULT_CONTENT_CHARS)}... [truncated: indicator metadata continues beyond this point]`;
}
