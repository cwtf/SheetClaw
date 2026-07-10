import { ToolNetworkError } from '../../workbook/executor';
import { MAX_RESULT_CONTENT_CHARS, resolveBaseUrl, type SearchProviderAdapter, type SearchResult } from './index';

interface UnSdgIndicator {
  code?: unknown;
  indicator?: unknown;
  description?: unknown;
  goal?: unknown;
  target?: unknown;
}

export const unSdgProvider: SearchProviderAdapter = {
  id: 'un-sdg',
  label: 'UN SDG API (keyless)',
  requiresKey: false,
  endpoint: 'https://unstats.un.org/sdgapi/v1/sdg/Indicator/List',
  signupUrl: 'https://unstats.un.org/sdgapi/swagger/',

  async search(query, opts): Promise<SearchResult[]> {
    const fetchImpl = opts.fetchImpl ?? fetch;
    const url = resolveBaseUrl(opts.baseUrl, this.endpoint);

    let response: Response;
    try {
      response = await fetchImpl(url, {
        method: 'GET',
        headers: { 'Accept': 'application/json' },
        signal: opts.signal,
      });
    } catch (e) {
      const message = e instanceof Error ? e.message : String(e);
      throw new ToolNetworkError(`un-sdg request failed: ${message}`);
    }

    if (!response.ok) throw new ToolNetworkError(`un-sdg request failed with HTTP ${response.status}`);
    const contentType = response.headers.get('content-type') ?? '';
    if (contentType.includes('text/html')) {
      throw new ToolNetworkError('UN SDG: web-access base URL looks misconfigured - response was HTML instead of JSON. Check your search provider settings.');
    }

    let json: UnSdgIndicator[];
    try {
      json = await response.json() as UnSdgIndicator[];
    } catch {
      throw new ToolNetworkError('un-sdg response was not valid JSON');
    }

    return rankIndicators(Array.isArray(json) ? json : [], query)
      .slice(0, opts.maxResults)
      .map(item => toResult(item, opts.includeContent))
      .filter((result): result is SearchResult => !!result);
  },
};

function rankIndicators(items: UnSdgIndicator[], query: string): UnSdgIndicator[] {
  const terms = tokenize(query);
  return items
    .map(item => ({ item, score: scoreIndicator(item, terms, query) }))
    .filter(hit => hit.score > 0)
    .sort((a, b) => b.score - a.score || stringValue(a.item.indicator).localeCompare(stringValue(b.item.indicator)))
    .map(hit => hit.item);
}

function scoreIndicator(item: UnSdgIndicator, terms: string[], rawQuery: string): number {
  const code = stringValue(item.code).toLowerCase();
  const indicator = stringValue(item.indicator).toLowerCase();
  const description = stringValue(item.description).toLowerCase();
  const goal = stringValue(item.goal).toLowerCase();
  const target = stringValue(item.target).toLowerCase();
  const query = rawQuery.trim().toLowerCase();
  let score = 0;
  if (code === query) score += 200;
  if (indicator.includes(query)) score += 70;
  for (const term of terms) {
    if (code.includes(term)) score += 18;
    if (indicator.includes(term)) score += 12;
    if (description.includes(term)) score += 6;
    if (goal.includes(term) || target.includes(term)) score += 4;
  }
  return score;
}

function toResult(item: UnSdgIndicator, includeContent?: boolean): SearchResult | null {
  const code = stringValue(item.code);
  const indicator = stringValue(item.indicator) || stringValue(item.description);
  if (!code || !indicator) return null;
  const url = `https://unstats.un.org/sdgapi/v1/sdg/Series/List?indicator=${encodeURIComponent(code)}`;
  const goal = stringValue(item.goal);
  const target = stringValue(item.target);
  return {
    title: `${indicator} (${code})`,
    url,
    snippet: [goal ? `Goal: ${goal}` : '', target ? `Target: ${target}` : '', stringValue(item.description)].filter(Boolean).join(' - '),
    ...(includeContent ? { content: capContent([
      `Indicator: ${indicator}`,
      `Code: ${code}`,
      goal ? `Goal: ${goal}` : '',
      target ? `Target: ${target}` : '',
      stringValue(item.description) ? `Description: ${stringValue(item.description)}` : '',
      `Series API URL: ${url}`,
    ].filter(Boolean).join('\n')) } : {}),
  };
}

function tokenize(query: string): string[] {
  return query.toLowerCase().split(/[^a-z0-9.]+/).map(t => t.trim()).filter(t => t.length >= 2);
}

function stringValue(value: unknown): string {
  if (typeof value === 'number') return String(value);
  return typeof value === 'string' ? value.trim() : '';
}

function capContent(raw: string): string {
  return raw.length <= MAX_RESULT_CONTENT_CHARS ? raw : `${raw.slice(0, MAX_RESULT_CONTENT_CHARS)}... [truncated: indicator metadata continues beyond this point]`;
}
