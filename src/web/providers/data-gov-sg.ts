import { ToolNetworkError } from '../../workbook/executor';
import { MAX_RESULT_CONTENT_CHARS, resolveBaseUrl, type SearchProviderAdapter, type SearchResult } from './index';

interface DataGovSgDataset {
  datasetId?: unknown;
  name?: unknown;
  status?: unknown;
  format?: unknown;
  managedByAgencyName?: unknown;
  managedBy?: unknown;
  lastUpdatedAt?: unknown;
  coverageStart?: unknown;
  coverageEnd?: unknown;
}

interface DataGovSgResponse {
  code?: unknown;
  data?: {
    datasets?: DataGovSgDataset[];
  };
  errorMsg?: unknown;
}

export const dataGovSgProvider: SearchProviderAdapter = {
  id: 'data-gov-sg',
  label: 'data.gov.sg datasets (keyless)',
  requiresKey: false,
  endpoint: 'https://api-production.data.gov.sg/v2/public/api/datasets',
  signupUrl: 'https://guide.data.gov.sg/developer-guide/api-overview',

  async search(query, opts): Promise<SearchResult[]> {
    const fetchImpl = opts.fetchImpl ?? fetch;
    const url = new URL(resolveBaseUrl(opts.baseUrl, this.endpoint));
    url.searchParams.set('page', '1');

    let response: Response;
    try {
      response = await fetchImpl(url.toString(), {
        method: 'GET',
        headers: { 'Accept': 'application/json' },
        signal: opts.signal,
      });
    } catch (e) {
      const message = e instanceof Error ? e.message : String(e);
      throw new ToolNetworkError(`data.gov.sg request failed: ${message}`);
    }

    if (!response.ok) throw new ToolNetworkError(`data.gov.sg request failed with HTTP ${response.status}`);
    const contentType = response.headers.get('content-type') ?? '';
    if (contentType.includes('text/html')) {
      throw new ToolNetworkError('data.gov.sg: web-access base URL looks misconfigured - response was HTML instead of JSON. Check your search provider settings.');
    }

    let json: DataGovSgResponse;
    try {
      json = await response.json() as DataGovSgResponse;
    } catch {
      throw new ToolNetworkError('data.gov.sg response was not valid JSON');
    }

    const error = stringValue(json.errorMsg);
    if (error) throw new ToolNetworkError(`data.gov.sg search failed: ${error}`);

    return rankDatasets(json.data?.datasets ?? [], query)
      .slice(0, opts.maxResults)
      .map(item => normalizeResult(item, opts.includeContent))
      .filter((result): result is SearchResult => !!result);
  },
};

function rankDatasets(items: DataGovSgDataset[], query: string): DataGovSgDataset[] {
  const terms = tokenize(query);
  return items
    .map(item => ({ item, score: scoreDataset(item, terms, query) }))
    .filter(hit => hit.score > 0)
    .sort((a, b) => b.score - a.score || stringValue(a.item.name).localeCompare(stringValue(b.item.name)))
    .map(hit => hit.item);
}

function scoreDataset(item: DataGovSgDataset, terms: string[], rawQuery: string): number {
  const id = stringValue(item.datasetId).toLowerCase();
  const name = stringValue(item.name).toLowerCase();
  const agency = (stringValue(item.managedByAgencyName) || stringValue(item.managedBy)).toLowerCase();
  const format = stringValue(item.format).toLowerCase();
  const query = rawQuery.trim().toLowerCase();
  let score = 0;
  if (id === query) score += 120;
  if (name === query) score += 100;
  if (name.includes(query)) score += 60;
  for (const term of terms) {
    if (id.includes(term)) score += 15;
    if (name.includes(term)) score += 12;
    if (agency.includes(term)) score += 6;
    if (format.includes(term)) score += 3;
  }
  return score;
}

function normalizeResult(item: DataGovSgDataset, includeContent?: boolean): SearchResult | null {
  const id = stringValue(item.datasetId);
  const name = stringValue(item.name);
  if (!id || !name) return null;
  const agency = stringValue(item.managedByAgencyName) || stringValue(item.managedBy);
  const format = stringValue(item.format);
  const metadataUrl = `https://api-production.data.gov.sg/v2/public/api/datasets/${encodeURIComponent(id)}/metadata`;
  const downloadStartUrl = `https://api-open.data.gov.sg/v1/public/api/datasets/${encodeURIComponent(id)}/initiate-download`;
  const downloadPollUrl = `https://api-open.data.gov.sg/v1/public/api/datasets/${encodeURIComponent(id)}/poll-download`;
  const coverage = [stringValue(item.coverageStart), stringValue(item.coverageEnd)].filter(Boolean).join(' to ');

  return {
    title: `${name} (${id})`,
    url: metadataUrl,
    snippet: [agency ? `Agency: ${agency}` : '', format ? `Format: ${format}` : '', coverage ? `Coverage: ${coverage}` : ''].filter(Boolean).join(' - '),
    ...(typeof item.lastUpdatedAt === 'string' ? { publishedAt: item.lastUpdatedAt } : {}),
    ...(includeContent ? { content: capContent([
      `Dataset: ${name}`,
      `ID: ${id}`,
      agency ? `Agency: ${agency}` : '',
      format ? `Format: ${format}` : '',
      coverage ? `Coverage: ${coverage}` : '',
      `Metadata API URL: ${metadataUrl}`,
      `Initiate download URL: ${downloadStartUrl}`,
      `Poll download URL: ${downloadPollUrl}`,
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
