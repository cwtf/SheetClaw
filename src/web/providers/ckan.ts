import { ToolNetworkError } from '../../workbook/executor';
import { MAX_RESULT_CONTENT_CHARS, resolveBaseUrl, type SearchProviderAdapter, type SearchResult } from './index';

interface CkanPackage {
  id?: unknown;
  name?: unknown;
  title?: unknown;
  notes?: unknown;
  organization?: { title?: unknown; name?: unknown };
  resources?: Array<{ name?: unknown; format?: unknown; url?: unknown }>;
  metadata_modified?: unknown;
}

interface CkanResponse {
  success?: unknown;
  result?: {
    results?: CkanPackage[];
  };
  error?: { message?: unknown };
}

export const ckanProvider: SearchProviderAdapter = {
  id: 'ckan',
  label: 'CKAN catalog (keyless, data.gov.au)',
  requiresKey: false,
  // catalog.data.gov retired its CKAN API (404, no CORS); data.gov.au serves
  // package_search with Access-Control-Allow-Origin: *, so it works from the task pane.
  endpoint: 'https://data.gov.au/data/api/3/action/package_search',
  signupUrl: 'https://docs.ckan.org/en/2.11/api/',

  async search(query, opts): Promise<SearchResult[]> {
    const fetchImpl = opts.fetchImpl ?? fetch;
    const endpoint = resolveBaseUrl(opts.baseUrl, this.endpoint);
    const url = new URL(endpoint);
    url.searchParams.set('q', query);
    url.searchParams.set('rows', String(Math.min(opts.maxResults, 10)));

    let response: Response;
    try {
      response = await fetchImpl(url.toString(), {
        method: 'GET',
        headers: { 'Accept': 'application/json' },
        signal: opts.signal,
      });
    } catch (e) {
      const message = e instanceof Error ? e.message : String(e);
      throw new ToolNetworkError(`ckan request failed: ${message}`);
    }

    if (!response.ok) throw new ToolNetworkError(`ckan request failed with HTTP ${response.status}`);
    const contentType = response.headers.get('content-type') ?? '';
    if (contentType.includes('text/html')) {
      throw new ToolNetworkError('CKAN: web-access base URL looks misconfigured - response was HTML instead of JSON. Check your search provider settings.');
    }

    let json: CkanResponse;
    try {
      json = await response.json() as CkanResponse;
    } catch {
      throw new ToolNetworkError('ckan response was not valid JSON');
    }
    if (json.success === false) {
      throw new ToolNetworkError(`ckan search failed: ${stringValue(json.error?.message) || 'unknown error'}`);
    }

    return (json.result?.results ?? [])
      .map(pkg => normalizeResult(pkg, endpoint, opts.includeContent))
      .filter((result): result is SearchResult => !!result);
  },
};

function normalizeResult(pkg: CkanPackage, endpoint: string, includeContent?: boolean): SearchResult | null {
  const id = stringValue(pkg.id) || stringValue(pkg.name);
  const title = stringValue(pkg.title) || stringValue(pkg.name) || id;
  if (!id || !title) return null;

  const resources = Array.isArray(pkg.resources) ? pkg.resources : [];
  const firstResourceUrl = resources.map(r => stringValue(r.url)).find(Boolean);
  const packageUrl = `${siteRootFromEndpoint(endpoint)}dataset/${encodeURIComponent(stringValue(pkg.name) || id)}`;
  const url = firstResourceUrl || packageUrl;
  const org = stringValue(pkg.organization?.title) || stringValue(pkg.organization?.name);
  const notes = stringValue(pkg.notes);
  const formats = [...new Set(resources.map(r => stringValue(r.format)).filter(Boolean))];

  return {
    title,
    url,
    ...(notes || org ? { snippet: capSnippet([org ? `Publisher: ${org}` : '', notes].filter(Boolean).join(' - ')) } : {}),
    ...(typeof pkg.metadata_modified === 'string' ? { publishedAt: pkg.metadata_modified } : {}),
    ...(includeContent ? { content: capContent([
      `Dataset: ${title}`,
      org ? `Publisher: ${org}` : '',
      notes ? `Notes: ${notes}` : '',
      formats.length ? `Formats: ${formats.join(', ')}` : '',
      `Package URL: ${packageUrl}`,
      firstResourceUrl ? `First resource URL: ${firstResourceUrl}` : '',
    ].filter(Boolean).join('\n')) } : {}),
  };
}

/**
 * Derives the CKAN site root from a package_search endpoint so dataset page
 * links stay correct when a custom base URL points at another CKAN instance,
 * e.g. "https://data.gov.au/data/api/3/action/package_search" → "https://data.gov.au/data/".
 */
function siteRootFromEndpoint(endpoint: string): string {
  const match = endpoint.match(/^(.*?)api\/(?:\d+\/)?action\/package_search/);
  if (match) return match[1];
  try {
    return `${new URL(endpoint).origin}/`;
  } catch {
    return endpoint.endsWith('/') ? endpoint : `${endpoint}/`;
  }
}

function stringValue(value: unknown): string {
  return typeof value === 'string' ? value.trim() : '';
}

function capSnippet(raw: string): string {
  return raw.length <= 500 ? raw : `${raw.slice(0, 500)}...`;
}

function capContent(raw: string): string {
  return raw.length <= MAX_RESULT_CONTENT_CHARS ? raw : `${raw.slice(0, MAX_RESULT_CONTENT_CHARS)}... [truncated: dataset metadata continues beyond this point]`;
}
