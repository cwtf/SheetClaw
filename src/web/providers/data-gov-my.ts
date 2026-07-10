import { MAX_RESULT_CONTENT_CHARS, type SearchProviderAdapter, type SearchResult } from './index';

interface Candidate {
  title: string;
  url: string;
  snippet: string;
  content: string;
}

export const dataGovMyProvider: SearchProviderAdapter = {
  id: 'data-gov-my',
  label: 'data.gov.my APIs (keyless)',
  requiresKey: false,
  endpoint: 'https://api.data.gov.my/data-catalogue',
  signupUrl: 'https://developer.data.gov.my/quickstart',

  async search(query, opts): Promise<SearchResult[]> {
    const id = toDatasetId(query);
    const candidates = buildCandidates(query, id);
    return candidates.slice(0, opts.maxResults).map(item => ({
      title: item.title,
      url: item.url,
      snippet: item.snippet,
      ...(opts.includeContent ? { content: capContent(item.content) } : {}),
    }));
  },
};

function buildCandidates(query: string, id: string): Candidate[] {
  const lowered = query.toLowerCase();
  const results: Candidate[] = [];
  if (lowered.includes('weather') || lowered.includes('forecast') || lowered.includes('rain')) {
    results.push(candidate(
      'Malaysia weather forecast',
      'https://api.data.gov.my/weather/forecast?limit=100',
      '7-day general forecast from MET Malaysia.',
      'Use contains=<place>@location__location_name to filter by place, or contains=St@location__location_id for state-level forecasts.'
    ));
  }
  if (lowered.includes('warning') || lowered.includes('storm') || lowered.includes('earthquake')) {
    results.push(candidate(
      'Malaysia weather warnings',
      lowered.includes('earthquake')
        ? 'https://api.data.gov.my/weather/warning/earthquake?limit=100'
        : 'https://api.data.gov.my/weather/warning?limit=100',
      'Live weather warning endpoint from MET Malaysia.',
      'Use timestamp filters documented by data.gov.my for warning_issue fields when narrowing a period.'
    ));
  }

  results.push(candidate(
    `data.gov.my data catalogue candidate: ${id}`,
    `https://api.data.gov.my/data-catalogue?id=${encodeURIComponent(id)}&limit=100`,
    'General public-sector data catalogue endpoint. The id is inferred from the query; refine it if the preview returns no rows.',
    'The Data Catalogue API requires an id parameter. Search result uses a normalized query-derived id as a starting point.'
  ));
  results.push(candidate(
    `OpenDOSM candidate: ${id}`,
    `https://api.data.gov.my/opendosm?id=${encodeURIComponent(id)}&limit=100`,
    'Department of Statistics Malaysia OpenDOSM endpoint. The id is inferred from the query; refine it if the preview returns no rows.',
    'The OpenDOSM API requires an id parameter shown on OpenDOSM dataset pages. Search result uses a normalized query-derived id as a starting point.'
  ));
  return results;
}

function candidate(title: string, url: string, snippet: string, note: string): Candidate {
  return {
    title,
    url,
    snippet,
    content: [`Title: ${title}`, `API URL: ${url}`, `Note: ${note}`].join('\n'),
  };
}

function toDatasetId(query: string): string {
  const id = query.toLowerCase()
    .replace(/[^a-z0-9]+/g, '_')
    .replace(/^_+|_+$/g, '')
    .replace(/_+/g, '_');
  return id || 'dataset';
}

function capContent(raw: string): string {
  return raw.length <= MAX_RESULT_CONTENT_CHARS ? raw : `${raw.slice(0, MAX_RESULT_CONTENT_CHARS)}... [truncated: API candidate metadata continues beyond this point]`;
}
