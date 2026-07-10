import { describe, expect, it, vi } from 'vitest';
import { createWebSearchHandler, WEB_SEARCH } from '../search';
import { MAX_RESULT_CONTENT_CHARS, resolveBaseUrl } from '../providers';
import { tavilyProvider } from '../providers/tavily';
import { googleCseProvider } from '../providers/google-cse';
import { jinaProvider } from '../providers/jina';
import { searxngProvider } from '../providers/searxng';
import { wikipediaProvider } from '../providers/wikipedia';
import { wikidataProvider } from '../providers/wikidata';
import { worldBankProvider } from '../providers/worldbank';
import { ckanProvider } from '../providers/ckan';
import { dataGovMyProvider } from '../providers/data-gov-my';
import { dataGovSgProvider } from '../providers/data-gov-sg';
import { ecbProvider } from '../providers/ecb';
import { eurostatProvider } from '../providers/eurostat';
import { imfProvider } from '../providers/imf';
import { openMeteoProvider } from '../providers/open-meteo';
import { unSdgProvider } from '../providers/un-sdg';
import { ToolExecutor, ToolValidationError } from '../../workbook/executor';
import { WorkbookRegistry } from '../../workbook/registry';
import type { ToolCall } from '../../types';

function jsonResponse(data: unknown, status = 200): Response {
  return new Response(JSON.stringify(data), {
    status,
    headers: { 'content-type': 'application/json' },
  });
}

function textResponse(body: string, contentType = 'text/html'): Response {
  return new Response(body, {
    status: 200,
    headers: { 'content-type': contentType },
  });
}

const DOMAIN_FIXTURES = [
  {
    title: 'Quarterly filing',
    url: 'https://finance.example/reports/q1',
    content: 'Revenue table and notes.',
  },
  {
    title: 'Hourly forecast',
    url: 'https://weather.example/api/hourly',
    content: 'Temperature and precipitation fields.',
  },
  {
    title: 'Match schedule',
    url: 'https://sports.example/schedule',
    content: 'Fixtures and scores.',
  },
];

describe('Tavily adapter', () => {
  it('parses documented JSON results from unrelated domains', async () => {
    const fetchImpl = vi.fn(async (_url: RequestInfo | URL, _init?: RequestInit) => jsonResponse({ results: DOMAIN_FIXTURES }));

    const results = await tavilyProvider.search('public data', {
      maxResults: 3,
      apiKey: 'key',
      signal: new AbortController().signal,
      fetchImpl,
    });

    expect(results).toEqual([
      { title: 'Quarterly filing', url: 'https://finance.example/reports/q1', snippet: 'Revenue table and notes.' },
      { title: 'Hourly forecast', url: 'https://weather.example/api/hourly', snippet: 'Temperature and precipitation fields.' },
      { title: 'Match schedule', url: 'https://sports.example/schedule', snippet: 'Fixtures and scores.' },
    ]);
  });

  it('preserves zero results as an empty array', async () => {
    const fetchImpl = vi.fn(async (_url: RequestInfo | URL, _init?: RequestInit) => jsonResponse({ results: [] }));

    const results = await tavilyProvider.search('no matches', {
      maxResults: 3,
      apiKey: 'key',
      signal: new AbortController().signal,
      fetchImpl,
    });

    expect(results).toEqual([]);
  });

  it('throws a clear config error when Tavily returns HTML instead of JSON', async () => {
    const fetchImpl = vi.fn(async (_url: RequestInfo | URL, _init?: RequestInit) =>
      textResponse('<html><title>Service unavailable</title></html>')
    );

    await expect(tavilyProvider.search('public data', {
      maxResults: 3,
      apiKey: 'key',
      signal: new AbortController().signal,
      fetchImpl,
    })).rejects.toThrow(/web-access base URL looks misconfigured.*HTML/);
  });

  it('keeps raw content disabled and omits the content field by default', async () => {
    const fetchImpl = vi.fn(async (_url: RequestInfo | URL, _init?: RequestInit) => jsonResponse({ results: DOMAIN_FIXTURES }));

    const results = await tavilyProvider.search('public data', {
      maxResults: 3,
      apiKey: 'key',
      signal: new AbortController().signal,
      fetchImpl,
    });

    const [, init] = fetchImpl.mock.calls[0];
    expect(JSON.parse((init as RequestInit).body as string)).toMatchObject({ include_raw_content: false });
    expect(results.every(r => r.content === undefined)).toBe(true);
  });

  it('requests text extraction and returns capped per-result content when includeContent is set', async () => {
    const longText = 'row data '.repeat(1000); // 9000 chars, well past the cap
    const fetchImpl = vi.fn(async (_url: RequestInfo | URL, _init?: RequestInit) => jsonResponse({
      results: [
        {
          title: 'Quarterly filing',
          url: 'https://finance.example/reports/q1',
          content: 'Revenue table and notes.',
          raw_content: longText,
        },
        {
          title: 'Match schedule',
          url: 'https://sports.example/schedule',
          content: 'Fixtures and scores.',
          raw_content: null,
        },
      ],
    }));

    const results = await tavilyProvider.search('public data', {
      maxResults: 2,
      apiKey: 'key',
      includeContent: true,
      signal: new AbortController().signal,
      fetchImpl,
    });

    const [, init] = fetchImpl.mock.calls[0];
    expect(JSON.parse((init as RequestInit).body as string)).toMatchObject({ include_raw_content: 'text' });
    expect(results[0].content!.startsWith(longText.slice(0, MAX_RESULT_CONTENT_CHARS))).toBe(true);
    expect(results[0].content!.endsWith('[truncated: page continues beyond this point]')).toBe(true);
    expect(results[0].snippet).toBe('Revenue table and notes.');
    expect(results[1].content).toBeUndefined();
    expect(results[1].snippet).toBe('Fixtures and scores.');
  });
});

describe('Google CSE adapter', () => {
  it('parses documented items, sends the key as a header, and caps num at 10', async () => {
    const fetchImpl = vi.fn(async (_url: RequestInfo | URL, _init?: RequestInit) => jsonResponse({
      items: DOMAIN_FIXTURES.map(f => ({ title: f.title, link: f.url, snippet: f.content })),
    }));

    const results = await googleCseProvider.search('public data', {
      maxResults: 10,
      apiKey: 'gkey',
      engineId: 'engine123',
      signal: new AbortController().signal,
      fetchImpl,
    });

    expect(results).toHaveLength(3);
    expect(results[0]).toEqual({
      title: 'Quarterly filing',
      url: 'https://finance.example/reports/q1',
      snippet: 'Revenue table and notes.',
    });
    const [url, init] = fetchImpl.mock.calls[0];
    expect(String(url)).toContain('cx=engine123');
    expect(String(url)).toContain('num=10');
    expect(String(url)).not.toContain('gkey');
    expect((init as RequestInit).headers).toMatchObject({ 'X-Goog-Api-Key': 'gkey' });
  });

  it('requires an engine id before issuing a request', async () => {
    const fetchImpl = vi.fn();
    await expect(googleCseProvider.search('public data', {
      maxResults: 5,
      apiKey: 'gkey',
      signal: new AbortController().signal,
      fetchImpl,
    })).rejects.toBeInstanceOf(ToolValidationError);
    expect(fetchImpl).not.toHaveBeenCalled();
  });
});

describe('Jina adapter', () => {
  it('parses documented data items and authenticates with a bearer header', async () => {
    const fetchImpl = vi.fn(async (_url: RequestInfo | URL, _init?: RequestInit) => jsonResponse({
      data: DOMAIN_FIXTURES.map(f => ({ title: f.title, url: f.url, description: f.content })),
    }));

    const results = await jinaProvider.search('public data', {
      maxResults: 3,
      apiKey: 'jkey',
      signal: new AbortController().signal,
      fetchImpl,
    });

    expect(results.map(r => r.url)).toEqual(DOMAIN_FIXTURES.map(f => f.url));
    const [, init] = fetchImpl.mock.calls[0];
    expect((init as RequestInit).headers).toMatchObject({ 'Authorization': 'Bearer jkey' });
  });
});

describe('SearXNG adapter', () => {
  it('parses documented results keylessly from a user-supplied instance URL', async () => {
    const fetchImpl = vi.fn(async (_url: RequestInfo | URL, _init?: RequestInit) => jsonResponse({
      results: DOMAIN_FIXTURES.map(f => ({ title: f.title, url: f.url, content: f.content })),
    }));

    const results = await searxngProvider.search('public data', {
      maxResults: 3,
      apiKey: '',
      baseUrl: 'http://localhost:8888/search',
      signal: new AbortController().signal,
      fetchImpl,
    });

    expect(results).toHaveLength(3);
    const [url] = fetchImpl.mock.calls[0];
    expect(String(url)).toContain('http://localhost:8888/search');
    expect(String(url)).toContain('format=json');
  });
});

describe('Wikipedia adapter', () => {
  it('parses documented search hits keylessly, building article URLs and stripping snippet HTML', async () => {
    const fetchImpl = vi.fn(async (_url: RequestInfo | URL, _init?: RequestInit) => jsonResponse({
      query: {
        search: [
          { title: 'Compound interest', snippet: 'the addition of <span class="searchmatch">interest</span> to principal' },
          { title: 'Gross domestic product', snippet: 'monetary measure of market value' },
        ],
      },
    }));

    const results = await wikipediaProvider.search('public data', {
      maxResults: 2,
      apiKey: '',
      signal: new AbortController().signal,
      fetchImpl,
    });

    expect(results).toEqual([
      {
        title: 'Compound interest',
        url: 'https://en.wikipedia.org/wiki/Compound_interest',
        snippet: 'the addition of interest to principal',
      },
      {
        title: 'Gross domestic product',
        url: 'https://en.wikipedia.org/wiki/Gross_domestic_product',
        snippet: 'monetary measure of market value',
      },
    ]);
    const [url] = fetchImpl.mock.calls[0];
    expect(String(url)).toContain('origin=*');
  });

  it('fetches intro extracts for all results when includeContent is set', async () => {
    const fetchImpl = vi.fn(async (url: RequestInfo | URL, _init?: RequestInit) => {
      if (String(url).includes('prop=extracts')) {
        return jsonResponse({
          query: {
            pages: {
              '1234': { title: 'Compound interest', extract: 'Compound interest is the addition of interest to principal.' },
            },
          },
        });
      }
      return jsonResponse({
        query: {
          search: [{ title: 'Compound interest', snippet: 'the addition of <span>interest</span> to principal' }],
        },
      });
    });

    const results = await wikipediaProvider.search('compound interest', {
      maxResults: 1,
      apiKey: '',
      includeContent: true,
      signal: new AbortController().signal,
      fetchImpl,
    });

    expect(results).toHaveLength(1);
    expect(results[0].content).toBe('Compound interest is the addition of interest to principal.');
    expect(results[0].snippet).toBe('the addition of interest to principal');
    expect(fetchImpl).toHaveBeenCalledTimes(2);
    const [extractsUrl] = fetchImpl.mock.calls[1];
    expect(String(extractsUrl)).toContain('prop=extracts');
    expect(String(extractsUrl)).toContain('explaintext=1');
    expect(String(extractsUrl)).toContain('exintro=1');
    expect(String(extractsUrl)).toContain('origin=*');
    expect(String(extractsUrl)).toContain('Compound+interest');
  });

  it('caps extract content at MAX_RESULT_CONTENT_CHARS', async () => {
    const longExtract = 'word '.repeat(2000);
    const fetchImpl = vi.fn(async (url: RequestInfo | URL, _init?: RequestInit) => {
      if (String(url).includes('prop=extracts')) {
        return jsonResponse({
          query: { pages: { '1': { title: 'Compound interest', extract: longExtract } } },
        });
      }
      return jsonResponse({
        query: { search: [{ title: 'Compound interest', snippet: 'snippet' }] },
      });
    });

    const results = await wikipediaProvider.search('compound interest', {
      maxResults: 1,
      apiKey: '',
      includeContent: true,
      signal: new AbortController().signal,
      fetchImpl,
    });

    expect(results[0].content!.startsWith(longExtract.slice(0, MAX_RESULT_CONTENT_CHARS))).toBe(true);
    expect(results[0].content!.endsWith('[truncated: article continues beyond this point]')).toBe(true);
  });

  it('falls back to snippets only if the extracts call fails', async () => {
    let callCount = 0;
    const fetchImpl = vi.fn(async (_url: RequestInfo | URL, _init?: RequestInit) => {
      callCount++;
      if (callCount === 2) throw new Error('network error');
      return jsonResponse({
        query: { search: [{ title: 'Compound interest', snippet: 'the addition of interest to principal' }] },
      });
    });

    const results = await wikipediaProvider.search('compound interest', {
      maxResults: 1,
      apiKey: '',
      includeContent: true,
      signal: new AbortController().signal,
      fetchImpl,
    });

    expect(results).toHaveLength(1);
    expect(results[0].content).toBeUndefined();
    expect(results[0].snippet).toBe('the addition of interest to principal');
  });

  it('skips the extracts call when includeContent is not set', async () => {
    const fetchImpl = vi.fn(async (_url: RequestInfo | URL, _init?: RequestInit) => jsonResponse({
      query: { search: [{ title: 'Compound interest', snippet: 'the addition of interest to principal' }] },
    }));

    const results = await wikipediaProvider.search('compound interest', {
      maxResults: 1,
      apiKey: '',
      signal: new AbortController().signal,
      fetchImpl,
    });

    expect(results[0].content).toBeUndefined();
    expect(fetchImpl).toHaveBeenCalledTimes(1);
  });
});

describe('Wikidata adapter', () => {
  it('parses entity hits keylessly, building Wikidata URLs and optional content summaries', async () => {
    const fetchImpl = vi.fn(async (_url: RequestInfo | URL, _init?: RequestInit) => jsonResponse({
      search: [
        {
          id: 'Q2',
          title: 'Q2',
          label: 'Earth',
          description: 'third planet from the Sun',
          concepturi: 'https://www.wikidata.org/wiki/Q2',
          aliases: ['Terra', 'World'],
        },
      ],
    }));

    const results = await wikidataProvider.search('earth', {
      maxResults: 1,
      apiKey: '',
      includeContent: true,
      signal: new AbortController().signal,
      fetchImpl,
    });

    expect(results).toEqual([{
      title: 'Earth (Q2)',
      url: 'https://www.wikidata.org/wiki/Q2',
      snippet: 'Q2 - third planet from the Sun',
      content: [
        'Entity: Earth',
        'ID: Q2',
        'Description: third planet from the Sun',
        'Aliases: Terra, World',
        'Wikidata URL: https://www.wikidata.org/wiki/Q2',
      ].join('\n'),
    }]);
    const [url] = fetchImpl.mock.calls[0];
    expect(String(url)).toContain('action=wbsearchentities');
    expect(String(url)).toContain('origin=*');
    expect(String(url)).toContain('language=en');
  });
});

describe('World Bank adapter', () => {
  it('ranks indicator metadata and returns API-ready data URLs', async () => {
    const fetchImpl = vi.fn(async (_url: RequestInfo | URL, _init?: RequestInit) => jsonResponse([
      { page: 1, pages: 1, per_page: '20000', total: 3 },
      [
        {
          id: 'SP.POP.TOTL',
          name: 'Population, total',
          sourceNote: 'Total population is based on the de facto definition of population.',
          sourceOrganization: 'World Bank staff estimates',
          topics: [{ value: 'Health' }],
        },
        {
          id: 'NY.GDP.MKTP.CD',
          name: 'GDP (current US$)',
          sourceNote: 'GDP at purchaser prices is the sum of gross value added.',
          sourceOrganization: 'World Bank national accounts data',
          topics: [{ value: 'Economic Policy & Debt' }],
        },
        {
          id: 'EN.ATM.CO2E.PC',
          name: 'CO2 emissions (metric tons per capita)',
          sourceNote: 'Carbon dioxide emissions are those stemming from fossil fuels.',
          topics: [{ value: 'Environment' }],
        },
      ],
    ]));

    const results = await worldBankProvider.search('GDP current', {
      maxResults: 2,
      apiKey: '',
      includeContent: true,
      signal: new AbortController().signal,
      fetchImpl,
    });

    expect(results).toHaveLength(1);
    expect(results[0].title).toBe('GDP (current US$) (NY.GDP.MKTP.CD)');
    expect(results[0].url).toBe('https://api.worldbank.org/v2/country/all/indicator/NY.GDP.MKTP.CD?format=json&per_page=20000');
    expect(results[0].snippet).toContain('Economic Policy & Debt');
    expect(results[0].content).toContain('Data API URL: https://api.worldbank.org/v2/country/all/indicator/NY.GDP.MKTP.CD?format=json&per_page=20000');
    const [url] = fetchImpl.mock.calls[0];
    expect(String(url)).toBe('https://api.worldbank.org/v2/indicator?format=json&per_page=20000');
  });
});

describe('CKAN adapter', () => {
  it('parses public package_search results and prefers resource URLs', async () => {
    const fetchImpl = vi.fn(async (_url: RequestInfo | URL, _init?: RequestInit) => jsonResponse({
      success: true,
      result: {
        results: [{
          id: 'pkg1',
          name: 'air-quality',
          title: 'Air Quality Measurements',
          notes: 'Hourly readings.',
          organization: { title: 'Environment Agency' },
          resources: [{ name: 'CSV', format: 'CSV', url: 'https://public.example/air.csv' }],
          metadata_modified: '2026-01-02T03:04:05',
        }],
      },
    }));

    const results = await ckanProvider.search('air quality', {
      maxResults: 1,
      apiKey: '',
      includeContent: true,
      signal: new AbortController().signal,
      fetchImpl,
    });

    expect(results[0].url).toBe('https://public.example/air.csv');
    expect(results[0].snippet).toContain('Environment Agency');
    expect(results[0].content).toContain('Package URL: https://data.gov.au/data/dataset/air-quality');
    expect(String(fetchImpl.mock.calls[0][0])).toContain('package_search');
  });

  it('derives dataset page URLs from a custom CKAN base URL', async () => {
    const fetchImpl = vi.fn(async (_url: RequestInfo | URL, _init?: RequestInit) => jsonResponse({
      success: true,
      result: {
        results: [{ id: 'pkg1', name: 'air-quality', title: 'Air Quality Measurements' }],
      },
    }));

    const results = await ckanProvider.search('air quality', {
      maxResults: 1,
      apiKey: '',
      baseUrl: 'https://public.example/api/3/action/package_search',
      signal: new AbortController().signal,
      fetchImpl,
    });

    expect(results[0].url).toBe('https://public.example/dataset/air-quality');
    expect(String(fetchImpl.mock.calls[0][0])).toContain('https://public.example/api/3/action/package_search');
  });
});

describe('data.gov.sg adapter', () => {
  it('ranks listed datasets and returns metadata/download API URLs', async () => {
    const fetchImpl = vi.fn(async (_url: RequestInfo | URL, _init?: RequestInit) => jsonResponse({
      code: 0,
      data: {
        datasets: [
          { datasetId: 'd_traffic', name: 'Traffic Images', format: 'API', managedByAgencyName: 'LTA' },
          { datasetId: 'd_rainfall', name: 'Rainfall Across Singapore', format: 'CSV', managedByAgencyName: 'NEA' },
        ],
      },
    }));

    const results = await dataGovSgProvider.search('rainfall', {
      maxResults: 1,
      apiKey: '',
      includeContent: true,
      signal: new AbortController().signal,
      fetchImpl,
    });

    expect(results[0].title).toBe('Rainfall Across Singapore (d_rainfall)');
    expect(results[0].url).toBe('https://api-production.data.gov.sg/v2/public/api/datasets/d_rainfall/metadata');
    expect(results[0].content).toContain('Initiate download URL: https://api-open.data.gov.sg/v1/public/api/datasets/d_rainfall/initiate-download');
  });
});

describe('data.gov.my adapter', () => {
  it('builds keyless Malaysia API candidates from the query', async () => {
    const results = await dataGovMyProvider.search('weather warning earthquake', {
      maxResults: 3,
      apiKey: '',
      includeContent: true,
      signal: new AbortController().signal,
      fetchImpl: vi.fn(),
    });

    expect(results.map(r => r.url)).toContain('https://api.data.gov.my/weather/forecast?limit=100');
    expect(results.map(r => r.url)).toContain('https://api.data.gov.my/weather/warning/earthquake?limit=100');
    expect(results[2].url).toBe('https://api.data.gov.my/data-catalogue?id=weather_warning_earthquake&limit=100');
  });
});

describe('IMF DataMapper adapter', () => {
  it('ranks indicators and returns direct DataMapper API URLs', async () => {
    const fetchImpl = vi.fn(async (_url: RequestInfo | URL, _init?: RequestInit) => jsonResponse({
      indicators: {
        NGDP_RPCH: { label: 'Real GDP growth', description: 'Annual percent change', unit: 'Percent' },
        PCPI_IX: { label: 'Consumer price index', description: 'Index' },
      },
    }));

    const results = await imfProvider.search('real gdp growth', {
      maxResults: 1,
      apiKey: '',
      includeContent: true,
      signal: new AbortController().signal,
      fetchImpl,
    });

    expect(results[0].title).toBe('Real GDP growth (NGDP_RPCH)');
    expect(results[0].url).toBe('https://www.imf.org/external/datamapper/api/v1/NGDP_RPCH');
    expect(results[0].content).toContain('Unit: Percent');
  });

  it('falls back to the reader proxy when the direct fetch is CORS-blocked', async () => {
    const indicators = { NGDP_RPCH: { label: 'Real GDP growth', description: 'Annual percent change', unit: 'Percent' } };
    const fetchImpl = vi.fn()
      .mockRejectedValueOnce(new TypeError('Failed to fetch'))
      .mockResolvedValueOnce(jsonResponse({
        code: 200,
        status: 20000,
        data: { title: '', url: 'https://www.imf.org/external/datamapper/api/v1/indicators', content: JSON.stringify({ indicators }) },
      }));

    const results = await imfProvider.search('real gdp growth', {
      maxResults: 1,
      apiKey: '',
      signal: new AbortController().signal,
      fetchImpl,
    });

    expect(String(fetchImpl.mock.calls[1][0])).toBe('https://r.jina.ai/https://www.imf.org/external/datamapper/api/v1/indicators');
    expect(results[0].title).toBe('Real GDP growth (NGDP_RPCH)');
  });

  it('reports both failures when the direct fetch and the reader proxy fail', async () => {
    const fetchImpl = vi.fn()
      .mockRejectedValueOnce(new TypeError('Failed to fetch'))
      .mockRejectedValueOnce(new TypeError('proxy unreachable'));

    await expect(imfProvider.search('gdp', {
      maxResults: 1,
      apiKey: '',
      signal: new AbortController().signal,
      fetchImpl,
    })).rejects.toThrow(/failed directly \(Failed to fetch\) and via the reader proxy \(proxy unreachable\)/);
  });
});

describe('Eurostat adapter', () => {
  it('ranks dataflows and returns statistics API URLs', async () => {
    const fetchImpl = vi.fn(async (_url: RequestInfo | URL, _init?: RequestInit) => jsonResponse({
      dataflows: [
        { id: 'nama_10_gdp', name: 'GDP and main components', description: 'National accounts' },
        { id: 'demo_pjan', name: 'Population on 1 January', description: 'Demography' },
      ],
    }));

    const results = await eurostatProvider.search('gdp', {
      maxResults: 1,
      apiKey: '',
      includeContent: true,
      signal: new AbortController().signal,
      fetchImpl,
    });

    expect(results[0].url).toBe('https://ec.europa.eu/eurostat/api/dissemination/statistics/1.0/data/nama_10_gdp?lang=en');
    expect(results[0].content).toContain('Dataset: GDP and main components');
  });
});

describe('ECB adapter', () => {
  // The ECB dataflow listing only serves SDMX-ML XML; application/json gets HTTP 406.
  // Namespace declarations omitted: the genericity guard rejects the sdmx.org xmlns URLs
  // and the parser matches on prefixed element names, not bound namespaces.
  const sdmxXml = `<?xml version='1.0' encoding='UTF-8'?>` +
    `<mes:Structure>` +
    `<mes:Structures><str:Dataflows>` +
    `<str:Dataflow agencyID="ECB" id="EXR" version="1.0"><com:Name xml:lang="en">Exchange rates &amp; more</com:Name></str:Dataflow>` +
    `<str:Dataflow agencyID="ECB" id="ICP" version="1.0"><com:Name xml:lang="en">Inflation</com:Name></str:Dataflow>` +
    `</str:Dataflows></mes:Structures></mes:Structure>`;

  it('parses the SDMX XML dataflow listing and returns ECB data API URLs', async () => {
    const fetchImpl = vi.fn(async (_url: RequestInfo | URL, _init?: RequestInit) =>
      textResponse(sdmxXml, 'application/vnd.sdmx.structure+xml;version=2.1')
    );

    const results = await ecbProvider.search('exchange rates', {
      maxResults: 1,
      apiKey: '',
      includeContent: true,
      signal: new AbortController().signal,
      fetchImpl,
    });

    const [, init] = fetchImpl.mock.calls[0];
    expect(new Headers((init as RequestInit).headers).get('accept')).toContain('vnd.sdmx.structure+xml');
    expect(results[0].title).toBe('Exchange rates & more (EXR)');
    expect(results[0].url).toBe('https://data-api.ecb.europa.eu/service/data/EXR?format=jsondata');
    expect(results[0].content).toContain('Dataflow: Exchange rates & more');
  });

  it('throws a clear error when the response contains no dataflows', async () => {
    const fetchImpl = vi.fn(async (_url: RequestInfo | URL, _init?: RequestInit) =>
      textResponse('<mes:Structure></mes:Structure>', 'application/xml')
    );

    await expect(ecbProvider.search('exchange rates', {
      maxResults: 1,
      apiKey: '',
      signal: new AbortController().signal,
      fetchImpl,
    })).rejects.toThrow(/did not contain any SDMX dataflows/);
  });
});

describe('Open-Meteo adapter', () => {
  it('builds forecast URLs when coordinates are present', async () => {
    const results = await openMeteoProvider.search('1.3521, 103.8198 rainfall', {
      maxResults: 2,
      apiKey: '',
      includeContent: true,
      signal: new AbortController().signal,
      fetchImpl: vi.fn(),
    });

    expect(results[0].url).toContain('latitude=1.3521');
    expect(results[0].url).toContain('longitude=103.8198');
    expect(results[0].url).toContain('hourly=temperature_2m');
  });
});

describe('UN SDG adapter', () => {
  it('ranks SDG indicators and returns series lookup URLs', async () => {
    const fetchImpl = vi.fn(async (_url: RequestInfo | URL, _init?: RequestInit) => jsonResponse([
      { code: '1.1.1', indicator: 'Proportion of population below the international poverty line', goal: '1', target: '1.1' },
      { code: '7.2.1', indicator: 'Renewable energy share in the total final energy consumption', goal: '7', target: '7.2' },
    ]));

    const results = await unSdgProvider.search('renewable energy', {
      maxResults: 1,
      apiKey: '',
      includeContent: true,
      signal: new AbortController().signal,
      fetchImpl,
    });

    expect(results[0].title).toContain('Renewable energy share');
    expect(results[0].url).toBe('https://unstats.un.org/sdgapi/v1/sdg/Series/List?indicator=7.2.1');
    expect(results[0].content).toContain('Goal: 7');
  });
});

describe('resolveBaseUrl', () => {
  const FALLBACK = 'https://finance.example';

  it('returns fallback for undefined', () => {
    expect(resolveBaseUrl(undefined, FALLBACK)).toBe(FALLBACK);
  });

  it('returns fallback for empty string', () => {
    expect(resolveBaseUrl('', FALLBACK)).toBe(FALLBACK);
  });

  it('returns fallback for a relative path', () => {
    expect(resolveBaseUrl('/api/search', FALLBACK)).toBe(FALLBACK);
  });

  it('returns fallback for a protocol-relative URL', () => {
    expect(resolveBaseUrl('//host/path', FALLBACK)).toBe(FALLBACK);
  });

  it('returns fallback for a non-http scheme', () => {
    expect(resolveBaseUrl('data:text/html,hi', FALLBACK)).toBe(FALLBACK);
  });

  it('returns the provided URL when it is a valid http URL', () => {
    expect(resolveBaseUrl('http://public.example/search', FALLBACK))
      .toBe('http://public.example/search');
  });

  it('returns the provided URL when it is a valid https URL', () => {
    expect(resolveBaseUrl('https://weather.example/api', FALLBACK))
      .toBe('https://weather.example/api');
  });
});

describe('base URL guard — empty/relative baseUrl falls back to provider endpoint', () => {
  it('Tavily: empty baseUrl sends request to the real Tavily endpoint', async () => {
    const fetchImpl = vi.fn(async (_url: RequestInfo | URL, _init?: RequestInit) =>
      jsonResponse({ results: [] })
    );
    await tavilyProvider.search('test', {
      maxResults: 1, apiKey: 'key', baseUrl: '', signal: new AbortController().signal, fetchImpl,
    });
    const [url] = fetchImpl.mock.calls[0];
    expect(String(url)).toBe('https://api.tavily.com/search');
  });

  it('SearXNG: relative baseUrl falls back to endpoint; valid custom URL is preserved', async () => {
    const fetchImpl = vi.fn(async (_url: RequestInfo | URL, _init?: RequestInit) =>
      jsonResponse({ results: [] })
    );
    // Relative URL → falls back to default endpoint
    await searxngProvider.search('test', {
      maxResults: 1, apiKey: '', baseUrl: '/search', signal: new AbortController().signal, fetchImpl,
    });
    expect(String(fetchImpl.mock.calls[0][0])).toContain('localhost:8080');

    fetchImpl.mockClear();
    // Valid absolute custom URL → preserved
    await searxngProvider.search('test', {
      maxResults: 1, apiKey: '', baseUrl: 'http://localhost:9090/search',
      signal: new AbortController().signal, fetchImpl,
    });
    expect(String(fetchImpl.mock.calls[0][0])).toContain('localhost:9090');
  });

  it('providers throw a clear config error when the endpoint returns HTML', async () => {
    const htmlResp = textResponse('<html><body>login required</body></html>');
    for (const [provider, opts] of [
      [tavilyProvider, { apiKey: 'k' }] as const,
      [searxngProvider, { apiKey: '' }] as const,
      [jinaProvider, { apiKey: 'k' }] as const,
      [wikidataProvider, { apiKey: '' }] as const,
      [worldBankProvider, { apiKey: '' }] as const,
      [ckanProvider, { apiKey: '' }] as const,
      [dataGovSgProvider, { apiKey: '' }] as const,
      [ecbProvider, { apiKey: '' }] as const,
      [eurostatProvider, { apiKey: '' }] as const,
      [imfProvider, { apiKey: '' }] as const,
      [unSdgProvider, { apiKey: '' }] as const,
    ] as const) {
      const fetchImpl = vi.fn(async () => htmlResp.clone());
      await expect(
        provider.search('q', { maxResults: 1, ...opts, signal: new AbortController().signal, fetchImpl })
      ).rejects.toThrow(/web-access base URL looks misconfigured.*HTML/);
    }
  });
});

describe('web_search handler', () => {
  it('returns NetworkError through the executor path without fabricated result URLs', async () => {
    const executor = new ToolExecutor(new WorkbookRegistry(), async () => {
      throw new Error('runtime none should not use Excel runner');
    });
    executor.register(WEB_SEARCH, createWebSearchHandler({
      getProvider: () => 'tavily',
      getApiKey: () => 'key',
      fetchImpl: vi.fn(async () => { throw new TypeError('network down'); }),
    }));

    const call: ToolCall = {
      id: 'search_call',
      name: 'web_search',
      arguments: { query: 'anything' },
      workbookId: 'host',
      mutating: false,
    };
    const result = await executor.execute(call, { workbookId: 'host' });

    expect(result.ok).toBe(false);
    expect(result.error?.code).toBe('NetworkError');
    expect(JSON.stringify(result)).not.toMatch(/https?:\/\//);
  });

  it('plumbs include_content through to the provider and surfaces content in the result', async () => {
    const fetchImpl = vi.fn(async (_url: RequestInfo | URL, _init?: RequestInit) => jsonResponse({
      results: [
        {
          title: 'Quarterly filing',
          url: 'https://finance.example/reports/q1',
          content: 'Revenue table and notes.',
          raw_content: 'Full extracted report text.',
        },
      ],
    }));
    const executor = new ToolExecutor(new WorkbookRegistry(), async () => {
      throw new Error('runtime none should not use Excel runner');
    });
    executor.register(WEB_SEARCH, createWebSearchHandler({
      getProvider: () => 'tavily',
      getApiKey: () => 'key',
      fetchImpl,
    }));

    const call: ToolCall = {
      id: 'search_call',
      name: 'web_search',
      arguments: { query: 'public data', include_content: true },
      workbookId: 'host',
      mutating: false,
    };
    const result = await executor.execute(call, { workbookId: 'host' });

    expect(result.ok).toBe(true);
    const [, init] = fetchImpl.mock.calls[0];
    expect(JSON.parse((init as RequestInit).body as string)).toMatchObject({ include_raw_content: 'text' });
    expect((result.data as { results: Array<{ content?: string }> }).results[0].content)
      .toBe('Full extracted report text.');
  });

  it('rejects a non-boolean include_content with a ValidationError', async () => {
    const fetchImpl = vi.fn();
    const executor = new ToolExecutor(new WorkbookRegistry(), async () => {
      throw new Error('runtime none should not use Excel runner');
    });
    executor.register(WEB_SEARCH, createWebSearchHandler({
      getProvider: () => 'tavily',
      getApiKey: () => 'key',
      fetchImpl,
    }));

    const call: ToolCall = {
      id: 'search_call',
      name: 'web_search',
      arguments: { query: 'public data', include_content: 'yes' },
      workbookId: 'host',
      mutating: false,
    };
    const result = await executor.execute(call, { workbookId: 'host' });

    expect(result.ok).toBe(false);
    expect(result.error?.code).toBe('ValidationError');
    expect(fetchImpl).not.toHaveBeenCalled();
  });
});

