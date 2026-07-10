import { tavilyProvider } from './tavily';
import { googleCseProvider } from './google-cse';
import { jinaProvider } from './jina';
import { searxngProvider } from './searxng';
import { wikipediaProvider } from './wikipedia';
import { wikidataProvider } from './wikidata';
import { worldBankProvider } from './worldbank';
import { ckanProvider } from './ckan';
import { dataGovMyProvider } from './data-gov-my';
import { dataGovSgProvider } from './data-gov-sg';
import { ecbProvider } from './ecb';
import { eurostatProvider } from './eurostat';
import { imfProvider } from './imf';
import { openMeteoProvider } from './open-meteo';
import { unSdgProvider } from './un-sdg';
import { keylessBundleProvider, KEYLESS_BUNDLE_ID, type KeylessBundleId } from './keyless';

export { keylessBundleProvider, KEYLESS_BUNDLE_ID, KEYLESS_SOURCE_HINTS, keylessSourceIds, type KeylessBundleId } from './keyless';

export type SearchProviderId =
  | 'tavily'
  | 'google-cse'
  | 'jina'
  | 'searxng'
  | 'wikipedia'
  | 'wikidata'
  | 'worldbank'
  | 'ckan'
  | 'data-gov-my'
  | 'data-gov-sg'
  | 'ecb'
  | 'eurostat'
  | 'imf'
  | 'open-meteo'
  | 'un-sdg';
export type WebAccessProvider = 'none' | SearchProviderId | KeylessBundleId;

export interface SearchResult {
  title: string;
  url: string;
  snippet?: string;
  publishedAt?: string;
  /** Extracted page text, present only when includeContent was requested and the provider supports it. */
  content?: string;
  /** Which keyless source produced this result; set only by the keyless bundle provider. */
  source?: SearchProviderId;
}

/** Cap on extracted page text per result so tool output stays bounded. */
export const MAX_RESULT_CONTENT_CHARS = 4000;

export interface SearchProviderAdapter {
  id: SearchProviderId | KeylessBundleId;
  label: string;
  requiresKey: boolean;
  /** Google CSE needs a Programmable Search Engine id (cx) in addition to the API key. */
  requiresEngineId?: boolean;
  /** Needs a user-supplied instance URL; listed individually and excluded from the keyless bundle. */
  selfHosted?: boolean;
  endpoint: string;
  signupUrl: string;
  search(
    query: string,
    opts: {
      maxResults: number;
      apiKey: string;
      baseUrl?: string;
      engineId?: string;
      /** Also return extracted page text per result. Providers without support ignore this. */
      includeContent?: boolean;
      /** Keyless source to target; honoured only by the keyless bundle provider. */
      source?: string;
      signal: AbortSignal;
      fetchImpl?: typeof fetch;
    }
  ): Promise<SearchResult[]>;
}

export const SEARCH_PROVIDERS: Record<SearchProviderId, SearchProviderAdapter> = {
  tavily: tavilyProvider,
  'google-cse': googleCseProvider,
  jina: jinaProvider,
  searxng: searxngProvider,
  wikipedia: wikipediaProvider,
  wikidata: wikidataProvider,
  worldbank: worldBankProvider,
  ckan: ckanProvider,
  'data-gov-my': dataGovMyProvider,
  'data-gov-sg': dataGovSgProvider,
  ecb: ecbProvider,
  eurostat: eurostatProvider,
  imf: imfProvider,
  'open-meteo': openMeteoProvider,
  'un-sdg': unSdgProvider,
};

export const SEARCH_PROVIDER_IDS = Object.keys(SEARCH_PROVIDERS) as SearchProviderId[];

export const READER_PROVIDER_ENDPOINT = 'https://r.jina.ai/';

export const PROVIDER_HOST_ALLOWLIST = Object.values(SEARCH_PROVIDERS).map(provider => {
  const url = new URL(provider.endpoint);
  return url.hostname;
});

export const PROVIDER_URL_HOST_ALLOWLIST = Object.values(SEARCH_PROVIDERS).flatMap(provider =>
  [provider.endpoint, provider.signupUrl].map(value => new URL(value).hostname)
).concat(new URL(READER_PROVIDER_ENDPOINT).hostname);

export function getSearchProvider(id: WebAccessProvider): SearchProviderAdapter | null {
  if (id === 'none') return null;
  if (id === KEYLESS_BUNDLE_ID) return keylessBundleProvider;
  return SEARCH_PROVIDERS[id] ?? null;
}

export function isKeylessSearchProvider(id: WebAccessProvider): boolean {
  const provider = getSearchProvider(id);
  return provider?.requiresKey === false;
}

/**
 * Returns `fallback` unless `baseUrl` is an absolute http(s) URL.
 * Prevents empty strings or same-origin paths from slipping through `opts.baseUrl ?? this.endpoint`.
 */
export function resolveBaseUrl(baseUrl: string | undefined, fallback: string): string {
  if (!baseUrl) return fallback;
  try {
    const u = new URL(baseUrl);
    if (u.protocol === 'http:' || u.protocol === 'https:') return baseUrl;
  } catch {
    // not a valid URL — fall through
  }
  return fallback;
}
