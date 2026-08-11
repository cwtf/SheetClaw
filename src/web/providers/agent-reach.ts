import { ToolNetworkError } from '../../workbook/executor';
import { resolveBaseUrl, type SearchProviderAdapter, type SearchResult } from './index';

/**
 * Platform search via a local Agent-Reach bridge (tools/agent-reach-bridge).
 *
 * Agent-Reach itself is CLI-only, so the pane never talks to it directly - the
 * bridge implements Agent-Reach's routing table over HTTP and shells out to the
 * upstream readers it provisions. Self-hosted and keyless, like SearXNG, which
 * also keeps it out of the keyless bundle (that bundle is remote catalogues; a
 * localhost service that may not be running has no business in the default fan-out).
 */

/** Platforms the bridge can route. Kept in sync with PLATFORMS in server.mjs. */
export const AGENT_REACH_PLATFORMS = [
  'youtube',
  'github',
  'reddit',
  'twitter',
  'bilibili',
  'xiaohongshu',
] as const;

export type AgentReachPlatform = typeof AGENT_REACH_PLATFORMS[number];

interface BridgeSearchResponse {
  query?: unknown;
  platform?: unknown;
  results?: Array<{
    title?: unknown;
    url?: unknown;
    snippet?: unknown;
    publishedAt?: unknown;
  }>;
}

/**
 * Pull a leading `platform:` selector off the query.
 *
 * web_search's `source` parameter is validated against the keyless catalogue
 * ids, so it cannot carry a platform. An inline prefix keeps platform targeting
 * available to the model without widening that shared enum for one provider.
 */
export function splitPlatformPrefix(query: string): { platform?: AgentReachPlatform; query: string } {
  const match = /^([a-z]+):\s*(.+)$/i.exec(query.trim());
  if (!match) return { query: query.trim() };
  const candidate = match[1].toLowerCase() as AgentReachPlatform;
  if (!(AGENT_REACH_PLATFORMS as readonly string[]).includes(candidate)) return { query: query.trim() };
  return { platform: candidate, query: match[2].trim() };
}

export const agentReachProvider: SearchProviderAdapter = {
  id: 'agent-reach',
  label: 'Agent-Reach (local bridge)',
  requiresKey: false,
  selfHosted: true,
  endpoint: 'http://localhost:8788',
  signupUrl: 'https://github.com/Panniantong/Agent-Reach',

  async search(query, opts): Promise<SearchResult[]> {
    const fetchImpl = opts.fetchImpl ?? fetch;
    const { platform, query: bare } = splitPlatformPrefix(query);

    const url = new URL('/search', resolveBaseUrl(opts.baseUrl, this.endpoint));
    url.searchParams.set('q', bare);
    url.searchParams.set('limit', String(opts.maxResults));
    if (platform) url.searchParams.set('platform', platform);

    let response: Response;
    try {
      response = await fetchImpl(url.toString(), {
        method: 'GET',
        headers: { Accept: 'application/json' },
        signal: opts.signal,
      });
    } catch (e) {
      const message = e instanceof Error ? e.message : String(e);
      throw new ToolNetworkError(
        `agent-reach bridge request failed: ${message}. Start it with \`npm run agent-reach-bridge\` and confirm the URL in Settings.`
      );
    }

    if (response.status === 501) {
      throw new ToolNetworkError(
        `agent-reach bridge: ${await errorText(response)} Prefix the query with a supported platform, e.g. "youtube: ${bare}".`
      );
    }
    if (!response.ok) {
      throw new ToolNetworkError(`agent-reach bridge returned HTTP ${response.status}. ${await errorText(response)}`);
    }

    let json: BridgeSearchResponse;
    try {
      json = await response.json() as BridgeSearchResponse;
    } catch {
      throw new ToolNetworkError('agent-reach bridge response was not valid JSON');
    }

    return (json.results ?? [])
      .map(normalizeResult)
      .filter((result): result is SearchResult => !!result);
  },
};

async function errorText(response: Response): Promise<string> {
  try {
    const body = await response.json() as { error?: unknown };
    return typeof body.error === 'string' ? body.error : '';
  } catch {
    return '';
  }
}

function normalizeResult(item: NonNullable<BridgeSearchResponse['results']>[number]): SearchResult | null {
  if (typeof item.title !== 'string' || typeof item.url !== 'string' || !item.url) return null;
  return {
    title: item.title,
    url: item.url,
    ...(typeof item.snippet === 'string' ? { snippet: item.snippet } : {}),
    ...(typeof item.publishedAt === 'string' ? { publishedAt: item.publishedAt } : {}),
  };
}
