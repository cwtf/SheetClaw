import { ToolNetworkError } from '../workbook/executor';
import { agentReachProvider, resolveBaseUrl } from './providers';

/**
 * fetch_url backend for platforms a plain cross-origin fetch cannot read.
 *
 * X, Reddit, YouTube and friends either block the request outright or return a
 * JavaScript shell with no content in it, so `fetch_url` on those hosts is a
 * guaranteed miss. When the Agent-Reach bridge is configured, route those hosts
 * through it first; everything else keeps the existing direct → reader path.
 */

/**
 * Hosts the bridge routes, `www.` stripped. Kept in sync with PLATFORMS in
 * server.mjs, and exported so the genericity guard can treat them as a declared
 * surface rather than stray hardcoded sites.
 */
export const PLATFORM_HOSTS: Record<string, string> = {
  'youtube.com': 'youtube',
  'youtu.be': 'youtube',
  'm.youtube.com': 'youtube',
  'github.com': 'github',
  'reddit.com': 'reddit',
  'old.reddit.com': 'reddit',
  'x.com': 'twitter',
  'twitter.com': 'twitter',
  'bilibili.com': 'bilibili',
  'b23.tv': 'bilibili',
  'xiaohongshu.com': 'xiaohongshu',
};

export interface AgentReachReadResult {
  url: string;
  platform: string;
  tool: string;
  title: string;
  text: string;
  truncated: boolean;
}

/** The platform id for a URL, or null when no bridge route covers that host. */
export function agentReachPlatformFor(url: string): string | null {
  let hostname: string;
  try {
    hostname = new URL(url).hostname.toLowerCase().replace(/^www\./, '');
  } catch {
    return null;
  }
  return PLATFORM_HOSTS[hostname] ?? null;
}

export async function readViaAgentReach(
  url: string,
  opts: { baseUrl?: string; signal?: AbortSignal; fetchImpl?: typeof fetch }
): Promise<AgentReachReadResult> {
  const fetchImpl = opts.fetchImpl ?? fetch;
  const endpoint = new URL('/read', resolveBaseUrl(opts.baseUrl, agentReachProvider.endpoint));
  endpoint.searchParams.set('url', url);

  let response: Response;
  try {
    response = await fetchImpl(endpoint.toString(), {
      method: 'GET',
      headers: { Accept: 'application/json' },
      signal: opts.signal,
    });
  } catch (e) {
    const message = e instanceof Error ? e.message : String(e);
    throw new ToolNetworkError(`agent-reach bridge is unreachable: ${message}`);
  }

  if (!response.ok) {
    let detail = '';
    try {
      const body = await response.json() as { error?: unknown };
      if (typeof body.error === 'string') detail = ` ${body.error}`;
    } catch {
      // keep the bare status
    }
    throw new ToolNetworkError(`agent-reach bridge returned HTTP ${response.status}.${detail}`);
  }

  let body: Record<string, unknown>;
  try {
    body = await response.json() as Record<string, unknown>;
  } catch {
    throw new ToolNetworkError('agent-reach bridge response was not valid JSON');
  }

  if (typeof body.text !== 'string') {
    throw new ToolNetworkError('agent-reach bridge response had no text field');
  }

  return {
    url,
    platform: typeof body.platform === 'string' ? body.platform : 'unknown',
    tool: typeof body.tool === 'string' ? body.tool : 'unknown',
    title: typeof body.title === 'string' ? body.title : '',
    text: body.text,
    truncated: body.truncated === true,
  };
}
