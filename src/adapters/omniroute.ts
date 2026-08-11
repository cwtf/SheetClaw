// OmniRoute is a self-hosted, OpenAI-compatible AI gateway. It runs as a local
// process and proxies to upstream providers, so it needs no adapter of its own:
// createAdapter() routes it through OpenAIAdapter like any other compatible API.
// What it does need is the same local-server diagnostics Ollama gets, because
// the two failure modes users hit are identical - the gateway is not running,
// or it is running but will not answer a cross-origin request from the pane.

/** OmniRoute's canonical local port. PORT / API_PORT in its .env override it. */
export const OMNIROUTE_DEFAULT_BASE_URL = 'http://localhost:20128/v1';

export const OMNIROUTE_NOT_RUNNING =
  'No OmniRoute gateway is listening on that address. Start it, or correct the base URL if you changed PORT / API_PORT.';

const BROWSER_ACCESS_PREFIX = 'The OmniRoute gateway is running, but this add-in cannot read it.';

function browserOrigin(): string | undefined {
  if (typeof window === 'undefined') return undefined;
  return window.location?.origin;
}

export function getOmniRouteBrowserAccessHint(origin = browserOrigin()): string {
  const originText = origin ?? '<add-in-origin>';
  return `${BROWSER_ACCESS_PREFIX} The gateway must return an Access-Control-Allow-Origin header for ${originText}; check the CORS and IP allowlist settings on the OmniRoute dashboard's Endpoints page, then retry.`;
}

export function isOmniRouteBrowserAccessError(message: string): boolean {
  return message.includes(BROWSER_ACCESS_PREFIX);
}

/**
 * Distinguish "gateway is down" from "gateway is up but blocked us".
 *
 * A no-cors request is opaque - we cannot read the response - but it only
 * resolves when something actually answered on that origin, which is exactly
 * the signal we need. Returns undefined when the address looks unreachable,
 * leaving the caller's original error as the better message.
 */
export async function diagnoseOmniRouteFailure(baseUrl: string): Promise<string> {
  try {
    await fetch(baseUrl, { method: 'GET', mode: 'no-cors', cache: 'no-store' });
    return getOmniRouteBrowserAccessHint();
  } catch {
    return OMNIROUTE_NOT_RUNNING;
  }
}
