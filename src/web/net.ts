import { ToolNetworkError, ToolValidationError } from '../workbook/executor';

export const FETCH_TIMEOUT_MS = 15_000;
export const MAX_BODY_BYTES = 1_000_000;
const PROBE_TIMEOUT_MS = 5_000;

const ALLOWED_PORTS = new Set(['', '80', '443', '8080']);

// Hosts confirmed CORS-blocked this session (learned at runtime, never hardcoded).
// Lets repeat fetches fail fast instead of burning the full timeout again.
const corsBlockedHosts = new Set<string>();

export function clearHostStatusCache(): void {
  corsBlockedHosts.clear();
}

export interface FetchResponse {
  url: string;
  finalUrl: string;
  status: number;
  statusText: string;
  contentType: string;
  bytesFetched: number;
  truncatedByBytes: boolean;
  text: string;
}

export function validatePublicHttpUrl(value: string): URL {
  let url: URL;
  try {
    url = new URL(value);
  } catch {
    throw new ToolValidationError('URL must be an absolute http(s) URL.');
  }

  if (url.protocol !== 'http:' && url.protocol !== 'https:') {
    throw new ToolValidationError('URL must use http or https.');
  }
  if (url.username || url.password) {
    throw new ToolValidationError('URL must not contain embedded credentials.');
  }
  if (!ALLOWED_PORTS.has(url.port)) {
    throw new ToolValidationError('URL port is not allowed. Use 80, 443, or 8080.');
  }

  const host = url.hostname.toLowerCase();
  if (!host || host === 'localhost' || host.endsWith('.localhost') || host.endsWith('.local')) {
    throw new ToolValidationError('URL host must be public, not localhost or local network.');
  }
  if (isBlockedIp(host)) {
    throw new ToolValidationError('URL host must be public, not private, loopback, or link-local.');
  }

  return url;
}

export async function fetchTextWithGuards(
  urlValue: string,
  opts: {
    fetchImpl?: typeof fetch;
    signal?: AbortSignal;
    timeoutMs?: number;
    maxBytes?: number;
    truncateAtMaxBytes?: boolean;
  } = {}
): Promise<FetchResponse> {
  const original = validatePublicHttpUrl(urlValue);
  const fetchImpl = opts.fetchImpl ?? fetch;
  const timeoutMs = opts.timeoutMs ?? FETCH_TIMEOUT_MS;
  const maxBytes = opts.maxBytes ?? MAX_BODY_BYTES;
  const host = original.hostname.toLowerCase();

  if (corsBlockedHosts.has(host)) {
    throw new ToolNetworkError(
      `Blocked by CORS (cached): ${host} was already confirmed CORS-blocked earlier in this session. ` +
      'Do not retry this host; choose a different source or ask the user how to proceed.'
    );
  }

  const controller = new AbortController();
  const timeout = globalThis.setTimeout(() => controller.abort(), timeoutMs);

  const abortFromParent = () => controller.abort();
  if (opts.signal?.aborted) controller.abort();
  opts.signal?.addEventListener('abort', abortFromParent, { once: true });

  try {
    const response = await fetchImpl(original.toString(), {
      method: 'GET',
      redirect: 'follow',
      signal: controller.signal,
    });
    const finalUrl = response.url || original.toString();
    validatePublicHttpUrl(finalUrl);

    const bytes = await readResponseBytes(response, maxBytes, opts.truncateAtMaxBytes ?? false);
    return {
      url: original.toString(),
      finalUrl,
      status: response.status,
      statusText: response.statusText,
      contentType: response.headers.get('content-type') ?? '',
      bytesFetched: bytes.byteLength,
      truncatedByBytes: bytes.truncated,
      text: new TextDecoder().decode(bytes),
    };
  } catch (e) {
    if (e instanceof ToolValidationError || e instanceof ToolNetworkError) throw e;

    // Parent abort (user stop): preserve abort semantics instead of inventing a network error.
    if (opts.signal?.aborted) throw e;

    if (isAbortError(e)) {
      throw new ToolNetworkError(
        `Request timed out after ${timeoutMs} ms. The host did not respond in time; one retry or a different source is reasonable.`
      );
    }

    if (e instanceof TypeError) {
      // fetch rejects with TypeError for both CORS blocks and DNS/offline failures.
      // A no-cors probe disambiguates: it ignores CORS, so it resolves iff the host is reachable.
      const reachable = await probeReachability(original, fetchImpl);
      if (reachable) {
        corsBlockedHosts.add(host);
        throw new ToolNetworkError(
          `Blocked by CORS: ${host} is reachable but does not allow browser requests from the add-in. ` +
          'Do not retry this host; choose a different source, or the user can enable the reader fallback in Settings.'
        );
      }
      throw new ToolNetworkError(
        `Network unreachable: could not connect to ${host} (DNS failure or no connectivity). Do not retry immediately.`
      );
    }

    const message = e instanceof Error ? e.message : String(e);
    throw new ToolNetworkError(`Network request failed: ${message}`);
  } finally {
    globalThis.clearTimeout(timeout);
    opts.signal?.removeEventListener('abort', abortFromParent);
  }
}

function isAbortError(e: unknown): boolean {
  return typeof e === 'object' && e !== null && (e as { name?: unknown }).name === 'AbortError';
}

async function probeReachability(url: URL, fetchImpl: typeof fetch): Promise<boolean> {
  const controller = new AbortController();
  const timeout = globalThis.setTimeout(() => controller.abort(), PROBE_TIMEOUT_MS);
  try {
    const response = await fetchImpl(url.toString(), {
      method: 'GET',
      mode: 'no-cors',
      redirect: 'follow',
      signal: controller.signal,
    });
    response.body?.cancel().catch(() => {});
    return true;
  } catch {
    return false;
  } finally {
    globalThis.clearTimeout(timeout);
  }
}

type ReadBytesResult = Uint8Array & { truncated: boolean };

async function readResponseBytes(response: Response, maxBytes: number, truncateAtMaxBytes: boolean): Promise<ReadBytesResult> {
  if (!response.body) {
    const buffer = await response.arrayBuffer();
    if (buffer.byteLength > maxBytes) {
      if (truncateAtMaxBytes) return withTruncatedFlag(new Uint8Array(buffer.slice(0, maxBytes)), true);
      throw new ToolNetworkError(`Response body exceeded ${maxBytes} bytes.`);
    }
    return withTruncatedFlag(new Uint8Array(buffer), false);
  }

  const reader = response.body.getReader();
  const chunks: Uint8Array[] = [];
  let total = 0;
  let truncated = false;

  while (true) {
    const { done, value } = await reader.read();
    if (done) break;
    if (!value) continue;
    const remaining = maxBytes - total;
    if (value.byteLength > remaining) {
      if (truncateAtMaxBytes && remaining > 0) {
        chunks.push(value.slice(0, remaining));
        total = maxBytes;
      }
      reader.cancel().catch(() => {});
      if (truncateAtMaxBytes) {
        truncated = true;
        break;
      }
      throw new ToolNetworkError(`Response body exceeded ${maxBytes} bytes.`);
    }
    total += value.byteLength;
    chunks.push(value);
  }

  const out = new Uint8Array(total);
  let offset = 0;
  for (const chunk of chunks) {
    out.set(chunk, offset);
    offset += chunk.byteLength;
  }
  return withTruncatedFlag(out, truncated);
}

function withTruncatedFlag(bytes: Uint8Array, truncated: boolean): ReadBytesResult {
  return Object.assign(bytes, { truncated });
}

function isBlockedIp(host: string): boolean {
  return isBlockedIpv4(host) || isBlockedIpv6(host);
}

function isBlockedIpv4(host: string): boolean {
  const parts = host.split('.');
  if (parts.length !== 4) return false;
  const nums = parts.map(p => Number(p));
  if (nums.some(n => !Number.isInteger(n) || n < 0 || n > 255)) return false;
  const [a, b] = nums;
  return (
    a === 0 ||
    a === 10 ||
    a === 127 ||
    (a === 100 && b >= 64 && b <= 127) ||
    (a === 169 && b === 254) ||
    (a === 172 && b >= 16 && b <= 31) ||
    (a === 192 && b === 168) ||
    (a === 198 && (b === 18 || b === 19))
  );
}

function isBlockedIpv6(host: string): boolean {
  const normalized = host.replace(/^\[|\]$/g, '').toLowerCase();
  return (
    normalized === '::1' ||
    normalized === '::' ||
    normalized.startsWith('fe80:') ||
    normalized.startsWith('fc') ||
    normalized.startsWith('fd')
  );
}
