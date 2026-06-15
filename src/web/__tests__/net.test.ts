import { beforeEach, describe, expect, it, vi } from 'vitest';
import { clearHostStatusCache, fetchTextWithGuards } from '../net';
import { ToolNetworkError } from '../../workbook/executor';

function abortError(): Error {
  const e = new Error('signal is aborted without reason');
  e.name = 'AbortError';
  return e;
}

function textResponse(body: string): Response {
  return new Response(body, { status: 200, headers: { 'content-type': 'text/plain' } });
}

// fetchImpl whose main (CORS-enforced) request rejects but whose no-cors probe resolves.
function corsBlockedFetch() {
  return vi.fn(async (_url: RequestInfo | URL, init?: RequestInit) => {
    if (init?.mode === 'no-cors') return textResponse('');
    throw new TypeError('Failed to fetch');
  });
}

beforeEach(() => {
  clearHostStatusCache();
});

describe('fetchTextWithGuards failure classification', () => {
  it('reports timeouts as timeouts, without a CORS claim or a probe', async () => {
    const fetchImpl = vi.fn((_url: RequestInfo | URL, init?: RequestInit) =>
      new Promise<Response>((_, reject) => {
        init?.signal?.addEventListener('abort', () => reject(abortError()), { once: true });
      })
    );

    await expect(fetchTextWithGuards('https://public.example/data', { fetchImpl, timeoutMs: 10 }))
      .rejects.toThrow(/timed out after 10 ms/);
    expect(fetchImpl).toHaveBeenCalledTimes(1);
  });

  it('confirms CORS blocks via a no-cors probe and instructs the model not to retry', async () => {
    const fetchImpl = corsBlockedFetch();

    const error = await fetchTextWithGuards('https://public.example/data', { fetchImpl })
      .then(() => null, e => e as Error);

    expect(error).toBeInstanceOf(ToolNetworkError);
    expect(error?.message).toMatch(/Blocked by CORS: public\.example/);
    expect(error?.message).toMatch(/Do not retry/);
    expect(fetchImpl).toHaveBeenCalledTimes(2);
    expect((fetchImpl.mock.calls[1][1] as RequestInit).mode).toBe('no-cors');
  });

  it('fails fast on hosts already confirmed CORS-blocked, without a network attempt', async () => {
    const fetchImpl = corsBlockedFetch();

    await expect(fetchTextWithGuards('https://public.example/data', { fetchImpl }))
      .rejects.toThrow(/Blocked by CORS/);
    expect(fetchImpl).toHaveBeenCalledTimes(2);

    await expect(fetchTextWithGuards('https://public.example/other-path', { fetchImpl }))
      .rejects.toThrow(/Blocked by CORS \(cached\).*Do not retry/);
    expect(fetchImpl).toHaveBeenCalledTimes(2);
  });

  it('classifies unreachable hosts distinctly and does not cache them', async () => {
    const fetchImpl = vi.fn(async () => {
      throw new TypeError('Failed to fetch');
    });

    await expect(fetchTextWithGuards('https://public.example/data', { fetchImpl }))
      .rejects.toThrow(/Network unreachable.*public\.example/);
    expect(fetchImpl).toHaveBeenCalledTimes(2);

    // Transient failure: a later attempt should hit the network again.
    await expect(fetchTextWithGuards('https://public.example/data', { fetchImpl }))
      .rejects.toThrow(/Network unreachable/);
    expect(fetchImpl).toHaveBeenCalledTimes(4);
  });

  it('rethrows parent aborts without converting them to network errors', async () => {
    const parent = new AbortController();
    parent.abort();
    const fetchImpl = vi.fn(async (_url: RequestInfo | URL, init?: RequestInit) => {
      if (init?.signal?.aborted) throw abortError();
      return textResponse('ok');
    });

    const error = await fetchTextWithGuards('https://public.example/data', { fetchImpl, signal: parent.signal })
      .then(() => null, e => e as Error);

    expect(error?.name).toBe('AbortError');
    expect(error).not.toBeInstanceOf(ToolNetworkError);
  });

  it('preserves the body-size ceiling message instead of relabeling it', async () => {
    const fetchImpl = vi.fn(async () => textResponse('x'.repeat(100)));

    await expect(fetchTextWithGuards('https://public.example/data', { fetchImpl, maxBytes: 10 }))
      .rejects.toThrow(/Response body exceeded 10 bytes/);
  });

  it('can return a byte-truncated response when requested by a bounded preview tool', async () => {
    const fetchImpl = vi.fn(async () => textResponse('x'.repeat(100)));

    const response = await fetchTextWithGuards('https://public.example/data', {
      fetchImpl,
      maxBytes: 10,
      truncateAtMaxBytes: true,
    });

    expect(response.text).toBe('x'.repeat(10));
    expect(response.bytesFetched).toBe(10);
    expect(response.truncatedByBytes).toBe(true);
  });
});
