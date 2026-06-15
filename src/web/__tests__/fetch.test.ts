import { describe, expect, it, vi } from 'vitest';
import { FETCH_URL, createFetchUrlHandler, handleFetchUrlWithOptions } from '../fetch';
import { validatePublicHttpUrl } from '../net';
import { ToolExecutor, ToolValidationError } from '../../workbook/executor';
import { WorkbookRegistry } from '../../workbook/registry';
import type { ToolCall } from '../../types';

const SCOPE = { workbookId: 'host' };

function makeCall(args: Record<string, unknown>): ToolCall {
  return {
    id: 'call_fetch',
    name: 'fetch_url',
    arguments: args,
    workbookId: 'host',
    mutating: false,
  };
}

function textResponse(body: string, init: ResponseInit & { url?: string } = {}): Response {
  const response = new Response(body, init);
  Object.defineProperty(response, 'url', {
    value: init.url ?? 'https://public.example/resource',
  });
  return response;
}

describe('fetch_url caps and previews', () => {
  it('caps oversized JSON at max_chars and returns a preview instead of parsed data', async () => {
    const data = Array.from({ length: 500 }, (_, i) => ({ id: i, value: `value-${i}` }));
    const fetchImpl = vi.fn(async () => textResponse(JSON.stringify(data), {
      status: 200,
      headers: { 'content-type': 'application/json' },
      url: 'https://public.example/data.json',
    }));

    const result = await handleFetchUrlWithOptions({
      url: 'https://public.example/data.json',
      format: 'json',
      max_chars: 1000,
    }, { fetchImpl });

    expect(result).toMatchObject({ format: 'json', truncated: true });
    expect((result as { returnedChars: number }).returnedChars).toBeLessThanOrEqual(1000);
    expect((result as { data?: unknown }).data).toBeUndefined();
    expect((result as { dataPreview: string }).dataPreview.length).toBeLessThanOrEqual(1000);
  });

  it('caps oversized text at max_chars', async () => {
    const fetchImpl = vi.fn(async () => textResponse('x'.repeat(3000), {
      status: 200,
      headers: { 'content-type': 'text/plain' },
    }));

    const result = await handleFetchUrlWithOptions({
      url: 'https://public.example/file.txt',
      format: 'text',
      max_chars: 1000,
    }, { fetchImpl });

    expect(result).toMatchObject({ format: 'text', truncated: true });
    expect((result as { returnedChars: number }).returnedChars).toBeLessThanOrEqual(1000);
    expect((result as { text: string }).text.length).toBeLessThanOrEqual(1000);
  });

  it('caps full CSV text while keeping a non-redundant preview shape', async () => {
    const csv = [
      'id,name',
      ...Array.from({ length: 200 }, (_, i) => `${i},name-${i}`),
    ].join('\n');
    const fetchImpl = vi.fn(async () => textResponse(csv, {
      status: 200,
      headers: { 'content-type': 'text/csv' },
    }));

    const result = await handleFetchUrlWithOptions({
      url: 'https://public.example/table.csv',
      format: 'csv',
      mode: 'full',
      max_chars: 1000,
    }, { fetchImpl });

    expect(result).toMatchObject({ format: 'csv', truncated: true });
    expect((result as { returnedChars: number }).returnedChars).toBeLessThanOrEqual(1000);
    expect((result as { text: string }).text.length).toBeLessThanOrEqual(1000);
    expect((result as { previewRows: unknown[] }).previewRows).toHaveLength(20);
  });

  it('returns a CSV preview when the source body exceeds the network byte cap', async () => {
    const csv = [
      'id,name',
      ...Array.from({ length: 120_000 }, (_, i) => `${i},name-${i}`),
    ].join('\n');
    const fetchImpl = vi.fn(async () => textResponse(csv, {
      status: 200,
      headers: { 'content-type': 'text/csv' },
    }));

    const result = await handleFetchUrlWithOptions({
      url: 'https://public.example/large.csv',
      format: 'csv',
    }, { fetchImpl });

    expect(result).toMatchObject({ format: 'csv', truncated: true });
    expect((result as { bytesFetched: number }).bytesFetched).toBe(1_000_000);
    expect((result as { headerRow: string[] }).headerRow).toEqual(['id', 'name']);
    expect((result as { previewRows: unknown[] }).previewRows).toHaveLength(20);
  });

  it('returns a JSON text preview when truncation happens before the JSON can be parsed', async () => {
    const body = `{"rows":[${'"x",'.repeat(1_000_000)}`;
    const fetchImpl = vi.fn(async () => textResponse(body, {
      status: 200,
      headers: { 'content-type': 'application/json' },
    }));

    const result = await handleFetchUrlWithOptions({
      url: 'https://public.example/large.json',
      format: 'json',
      max_chars: 1000,
    }, { fetchImpl });

    expect(result).toMatchObject({ format: 'json', truncated: true });
    expect((result as { returnedChars: number }).returnedChars).toBeLessThanOrEqual(1000);
    expect((result as { dataPreview: string }).dataPreview).toContain('"rows"');
    expect((result as { hint: string }).hint).toMatch(/byte cap/);
  });
});

describe('fetch_url SSRF guards', () => {
  it('rejects private IP and localhost URLs before issuing a request', async () => {
    const fetchImpl = vi.fn();
    await expect(handleFetchUrlWithOptions({
      url: 'http://127.0.0.1:8080/data',
    }, { fetchImpl })).rejects.toBeInstanceOf(ToolValidationError);
    await expect(handleFetchUrlWithOptions({
      url: 'http://localhost:8080/data',
    }, { fetchImpl })).rejects.toBeInstanceOf(ToolValidationError);
    expect(fetchImpl).not.toHaveBeenCalled();
  });

  it('rejects post-redirect private URLs', async () => {
    const fetchImpl = vi.fn(async () => textResponse('nope', {
      status: 200,
      headers: { 'content-type': 'text/plain' },
      url: 'http://192.168.1.5/data',
    }));

    await expect(handleFetchUrlWithOptions({
      url: 'https://public.example/redirect',
    }, { fetchImpl })).rejects.toBeInstanceOf(ToolValidationError);
    expect(fetchImpl).toHaveBeenCalledTimes(1);
  });

  it('blocks credentials and disallowed ports', () => {
    expect(() => validatePublicHttpUrl('https://user:pass@public.example/data')).toThrow(ToolValidationError);
    expect(() => validatePublicHttpUrl('https://public.example:444/data')).toThrow(ToolValidationError);
  });
});

describe('ToolExecutor runtime none', () => {
  it('runs runtime:none tools without invoking the Excel runner', async () => {
    const registry = new WorkbookRegistry();
    const runner = vi.fn();
    const executor = new ToolExecutor(registry, runner);
    const fetchImpl = vi.fn(async () => textResponse('hello', {
      status: 200,
      headers: { 'content-type': 'text/plain' },
    }));
    executor.register(FETCH_URL, createFetchUrlHandler({ fetchImpl }));

    const result = await executor.execute(makeCall({
      url: 'https://public.example/hello.txt',
      format: 'text',
    }), SCOPE);

    expect(result.ok).toBe(true);
    expect(runner).not.toHaveBeenCalled();
    expect(fetchImpl).toHaveBeenCalledTimes(1);
  });
});
