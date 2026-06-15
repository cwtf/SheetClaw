import type { ToolSpec } from '../types';
import type { ToolHandler } from '../workbook/executor';
import { ToolNetworkError, ToolValidationError } from '../workbook/executor';
import { fetchTextWithGuards, type FetchResponse } from './net';
import { READER_PROVIDER_ENDPOINT } from './providers';

type FetchFormat = 'auto' | 'text' | 'json' | 'csv';
type FetchMode = 'preview' | 'full';

const MIN_CHARS = 1_000;
const MAX_CHARS = 20_000;
const PREVIEW_DEFAULT_CHARS = 4_000;
const FULL_DEFAULT_CHARS = 16_000;
const MAX_PREVIEW_ROWS = 20;

export const FETCH_URL: ToolSpec = {
  name: 'fetch_url',
  description: "Fetch a public HTTP(S) URL. Defaults to preview mode, which returns metadata plus a bounded sample (CSV: header + first rows; JSON: truncated structure; text/HTML: leading text). Call again with mode:'full' only after the scope is confirmed; output is still capped.",
  parameters: {
    type: 'object',
    properties: {
      url: { type: 'string', description: 'Absolute public http(s) URL to fetch' },
      mode: { type: 'string', enum: ['preview', 'full'], description: 'preview or full; default preview' },
      format: { type: 'string', enum: ['auto', 'text', 'json', 'csv'], description: 'auto, text, json, or csv; default auto' },
      max_chars: { type: 'number', description: 'Clamp between 1000 and 20000 characters' },
    },
    required: ['url'],
    additionalProperties: false,
  },
  mutating: false,
  runtime: 'none',
};

export interface FetchUrlOptions {
  fetchImpl?: typeof fetch;
  readerFallback?: boolean | (() => boolean);
  signal?: AbortSignal;
}

export function createFetchUrlHandler(options: FetchUrlOptions = {}): ToolHandler {
  return async (args) => handleFetchUrlWithOptions(args, options);
}

export const handleFetchUrl: ToolHandler = async (args) => {
  return handleFetchUrlWithOptions(args, {});
};

export async function handleFetchUrlWithOptions(
  args: Record<string, unknown>,
  options: FetchUrlOptions
): Promise<unknown> {
  const url = requiredString(args.url, 'url');
  const mode = optionalEnum<FetchMode>(args.mode, 'mode', ['preview', 'full']) ?? 'preview';
  const requestedFormat = optionalEnum<FetchFormat>(args.format, 'format', ['auto', 'text', 'json', 'csv']) ?? 'auto';
  const maxChars = clampMaxChars(args.max_chars, mode);

  try {
    const direct = await fetchTextWithGuards(url, {
      fetchImpl: options.fetchImpl,
      signal: options.signal,
      truncateAtMaxBytes: true,
    });
    return formatFetchResult(direct, requestedFormat, mode, maxChars, 'direct');
  } catch (e) {
    if (!(e instanceof ToolNetworkError) || !readerFallbackEnabled(options.readerFallback)) throw e;
    const reader = await fetchTextWithGuards(`${READER_PROVIDER_ENDPOINT}${url}`, {
      fetchImpl: options.fetchImpl,
      signal: options.signal,
      truncateAtMaxBytes: true,
    });
    return formatFetchResult(reader, requestedFormat, mode, maxChars, 'reader');
  }
}

function formatFetchResult(
  response: FetchResponse,
  requestedFormat: FetchFormat,
  mode: FetchMode,
  maxChars: number,
  source: 'direct' | 'reader'
): Record<string, unknown> {
  const format = requestedFormat === 'auto'
    ? detectFormat(response.contentType, response.finalUrl, response.text)
    : requestedFormat;
  const base = {
    url: response.url,
    finalUrl: response.finalUrl,
    status: response.status,
    contentType: response.contentType,
    source,
    format,
    bytesFetched: response.bytesFetched,
  };

  if (format === 'json') {
    return { ...base, ...formatJsonPayload(response.text, maxChars, response.truncatedByBytes) };
  }
  if (format === 'csv') {
    return { ...base, ...formatCsvPayload(response.text, mode, maxChars) };
  }
  return { ...base, ...formatTextPayload(stripHtmlIfNeeded(response.text, response.contentType), maxChars) };
}

function formatTextPayload(text: string, maxChars: number) {
  const capped = cleanTruncate(text, maxChars);
  return {
    returnedChars: capped.text.length,
    truncated: capped.truncated,
    text: capped.text,
  };
}

function formatJsonPayload(text: string, maxChars: number, truncatedByBytes: boolean) {
  let data: unknown;
  try {
    data = JSON.parse(text);
  } catch {
    if (truncatedByBytes) {
      const capped = cleanTruncate(text, maxChars);
      return {
        returnedChars: capped.text.length,
        truncated: true,
        dataPreview: capped.text,
        hint: 'JSON response exceeded the network byte cap before it could be parsed; refetch with a narrower API query or a smaller dataset.',
      };
    }
    throw new ToolValidationError('Response is not valid JSON.');
  }
  const serialized = JSON.stringify(data, null, 2);
  const capped = cleanTruncate(serialized, maxChars);
  if (!capped.truncated) {
    return {
      returnedChars: capped.text.length,
      truncated: false,
      data,
    };
  }
  return {
    returnedChars: capped.text.length,
    truncated: true,
    dataPreview: capped.text,
    hint: `${describeJsonSize(data)}; refetch with a narrower API query or ask the user to choose a scope.`,
  };
}

function formatCsvPayload(text: string, mode: FetchMode, maxChars: number) {
  const rows = parseCsvRows(text);
  const headerRow = rows[0] ?? [];
  const previewRows = rows.slice(1, MAX_PREVIEW_ROWS + 1);
  const capped = mode === 'full' ? cleanTruncate(text, maxChars) : { text: '', truncated: text.length > maxChars };
  return {
    returnedChars: mode === 'full' ? capped.text.length : JSON.stringify({ headerRow, previewRows }).length,
    truncated: mode === 'full' ? capped.truncated : text.length > maxChars,
    headerRow,
    previewRows,
    approxRowCount: rows.length,
    ...(mode === 'full' ? { text: capped.text } : {}),
  };
}

function detectFormat(contentType: string, url: string, text: string): Exclude<FetchFormat, 'auto'> {
  const ct = contentType.toLowerCase();
  if (ct.includes('json') || looksLikeJson(text)) return 'json';
  if (ct.includes('csv') || /\.csv(?:$|[?#])/i.test(url)) return 'csv';
  return 'text';
}

function stripHtmlIfNeeded(text: string, contentType: string): string {
  if (!contentType.toLowerCase().includes('html')) return text;
  return text
    .replace(/<script[\s\S]*?<\/script>/gi, ' ')
    .replace(/<style[\s\S]*?<\/style>/gi, ' ')
    .replace(/<[^>]+>/g, ' ')
    .replace(/\s+/g, ' ')
    .trim();
}

function cleanTruncate(text: string, maxChars: number): { text: string; truncated: boolean } {
  if (text.length <= maxChars) return { text, truncated: false };
  const cut = text.slice(0, maxChars);
  const boundary = Math.max(cut.lastIndexOf('\n'), cut.lastIndexOf(' '), cut.lastIndexOf(','));
  const safe = boundary >= Math.floor(maxChars * 0.8) ? cut.slice(0, boundary) : cut;
  return { text: safe, truncated: true };
}

function parseCsvRows(text: string): string[][] {
  const rows: string[][] = [];
  let row: string[] = [];
  let cell = '';
  let quoted = false;

  for (let i = 0; i < text.length; i++) {
    const ch = text[i];
    const next = text[i + 1];
    if (quoted) {
      if (ch === '"' && next === '"') {
        cell += '"';
        i++;
      } else if (ch === '"') {
        quoted = false;
      } else {
        cell += ch;
      }
    } else if (ch === '"') {
      quoted = true;
    } else if (ch === ',') {
      row.push(cell);
      cell = '';
    } else if (ch === '\n') {
      row.push(cell);
      rows.push(row);
      row = [];
      cell = '';
    } else if (ch !== '\r') {
      cell += ch;
    }
  }

  if (cell.length > 0 || row.length > 0) {
    row.push(cell);
    rows.push(row);
  }
  return rows;
}

function requiredString(value: unknown, name: string): string {
  if (typeof value !== 'string' || !value.trim()) {
    throw new ToolValidationError(`"${name}" must be a non-empty string.`);
  }
  return value.trim();
}

function optionalEnum<T extends string>(value: unknown, name: string, allowed: readonly T[]): T | undefined {
  if (value === undefined || value === null) return undefined;
  if (typeof value !== 'string' || !allowed.includes(value as T)) {
    throw new ToolValidationError(`"${name}" must be one of: ${allowed.join(', ')}.`);
  }
  return value as T;
}

function clampMaxChars(value: unknown, mode: FetchMode): number {
  if (value === undefined || value === null) return mode === 'preview' ? PREVIEW_DEFAULT_CHARS : FULL_DEFAULT_CHARS;
  if (typeof value !== 'number' || !Number.isFinite(value)) {
    throw new ToolValidationError('"max_chars" must be a number.');
  }
  return Math.min(Math.max(Math.floor(value), MIN_CHARS), MAX_CHARS);
}

function looksLikeJson(text: string): boolean {
  const trimmed = text.trim();
  return trimmed.startsWith('{') || trimmed.startsWith('[');
}

function describeJsonSize(value: unknown): string {
  if (Array.isArray(value)) return `${value.length} array items`;
  if (value && typeof value === 'object') return `${Object.keys(value).length} top-level keys`;
  return 'JSON value is too large';
}

function readerFallbackEnabled(value: FetchUrlOptions['readerFallback']): boolean {
  return typeof value === 'function' ? value() : !!value;
}
