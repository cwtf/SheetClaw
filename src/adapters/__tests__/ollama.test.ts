import { afterEach, describe, expect, it } from 'vitest';
import { OllamaAdapter, getOllamaBrowserAccessCommand, isOllamaBrowserAccessError } from '../ollama';

const originalFetch = globalThis.fetch;

afterEach(() => {
  globalThis.fetch = originalFetch;
});

describe('Ollama adapter - model listing', () => {
  it('loads models from the native tags endpoint', async () => {
    const urls: string[] = [];
    globalThis.fetch = async (url: RequestInfo | URL) => {
      urls.push(String(url));
      return new Response(JSON.stringify({
        models: [
          { name: 'llama3.2:latest', model: 'llama3.2:latest' },
          { name: 'qwen2.5:latest', model: 'qwen2.5:latest' },
        ],
      }), { status: 200, headers: { 'Content-Type': 'application/json' } });
    };

    const adapter = new OllamaAdapter({ baseUrl: 'http://localhost:11434/' });
    await expect(adapter.listModels()).resolves.toEqual([
      { id: 'llama3.2:latest', name: 'llama3.2:latest' },
      { id: 'qwen2.5:latest', name: 'qwen2.5:latest' },
    ]);
    expect(urls).toEqual(['http://localhost:11434/api/tags']);
  });

  it('reports browser access guidance when Ollama is reachable but readable fetch is blocked', async () => {
    const requests: Array<{ url: string; mode?: RequestMode }> = [];
    globalThis.fetch = async (url: RequestInfo | URL, init?: RequestInit) => {
      requests.push({ url: String(url), mode: init?.mode });
      if (init?.mode === 'no-cors') return new Response(null, { status: 200 });
      throw new TypeError('Failed to fetch');
    };

    const adapter = new OllamaAdapter({ baseUrl: 'http://localhost:11434' });
    await expect(adapter.listModels()).rejects.toThrow(/Ollama is reachable/);
    expect(requests).toEqual([
      { url: 'http://localhost:11434/api/tags', mode: undefined },
      { url: 'http://localhost:11434', mode: 'no-cors' },
    ]);
  });
});

describe('Ollama browser access helpers', () => {
  it('builds a PowerShell command for allowing the add-in origin', () => {
    expect(getOllamaBrowserAccessCommand('https://cwtf.github.io')).toBe(
      "$env:OLLAMA_ORIGINS='https://cwtf.github.io'; ollama serve"
    );
  });

  it('identifies browser access errors', () => {
    expect(isOllamaBrowserAccessError(
      'Ollama is reachable, but this add-in cannot read it. Allow the add-in origin.'
    )).toBe(true);
  });
});
