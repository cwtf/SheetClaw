import { describe, expect, it } from 'vitest';
import {
  NATIVE_SEARCH,
  getByokSectionNote,
  getNativeSearchCapability,
  getProviderNativeSearchCaption,
  getSearchSettingsStatusText,
  getUnavailableSearchToggleHint,
  resolveSearchTier,
  resolveSearchToggle,
} from '../native-search';
import type { ProviderKey } from '../../types';

describe('native search registry', () => {
  it('declares the v1 native providers and mutation kinds in one place', () => {
    expect(NATIVE_SEARCH.generic?.kind).toBe('openrouter-server-tool');
    expect(NATIVE_SEARCH.anthropic?.kind).toBe('anthropic-server-tool');
    expect(NATIVE_SEARCH.kimi?.kind).toBe('kimi-builtin-function');
    expect(NATIVE_SEARCH.qwen?.kind).toBe('qwen-enable-search');
    expect(NATIVE_SEARCH.glm?.kind).toBe('glm-web-search-tool');
  });

  it('resolves native tier for native providers before considering BYOK readiness', () => {
    expect(resolveSearchToggle({ provider: 'anthropic', model: 'claude-sonnet-4-6', byokReady: false })).toMatchObject({
      tier: 'native',
      available: true,
    });
    expect(resolveSearchToggle({ provider: 'generic', model: 'openai/gpt-4o-mini', byokReady: true })).toMatchObject({
      tier: 'native',
      available: true,
    });
  });

  it('resolves non-native providers to the BYOK tier', () => {
    const nonNative: ProviderKey[] = ['ollama', 'openai', 'deepseek', 'groq', 'mistral', 'together', 'llama'];
    for (const provider of nonNative) {
      expect(resolveSearchTier(provider, 'model')).toEqual({ tier: 'byok', provider, model: 'model' });
    }
  });

  it('keeps BYOK toggle availability dependent on a ready BYOK provider', () => {
    expect(resolveSearchToggle({ provider: 'ollama', model: 'llama3.2', byokReady: false })).toMatchObject({
      tier: 'byok',
      available: false,
    });
    expect(resolveSearchToggle({ provider: 'ollama', model: 'llama3.2', byokReady: true })).toMatchObject({
      tier: 'byok',
      available: true,
    });
  });

  it('applies Qwen native search only to the supported model family', () => {
    expect(getNativeSearchCapability('qwen', 'qwen3.5-plus')?.kind).toBe('qwen-enable-search');
    expect(getNativeSearchCapability('qwen', 'qwen3-max-2026')?.kind).toBe('qwen-enable-search');

    const resolution = resolveSearchTier('qwen', 'qwen-plus');
    expect(resolution.tier).toBe('byok');
    expect(resolution.nativeUnavailableReason).toContain('qwen3.5-plus');
  });
});

describe('native search copy helpers', () => {
  it('returns Settings status text and BYOK annotation for native providers', () => {
    const resolution = resolveSearchTier('generic', 'openai/gpt-4o-mini');

    expect(getSearchSettingsStatusText('OpenRouter', resolution)).toContain('OpenRouter has native web search');
    expect(getSearchSettingsStatusText('OpenRouter', resolution)).toContain('provider key');
    expect(getByokSectionNote('OpenRouter', resolution)).toBe(
      'OpenRouter also supports native search; selected keyless providers remain available, while keyed providers are secondary.'
    );
    expect(getProviderNativeSearchCaption('generic', 'openai/gpt-4o-mini')).toBe('Native web search: yes');
  });

  it('returns keyless BYOK-tier hint text for non-native providers', () => {
    const resolution = resolveSearchTier('ollama', 'llama3.2');

    expect(getSearchSettingsStatusText('Ollama', resolution)).toContain('Ollama has no native web search');
    expect(getByokSectionNote('Ollama', resolution)).toBeUndefined();
    expect(getProviderNativeSearchCaption('ollama', 'llama3.2')).toBe(
      'Native web search: no - uses the configured Search tab provider'
    );
    expect(getUnavailableSearchToggleHint('Ollama', resolution)).toBe(
      'Ollama has no native web search. Configure a Search tab provider to enable search.'
    );
  });

  it('names the Qwen model constraint when native search is unavailable for the model', () => {
    const resolution = resolveSearchTier('qwen', 'qwen-plus');

    expect(getSearchSettingsStatusText('Qwen', resolution)).toContain('qwen3.5-plus');
    expect(getProviderNativeSearchCaption('qwen', 'qwen-plus')).toContain('qwen3.5-plus');
    expect(getUnavailableSearchToggleHint('Qwen', resolution)).toContain('Select a supported model');
  });
});
