import type { ProviderKey } from '../types/usage';

export type NativeSearchKind =
  | 'openrouter-server-tool'
  | 'anthropic-server-tool'
  | 'kimi-builtin-function'
  | 'qwen-enable-search'
  | 'glm-web-search-tool';

export interface NativeSearchCapability {
  provider: ProviderKey;
  kind: NativeSearchKind;
  supportsModel?: (model: string) => boolean;
  modelSupportLabel?: string;
  costNote: string;
}

export type OpenAINativeToolEntry =
  | { type: 'openrouter:web_search' }
  | { type: 'builtin_function'; function: { name: '$web_search' } }
  | {
      type: 'web_search';
      web_search: {
        enable: boolean;
        search_engine: string;
        search_result: boolean;
      };
    };

export type AnthropicNativeToolEntry = {
  type: 'web_search_20250305';
  name: 'web_search';
  max_uses: number;
};

export type OpenAINativeBodyPatch = {
  tools?: OpenAINativeToolEntry[];
  body?: Record<string, unknown>;
};

export type SearchTier = 'native' | 'byok';

export type SearchTierResolution =
  | {
      tier: 'native';
      provider: ProviderKey;
      model: string;
      capability: NativeSearchCapability;
      nativeUnavailableReason?: undefined;
    }
  | {
      tier: 'byok';
      provider: ProviderKey;
      model: string;
      capability?: NativeSearchCapability;
      nativeUnavailableReason?: string;
    };

export type SearchToggleResolution = SearchTierResolution & {
  byokReady: boolean;
  available: boolean;
};

const QWEN_NATIVE_MODELS = ['qwen3.5-plus', 'qwen3.5-flash', 'qwen3-max'];

function supportsQwenNativeSearch(model: string): boolean {
  const normalized = model.trim().toLowerCase();
  return QWEN_NATIVE_MODELS.some(prefix => normalized === prefix || normalized.startsWith(`${prefix}-`));
}

export const NATIVE_SEARCH: Partial<Record<ProviderKey, NativeSearchCapability>> = {
  generic: {
    provider: 'generic',
    kind: 'openrouter-server-tool',
    costNote: 'OpenRouter native search is billed to your OpenRouter key, currently about $0.005 per search call.',
  },
  anthropic: {
    provider: 'anthropic',
    kind: 'anthropic-server-tool',
    costNote: 'Anthropic native search is billed to your Anthropic key, currently about $10 per 1,000 searches plus result tokens.',
  },
  kimi: {
    provider: 'kimi',
    kind: 'kimi-builtin-function',
    costNote: 'Kimi native search is billed to your Moonshot key, currently about $0.005 per search call.',
  },
  qwen: {
    provider: 'qwen',
    kind: 'qwen-enable-search',
    supportsModel: supportsQwenNativeSearch,
    modelSupportLabel: 'qwen3.5-plus, qwen3.5-flash, or qwen3-max in thinking mode',
    costNote: 'Qwen native search is billed to your DashScope key under the provider search pricing for the selected model.',
  },
  glm: {
    provider: 'glm',
    kind: 'glm-web-search-tool',
    costNote: 'GLM native search is billed to your Z.AI key under the provider search pricing for the selected model.',
  },
};

export function getNativeSearchCapability(
  provider: ProviderKey,
  model: string
): NativeSearchCapability | undefined {
  const capability = NATIVE_SEARCH[provider];
  if (!capability) return undefined;
  if (capability.supportsModel && !capability.supportsModel(model)) return undefined;
  return capability;
}

export function resolveSearchTier(provider: ProviderKey, model: string): SearchTierResolution {
  const capability = NATIVE_SEARCH[provider];
  if (!capability) return { tier: 'byok', provider, model };
  if (capability.supportsModel && !capability.supportsModel(model)) {
    return {
      tier: 'byok',
      provider,
      model,
      capability,
      nativeUnavailableReason: capability.modelSupportLabel
        ? `Native search requires ${capability.modelSupportLabel}.`
        : 'Native search is not available for this model.',
    };
  }
  return { tier: 'native', provider, model, capability };
}

export function resolveSearchToggle(input: {
  provider: ProviderKey;
  model: string;
  byokReady: boolean;
}): SearchToggleResolution {
  const resolution = resolveSearchTier(input.provider, input.model);
  return {
    ...resolution,
    byokReady: input.byokReady,
    available: resolution.tier === 'native' || input.byokReady,
  };
}

export function getSearchSettingsStatusText(
  providerLabel: string,
  resolution: SearchTierResolution
): string {
  if (resolution.tier === 'native') {
    return `${providerLabel} has native web search; searches run on your provider key. ${resolution.capability.costNote}`;
  }
  if (resolution.nativeUnavailableReason) {
    return `${providerLabel} native web search is unavailable for ${resolution.model || 'this model'}: ${resolution.nativeUnavailableReason} Configure a search provider below to use the BYOK tier.`;
  }
  return `${providerLabel} has no native web search - configure a search provider below to enable search.`;
}

export function getByokSectionNote(
  providerLabel: string,
  resolution: SearchTierResolution
): string | undefined {
  if (resolution.tier !== 'native') return undefined;
  return `${providerLabel} also supports native search; selected keyless providers remain available, while keyed providers are secondary.`;
}

export function getProviderNativeSearchCaption(provider: ProviderKey, model: string): string {
  const resolution = resolveSearchTier(provider, model);
  if (resolution.tier === 'native') return 'Native web search: yes';
  if (resolution.nativeUnavailableReason) {
    return `Native web search: no for this model - ${resolution.nativeUnavailableReason}`;
  }
  return 'Native web search: no - uses the configured Search tab provider';
}

export function getUnavailableSearchToggleHint(
  providerLabel: string,
  resolution: SearchTierResolution
): string {
  if (resolution.nativeUnavailableReason) {
    return `${providerLabel} native search is not available for ${resolution.model || 'this model'}. ${resolution.nativeUnavailableReason} Select a supported model or configure a Search tab provider.`;
  }
  return `${providerLabel} has no native web search. Configure a Search tab provider to enable search.`;
}

export function getOpenAINativeSearchPatch(
  activeProvider: ProviderKey | undefined,
  capability: NativeSearchCapability | undefined
): OpenAINativeBodyPatch {
  if (!capability || capability.provider !== activeProvider) return {};

  switch (capability.kind) {
    case 'openrouter-server-tool':
      return { tools: [{ type: 'openrouter:web_search' }] };
    case 'kimi-builtin-function':
      return { tools: [{ type: 'builtin_function', function: { name: '$web_search' } }] };
    case 'qwen-enable-search':
      return {
        body: {
          enable_search: true,
          search_options: { forced_search: false, search_strategy: 'turbo' },
        },
      };
    case 'glm-web-search-tool':
      return {
        tools: [{
          type: 'web_search',
          web_search: {
            enable: true,
            search_engine: 'search-prime',
            search_result: true,
          },
        }],
      };
    case 'anthropic-server-tool':
      return {};
  }
}

export function getAnthropicNativeSearchTool(
  activeProvider: ProviderKey | undefined,
  capability: NativeSearchCapability | undefined
): AnthropicNativeToolEntry | undefined {
  if (!capability || capability.provider !== activeProvider || capability.kind !== 'anthropic-server-tool') {
    return undefined;
  }
  return { type: 'web_search_20250305', name: 'web_search', max_uses: 5 };
}
