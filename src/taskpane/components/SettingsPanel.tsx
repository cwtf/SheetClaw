import { useState, useEffect } from 'react';
import {
  Button,
  Caption1,
  Checkbox,
  Combobox,
  Field,
  Input,
  Label,
  MessageBar,
  MessageBarBody,
  Option,
  Select,
  Spinner,
  Tab,
  TabList,
  tokens,
  type SelectTabData,
} from '@fluentui/react-components';
import { useStore } from '../../store/index';
import { createAdapter } from '../../adapters/index';
import { getOllamaBrowserAccessCommand, isOllamaBrowserAccessError } from '../../adapters/ollama';
import type { AuthState, ProviderConfig, ProviderKey } from '../../types';
import { getAuthCredential } from '../../auth/credentials';
import { signInWithOpenRouter } from '../../auth/oauthFlow';
import {
  getByokSectionNote,
  getProviderNativeSearchCaption,
  getSearchSettingsStatusText,
  resolveSearchTier,
} from '../../adapters/native-search';
import { getSearchProvider, SEARCH_PROVIDERS, type SearchProviderId, type WebAccessProvider } from '../../web/providers';

type ApiKeyProvider = Exclude<ProviderKey, 'ollama' | 'generic'>;
export type SettingsTabKey = 'ollama' | 'apiKeys' | 'generic' | 'search';

const SETTINGS_TABS: { key: SettingsTabKey; label: string }[] = [
  { key: 'ollama', label: 'Ollama' },
  { key: 'generic', label: 'OpenRouter' },
  { key: 'apiKeys', label: 'Other API' },
  { key: 'search', label: 'Search' },
];

const API_KEY_PROVIDERS: { key: ApiKeyProvider; label: string }[] = [
  { key: 'openai', label: 'OpenAI' },
  { key: 'anthropic', label: 'Anthropic' },
  { key: 'deepseek', label: 'DeepSeek' },
  { key: 'groq', label: 'Groq' },
  { key: 'mistral', label: 'Mistral' },
  { key: 'together', label: 'Together AI' },
  { key: 'kimi', label: 'Kimi' },
  { key: 'glm', label: 'GLM' },
  { key: 'qwen', label: 'Qwen' },
  { key: 'llama', label: 'Llama' },
];

const API_KEY_SIGNUP_LINKS: Partial<Record<ProviderKey, { label: string; url: string }>> = {
  openai: { label: 'Get an OpenAI key', url: 'https://platform.openai.com/api-keys' },
  anthropic: { label: 'Get an Anthropic key', url: 'https://console.anthropic.com/settings/keys' },
  deepseek: { label: 'Get a DeepSeek key', url: 'https://platform.deepseek.com/api_keys' },
  groq: { label: 'Get a Groq key', url: 'https://console.groq.com/keys' },
  mistral: { label: 'Get a Mistral key', url: 'https://console.mistral.ai/api-keys' },
  together: { label: 'Get a Together AI key', url: 'https://api.together.ai/settings/api-keys' },
  kimi: { label: 'Get a Kimi key', url: 'https://platform.moonshot.ai/console/api-keys' },
  glm: { label: 'Get a GLM key', url: 'https://z.ai/manage-apikey/apikey-list' },
  qwen: { label: 'Get a Qwen key', url: 'https://bailian.console.aliyun.com/' },
  llama: { label: 'Get a Llama key', url: 'https://llama.developer.meta.com/' },
};

const STATIC_MODELS: Partial<Record<ProviderKey, string[]>> = {
  openai: [
    'gpt-4o', 'gpt-4o-mini',
    'gpt-4.1', 'gpt-4.1-mini', 'gpt-4.1-nano',
    'o3', 'o3-mini', 'o4-mini',
  ],
  generic: [
    'openai/gpt-4o', 'openai/gpt-4o-mini',
    'anthropic/claude-sonnet-4-6', 'anthropic/claude-opus-4-8',
    'deepseek/deepseek-chat', 'deepseek/deepseek-r1',
    'qwen/qwen3-235b-a22b', 'qwen/qwen3.7-max',
    'meta-llama/llama-3.3-70b-instruct',
    'google/gemini-2.0-flash-001',
  ],
  deepseek: ['deepseek-v4-flash', 'deepseek-v4-pro', 'deepseek-chat', 'deepseek-reasoner'],
  groq: ['llama-3.3-70b-versatile', 'llama-3.1-8b-instant', 'openai/gpt-oss-120b', 'moonshotai/kimi-k2-instruct'],
  mistral: ['mistral-large-latest', 'mistral-small-latest', 'codestral-latest'],
  together: [
    'meta-llama/Llama-3.3-70B-Instruct-Turbo',
    'meta-llama/Llama-4-Maverick-17B-128E-Instruct-FP8',
    'deepseek-ai/DeepSeek-V3',
    'Qwen/Qwen2.5-Coder-32B-Instruct',
  ],
  kimi: ['kimi-k2.6', 'kimi-k2.5', 'kimi-k2-turbo-preview'],
  glm: ['glm-4.7', 'glm-4.6', 'glm-4.5', 'glm-4.5-flash'],
  qwen: ['qwen-plus', 'qwen-max', 'qwen-turbo', 'qwen3-coder-plus'],
  llama: ['Llama-3.3-70B-Instruct', 'Llama-4-Maverick-17B-128E-Instruct-FP8', 'Llama-4-Scout-17B-16E-Instruct-FP8'],
};

const OPENAI_CHAT_PREFIXES = ['gpt-', 'o1', 'o3', 'o4', 'chatgpt-'];
function isOpenAIChatModel(id: string): boolean {
  return OPENAI_CHAT_PREFIXES.some(p => id.startsWith(p));
}

function isOpenRouterBaseUrl(url: string): boolean {
  try {
    const parsed = new URL(url);
    return parsed.origin === 'https://openrouter.ai';
  } catch {
    return false;
  }
}

function chooseDefaultModel(providerKey: ProviderKey, baseUrl: string, ids: string[]): string {
  const preferredByProvider: Partial<Record<ProviderKey, string[]>> = {
    openai: ['gpt-4o', 'gpt-4o-mini'],
    anthropic: ['claude-sonnet-4-6', 'claude-3-5-sonnet-latest'],
    deepseek: ['deepseek-v4-flash', 'deepseek-v4-pro', 'deepseek-chat'],
    groq: ['llama-3.3-70b-versatile', 'llama-3.1-8b-instant'],
    mistral: ['mistral-large-latest', 'mistral-small-latest'],
    together: ['meta-llama/Llama-3.3-70B-Instruct-Turbo', 'meta-llama/Llama-4-Maverick-17B-128E-Instruct-FP8'],
    kimi: ['kimi-k2.6', 'kimi-k2.5'],
    glm: ['glm-4.7', 'glm-4.6'],
    qwen: ['qwen-plus', 'qwen-max'],
    llama: ['Llama-3.3-70B-Instruct', 'Llama-4-Maverick-17B-128E-Instruct-FP8'],
  };
  const preferred = preferredByProvider[providerKey];
  if (preferred) {
    const match = preferred.find(id => ids.includes(id));
    if (match) return match;
  }
  if (providerKey === 'generic' && isOpenRouterBaseUrl(baseUrl)) {
    const openRouterPreferred = [
      'openai/gpt-4o-mini',
      'openai/gpt-4o',
      'google/gemini-2.0-flash-001',
      'anthropic/claude-sonnet-4-6',
    ];
    return openRouterPreferred.find(id => ids.includes(id)) ?? ids[0] ?? '';
  }
  return ids[0] ?? '';
}

function providerToTab(provider: ProviderKey): SettingsTabKey {
  if (provider !== 'ollama' && provider !== 'generic') return 'apiKeys';
  return provider;
}

function isApiKeyProvider(provider: ProviderKey): provider is ApiKeyProvider {
  return provider !== 'ollama' && provider !== 'generic';
}

export default function SettingsPanel({ initialTab }: { initialTab?: SettingsTabKey }) {
  const providers = useStore(s => s.providers);
  const appConfig = useStore(s => s.appConfig);
  const authStates = useStore(s => s.authStates);
  const searchAuthStates = useStore(s => s.searchAuthStates);
  const setProvider = useStore(s => s.setProvider);
  const setActiveProvider = useStore(s => s.setActiveProvider);
  const saveApiKey = useStore(s => s.saveApiKey);
  const saveOAuthCredential = useStore(s => s.saveOAuthCredential);
  const setAuthState = useStore(s => s.setAuthState);
  const clearApiKey = useStore(s => s.clearApiKey);
  const saveSearchApiKey = useStore(s => s.saveSearchApiKey);
  const clearSearchApiKey = useStore(s => s.clearSearchApiKey);
  const setAppConfig = useStore(s => s.setAppConfig);
  const setWebSearchEnabled = useStore(s => s.setWebSearchEnabled);

  const [selectedTab, setSelectedTab] = useState<SettingsTabKey>(initialTab ?? providerToTab(appConfig.activeProvider));
  const [selectedApiProvider, setSelectedApiProvider] = useState<ApiKeyProvider>(
    isApiKeyProvider(appConfig.activeProvider) ? appConfig.activeProvider : 'openai'
  );

  useEffect(() => {
    if (initialTab) setSelectedTab(initialTab);
  }, [initialTab]);

  function renderProviderForm(providerKey: ProviderKey, key: string = providerKey) {
    return (
      <ProviderForm
        key={key}
        providerKey={providerKey}
        cfg={providers[providerKey]}
        auth={authStates[providerKey]}
        showActiveButton={selectedTab !== 'apiKeys'}
        isActive={appConfig.activeProvider === providerKey}
        onSetActive={() => setActiveProvider(providerKey)}
        onSave={(patch) => setProvider(providerKey, patch)}
        onSaveKey={(apiKey) => saveApiKey(providerKey, apiKey)}
        onSaveOAuthCredential={(credential) => saveOAuthCredential(providerKey, credential)}
        onSetAuthState={(patch) => setAuthState(providerKey, patch)}
        onClearKey={() => clearApiKey(providerKey)}
      />
    );
  }

  return (
    <div style={{ display: 'flex', flexDirection: 'column', height: '100%', overflow: 'hidden' }}>
      <TabList
        selectedValue={selectedTab}
        onTabSelect={(_, d: SelectTabData) => setSelectedTab(d.value as SettingsTabKey)}
        size="small"
        style={{ flexShrink: 0, paddingLeft: 4, borderBottom: `1px solid ${tokens.colorNeutralStroke1}` }}
      >
        {SETTINGS_TABS.map(p => (
          <Tab key={p.key} value={p.key}>
            {p.label}
            {providerToTab(appConfig.activeProvider) === p.key && (
              <span style={{ marginLeft: 4, color: tokens.colorBrandForeground1, fontSize: 10 }}>*</span>
            )}
          </Tab>
        ))}
      </TabList>

      <div style={{ flex: 1, overflowY: 'auto', padding: 12 }}>
        {selectedTab === 'search' ? (
          <SearchSettingsForm
            provider={appConfig.webAccess.provider}
            baseUrl={appConfig.webAccess.baseUrl ?? ''}
            engineId={appConfig.webAccess.engineId ?? ''}
            readerFallback={appConfig.webAccess.readerFallback}
            activeProvider={providers[appConfig.activeProvider]}
            searchAuthStates={searchAuthStates}
            onSaveConfig={(patch) => setAppConfig({ webAccess: { ...appConfig.webAccess, ...patch } })}
            onSaveKey={saveSearchApiKey}
            onClearKey={(provider) => {
              clearSearchApiKey(provider);
              setAppConfig({ webAccess: { ...appConfig.webAccess, provider: 'none' } });
              setWebSearchEnabled(false);
            }}
          />
        ) : selectedTab === 'apiKeys' ? (
          <div style={{ display: 'flex', flexDirection: 'column', gap: 14 }}>
            <ActiveProviderButton
              isActive={appConfig.activeProvider === selectedApiProvider}
              onSetActive={() => setActiveProvider(selectedApiProvider)}
            />
            <Field label="Provider">
              <Select
                value={selectedApiProvider}
                onChange={(_, d) => setSelectedApiProvider(d.value as ApiKeyProvider)}
                size="small"
              >
                {API_KEY_PROVIDERS.map(p => (
                  <option key={p.key} value={p.key}>{p.label}</option>
                ))}
              </Select>
            </Field>
            {renderProviderForm(selectedApiProvider, `api-${selectedApiProvider}`)}
          </div>
        ) : (
          renderProviderForm(selectedTab)
        )}
      </div>

    </div>
  );
}

interface ProviderFormProps {
  providerKey: ProviderKey;
  cfg: ProviderConfig;
  auth: AuthState;
  showActiveButton?: boolean;
  isActive: boolean;
  onSetActive: () => void;
  onSave: (patch: Partial<ProviderConfig>) => void;
  onSaveKey: (key: string) => void;
  onSaveOAuthCredential: (credential: {
    accessToken: string;
    oauthProvider?: 'openrouter';
    userId?: string;
  }) => void;
  onSetAuthState: (patch: Partial<AuthState>) => void;
  onClearKey: () => void;
}

function SearchSettingsForm({
  provider,
  baseUrl,
  engineId,
  readerFallback,
  activeProvider,
  searchAuthStates,
  onSaveConfig,
  onSaveKey,
  onClearKey,
}: {
  provider: WebAccessProvider;
  baseUrl: string;
  engineId: string;
  readerFallback: boolean;
  activeProvider: ProviderConfig;
  searchAuthStates: Record<SearchProviderId, AuthState>;
  onSaveConfig: (patch: { provider?: WebAccessProvider; baseUrl?: string; engineId?: string; readerFallback?: boolean }) => void;
  onSaveKey: (provider: SearchProviderId, key: string) => void;
  onClearKey: (provider: SearchProviderId) => void;
}) {
  const selectedProvider = provider === 'none' ? 'tavily' : provider;
  const adapter = getSearchProvider(selectedProvider);
  const auth = searchAuthStates[selectedProvider];
  const activeSearchTier = resolveSearchTier(activeProvider.provider, activeProvider.model);
  const activeProviderLabel = activeProvider.label || activeProvider.provider;
  const statusText = getSearchSettingsStatusText(activeProviderLabel, activeSearchTier);
  const byokNote = getByokSectionNote(activeProviderLabel, activeSearchTier);
  const [apiKey, setApiKey] = useState('');
  const [showKey, setShowKey] = useState(false);
  const [localBaseUrl, setLocalBaseUrl] = useState(baseUrl);
  const [localEngineId, setLocalEngineId] = useState(engineId);
  const [testStatus, setTestStatus] = useState<'idle' | 'testing' | 'ok' | 'error'>('idle');
  const [testMsg, setTestMsg] = useState('');
  const keySet = !!getAuthCredential(auth);

  useEffect(() => {
    setLocalBaseUrl(baseUrl);
  }, [baseUrl]);

  useEffect(() => {
    setLocalEngineId(engineId);
  }, [engineId]);

  function saveKey() {
    if (!apiKey.trim()) return;
    onSaveKey(selectedProvider, apiKey);
    onSaveConfig({ provider: selectedProvider });
    setApiKey('');
  }

  async function testKey() {
    if (!adapter) return;
    const key = apiKey || getAuthCredential(auth);
    setTestStatus('testing');
    setTestMsg('');
    try {
      const results = await adapter.search('spreadsheet public data', {
        maxResults: 1,
        apiKey: key,
        baseUrl: localBaseUrl || undefined,
        engineId: localEngineId || undefined,
        signal: new AbortController().signal,
      });
      setTestStatus('ok');
      setTestMsg(`Connected - ${results.length} result${results.length !== 1 ? 's' : ''} returned`);
    } catch (e) {
      setTestStatus('error');
      setTestMsg(e instanceof Error ? e.message : String(e));
    }
  }

  return (
    <div style={{ display: 'flex', flexDirection: 'column', gap: 14 }}>
      <MessageBar intent={activeSearchTier.tier === 'native' ? 'success' : 'info'}>
        <MessageBarBody>
          <Caption1>{statusText}</Caption1>
        </MessageBarBody>
      </MessageBar>

      <div style={{ display: 'flex', flexDirection: 'column', gap: 2 }}>
        <Label size="small">BYOK search provider</Label>
        <Caption1 style={{ color: tokens.colorNeutralForeground3 }}>
          Used for providers without native web search.
        </Caption1>
        {byokNote && (
          <Caption1 style={{ color: tokens.colorNeutralForeground3 }}>
            {byokNote}
          </Caption1>
        )}
      </div>

      <Field label="Provider">
        <Select
          value={provider}
          onChange={(_, d) => onSaveConfig({ provider: d.value as WebAccessProvider })}
          size="small"
        >
          <option value="none">None</option>
          {Object.values(SEARCH_PROVIDERS).map(p => (
            <option key={p.id} value={p.id}>{p.label}</option>
          ))}
        </Select>
      </Field>

      {adapter && (
        <>
          <Caption1 style={{ color: tokens.colorNeutralForeground3 }}>
            Search is optional and uses your own provider key. It is off for each new session until you enable it in Chat.
          </Caption1>
          <a href={adapter.signupUrl} target="_blank" rel="noreferrer" style={{ fontSize: 12 }}>
            {adapter.requiresKey ? `Get a ${adapter.label} key` : `${adapter.label} setup guide`}
          </a>

          {adapter.requiresKey && (
            <Field label="API Key">
              <div style={{ display: 'flex', gap: 6, alignItems: 'center' }}>
                <Input
                  type={showKey ? 'text' : 'password'}
                  placeholder={keySet ? auth.apiKeyMasked : 'Enter search API key...'}
                  value={apiKey}
                  onChange={(_, d) => setApiKey(d.value)}
                  size="small"
                  style={{ flex: 1, fontFamily: 'monospace', fontSize: 12 }}
                />
                <Button size="small" appearance="subtle" onClick={() => setShowKey(s => !s)}>
                  {showKey ? 'Hide' : 'Show'}
                </Button>
                {keySet && (
                  <Button size="small" appearance="subtle" onClick={() => onClearKey(selectedProvider)}>Clear</Button>
                )}
              </div>
            </Field>
          )}

          {adapter.requiresEngineId && (
            <Field label="Engine ID (cx)">
              <Input
                value={localEngineId}
                onChange={(_, d) => setLocalEngineId(d.value)}
                onBlur={() => onSaveConfig({ engineId: localEngineId })}
                placeholder="Programmable Search Engine id"
                size="small"
                style={{ fontFamily: 'monospace', fontSize: 12 }}
              />
            </Field>
          )}

          <Field label="Base URL (optional)">
            <Input
              value={localBaseUrl}
              onChange={(_, d) => setLocalBaseUrl(d.value)}
              onBlur={() => onSaveConfig({ baseUrl: localBaseUrl })}
              placeholder={adapter.endpoint}
              size="small"
              style={{ fontFamily: 'monospace', fontSize: 12 }}
            />
          </Field>

          <Checkbox
            label="Allow reader fallback for fetched URLs"
            checked={readerFallback}
            onChange={(_, d) => onSaveConfig({ readerFallback: !!d.checked })}
          />
          <Caption1 style={{ color: tokens.colorNeutralForeground3 }}>
            Reader fallback routes fetched URLs through the configured reader service if direct browser fetch is blocked.
          </Caption1>

          <div style={{ display: 'flex', gap: 8 }}>
            {apiKey && <Button appearance="primary" size="small" onClick={saveKey}>Save key</Button>}
            <Button
              appearance="secondary"
              size="small"
              disabled={testStatus === 'testing' || (adapter.requiresKey && !apiKey && !keySet)}
              onClick={() => void testKey()}
            >
              {testStatus === 'testing' ? 'Testing...' : adapter.requiresKey ? 'Test key' : 'Test search'}
            </Button>
          </div>

          {testStatus !== 'idle' && testStatus !== 'testing' && (
            <MessageBar intent={testStatus === 'ok' ? 'success' : 'error'}>
              <MessageBarBody>
                <Caption1>{testMsg}</Caption1>
              </MessageBarBody>
            </MessageBar>
          )}
        </>
      )}
    </div>
  );
}

type LoadState = 'idle' | 'loading' | 'loaded' | 'error';

function ProviderForm({
  providerKey, cfg, auth, isActive,
  showActiveButton = true,
  onSetActive, onSave, onSaveKey, onSaveOAuthCredential, onSetAuthState, onClearKey,
}: ProviderFormProps) {
  const [baseUrl, setBaseUrl] = useState(cfg.baseUrl);
  const [model, setModel] = useState(cfg.model);
  const [apiKey, setApiKey] = useState('');
  const [showKey, setShowKey] = useState(false);

  const initialList = cfg.knownModels?.map(m => m.id) ?? STATIC_MODELS[providerKey] ?? [];
  const [modelList, setModelList] = useState<string[]>(initialList);
  const [loadState, setLoadState] = useState<LoadState>(initialList.length > 0 ? 'loaded' : 'idle');
  const [loadError, setLoadError] = useState('');

  const [testStatus, setTestStatus] = useState<'idle' | 'testing' | 'ok' | 'error'>('idle');
  const [testMsg, setTestMsg] = useState('');
  const [oauthStatus, setOAuthStatus] = useState<'idle' | 'authenticating' | 'ok' | 'error'>('idle');
  const [copiedOllamaCmd, setCopiedOllamaCmd] = useState(false);

  const needsKey = providerKey !== 'ollama';
  const storedCredential = getAuthCredential(auth);
  const keySet = !!storedCredential;
  const supportsOpenRouterOAuth = providerKey === 'generic' && isOpenRouterBaseUrl(baseUrl);
  const signupLink = API_KEY_SIGNUP_LINKS[providerKey];

  useEffect(() => {
    const canLoad = providerKey === 'ollama'
      || providerKey === 'anthropic'
      || !!getAuthCredential(auth);
    if (!canLoad) return;
    void fetchModels(cfg.baseUrl, getAuthCredential(auth)).catch(() => undefined);
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, []);

  function commitBaseUrl(url: string) {
    onSave({ baseUrl: url, enabled: true });
  }

  function commitModel(m: string) {
    onSave({ model: m, enabled: true });
  }

  async function fetchModels(url: string, key: string): Promise<string[]> {
    setLoadState('loading');
    setLoadError('');
    try {
      const adapter = createAdapter({ ...cfg, baseUrl: url }, key);
      let found = await adapter.listModels();

      if (providerKey === 'openai') {
        const chat = found.filter(m => isOpenAIChatModel(m.id));
        if (chat.length > 0) found = chat;
      }

      const ids = found.map(m => m.id).sort();
      const fallbackModel = model.trim() ? '' : chooseDefaultModel(providerKey, url, ids);
      setModelList(ids);
      if (fallbackModel) setModel(fallbackModel);
      setLoadState('loaded');
      onSave({
        knownModels: found,
        ...(fallbackModel ? { model: fallbackModel, enabled: true } : {}),
      });
      return ids;
    } catch (e) {
      setLoadState('error');
      setLoadError(e instanceof Error ? e.message : String(e));
      throw e;
    }
  }

  async function test() {
    setTestStatus('testing');
    setTestMsg('');
    try {
      const key = apiKey || getAuthCredential(auth);
      const ids = await fetchModels(baseUrl, key);
      setTestStatus('ok');
      setTestMsg(`Connected - ${ids.length} model${ids.length !== 1 ? 's' : ''} available`);
    } catch (e) {
      setTestStatus('error');
      setTestMsg(e instanceof Error ? e.message : String(e));
    }
  }

  async function startOpenRouterOAuth() {
    setOAuthStatus('authenticating');
    setTestStatus('idle');
    setTestMsg('');
    onSetAuthState({ state: 'authenticating', error: undefined });

    try {
      const result = await signInWithOpenRouter();
      onSave({ authMode: 'oauth', enabled: true });
      onSaveOAuthCredential({
        accessToken: result.key,
        oauthProvider: 'openrouter',
        userId: result.userId,
      });
      setOAuthStatus('ok');
      const ids = await fetchModels(baseUrl, result.key);
      setTestStatus('ok');
      setTestMsg(`OpenRouter connected - ${ids.length} model${ids.length !== 1 ? 's' : ''} available`);
    } catch (e) {
      const message = e instanceof Error ? e.message : String(e);
      setOAuthStatus('error');
      onSetAuthState({ state: 'error', error: message });
      setTestStatus('error');
      setTestMsg(message);
    }
  }

  function saveKey() {
    if (apiKey) {
      onSaveKey(apiKey);
      setApiKey('');
    }
  }

  const modelLabel = loadState === 'loading'
    ? 'Model (loading...)'
    : loadState === 'loaded' && modelList.length > 0
    ? `Model (${modelList.length} available)`
    : 'Model';

  return (
    <div style={{ display: 'flex', flexDirection: 'column', gap: 14 }}>
      {showActiveButton && (
        <ActiveProviderButton isActive={isActive} onSetActive={onSetActive} />
      )}

      {signupLink && (
        <a href={signupLink.url} target="_blank" rel="noreferrer" style={{ fontSize: 12 }}>
          {signupLink.label}
        </a>
      )}

      <Caption1 style={{ color: tokens.colorNeutralForeground3 }}>
        {getProviderNativeSearchCaption(providerKey, model)}
      </Caption1>

      <Field label="Base URL">
        <Input
          value={baseUrl}
          onChange={(_, d) => setBaseUrl(d.value)}
          onBlur={() => commitBaseUrl(baseUrl)}
          size="small"
          style={{ fontFamily: 'monospace', fontSize: 12 }}
        />
      </Field>

      <Field label={modelLabel}>
        {loadState === 'loading' ? (
          <div style={{ display: 'flex', alignItems: 'center', gap: 8, height: 32 }}>
            <Spinner size="extra-small" />
            <Caption1 style={{ color: tokens.colorNeutralForeground3 }}>Fetching models...</Caption1>
          </div>
        ) : modelList.length > 0 ? (
          <div style={{ display: 'flex', gap: 6, alignItems: 'center' }}>
            <Combobox
              value={model}
              selectedOptions={model ? [model] : []}
              onOptionSelect={(_, d) => {
                const m = d.optionValue ?? '';
                setModel(m);
                commitModel(m);
              }}
              onChange={(e) => setModel(e.target.value)}
              onBlur={() => commitModel(model)}
              placeholder="Select or type a model..."
              size="small"
              style={{ flex: 1, minWidth: 0 }}
              freeform
            >
              {modelList.map(id => (
                <Option key={id} value={id}>{id}</Option>
              ))}
            </Combobox>
            <Button
              size="small"
              appearance="subtle"
              title="Refresh model list"
              onClick={() => void fetchModels(baseUrl, getAuthCredential(auth)).catch(() => undefined)}
            >Refresh</Button>
          </div>
        ) : (
          <div style={{ display: 'flex', gap: 6, alignItems: 'center' }}>
            <Input
              value={model}
              onChange={(_, d) => setModel(d.value)}
              onBlur={() => commitModel(model)}
              placeholder={providerKey === 'ollama' ? 'e.g. llama3.2' : 'e.g. gpt-4o'}
              size="small"
              style={{ flex: 1, fontFamily: 'monospace', fontSize: 12 }}
            />
            <Button
              size="small"
              appearance="subtle"
              title="Fetch available models"
              disabled={needsKey && !keySet && !apiKey}
              onClick={() => void fetchModels(baseUrl, apiKey || getAuthCredential(auth)).catch(() => undefined)}
            >Refresh</Button>
          </div>
        )}
        {loadState === 'error' && (
          <Caption1 style={{ color: tokens.colorPaletteRedForeground1, marginTop: 2 }}>
            {loadError}
          </Caption1>
        )}
      </Field>

      {supportsOpenRouterOAuth && (
        <div style={{ display: 'flex', flexDirection: 'column', gap: 6 }}>
          <Button
            appearance="primary"
            size="small"
            disabled={oauthStatus === 'authenticating'}
            onClick={() => void startOpenRouterOAuth()}
          >
            {oauthStatus === 'authenticating' ? 'Signing in...' : 'Sign in with OpenRouter'}
          </Button>
        </div>
      )}

      {needsKey && (
        <Field label={supportsOpenRouterOAuth ? 'API Key fallback' : 'API Key'}>
          <div style={{ display: 'flex', gap: 6, alignItems: 'center' }}>
            <Input
              type={showKey ? 'text' : 'password'}
              placeholder={keySet ? auth.apiKeyMasked : 'Enter API key...'}
              value={apiKey}
              onChange={(_, d) => setApiKey(d.value)}
              size="small"
              style={{ flex: 1, fontFamily: 'monospace', fontSize: 12 }}
            />
            <Button size="small" appearance="subtle" onClick={() => setShowKey(s => !s)}>
              {showKey ? 'Hide' : 'Show'}
            </Button>
            {keySet && (
              <Button size="small" appearance="subtle" onClick={onClearKey}>Clear</Button>
            )}
          </div>
        </Field>
      )}

      <div style={{ display: 'flex', alignItems: 'center', gap: 6 }}>
        <Label size="small" style={{ color: tokens.colorNeutralForeground3 }}>Status:</Label>
        <Caption1 style={{
          color: auth.state === 'authenticated'
            ? tokens.colorPaletteGreenForeground1
            : auth.state === 'unauthenticated' && !needsKey
            ? tokens.colorPaletteGreenForeground1
            : auth.state === 'error'
            ? tokens.colorPaletteRedForeground1
            : tokens.colorNeutralForeground3,
        }}>
          {auth.state === 'authenticated' ? 'authenticated'
            : auth.state === 'unauthenticated' && !needsKey ? 'no auth needed'
            : auth.error ? `${auth.state}: ${auth.error}`
            : auth.state}
        </Caption1>
      </div>

      <div style={{ display: 'flex', gap: 8 }}>
        {needsKey && apiKey && (
          <Button appearance="primary" size="small" onClick={saveKey}>Save key</Button>
        )}
        <Button
          appearance="secondary"
          size="small"
          disabled={testStatus === 'testing'}
          onClick={() => void test()}
        >
          {testStatus === 'testing' ? 'Testing...' : 'Test connection'}
        </Button>
      </div>

      {testStatus !== 'idle' && testStatus !== 'testing' && (
        <MessageBar intent={testStatus === 'ok' ? 'success' : 'error'}>
          <MessageBarBody>
            <Caption1>{testMsg}</Caption1>
          </MessageBarBody>
        </MessageBar>
      )}

      {providerKey === 'ollama' && (loadState === 'error' || testStatus === 'error') && (
        <MessageBar intent="warning">
          <MessageBarBody>
            <Caption1>
              {isOllamaBrowserAccessError(testMsg || loadError)
                ? 'Ollama is running, but browser access is blocked. Quit Ollama, then restart it with:'
                : 'Ollama may not be running. Start it with:'}
            </Caption1>
            <div style={{ display: 'flex', alignItems: 'center', gap: 6, marginTop: 4 }}>
              <span style={{
                fontFamily: 'monospace',
                fontSize: 12,
                background: tokens.colorNeutralBackground3,
                padding: '2px 6px',
                borderRadius: 3,
                flex: 1,
              }}>
                {isOllamaBrowserAccessError(testMsg || loadError) ? getOllamaBrowserAccessCommand() : 'ollama serve'}
              </span>
              <Button
                size="small"
                appearance="subtle"
                onClick={() => {
                  const command = isOllamaBrowserAccessError(testMsg || loadError)
                    ? getOllamaBrowserAccessCommand()
                    : 'ollama serve';
                  void navigator.clipboard.writeText(command).then(() => {
                    setCopiedOllamaCmd(true);
                    setTimeout(() => setCopiedOllamaCmd(false), 2000);
                  });
                }}
              >
                {copiedOllamaCmd ? 'Copied!' : 'Copy'}
              </Button>
            </div>
            {isOllamaBrowserAccessError(testMsg || loadError) && (
              <Caption1 style={{ display: 'block', marginTop: 6, color: tokens.colorNeutralForeground3 }}>
                If it still fails inside Excel, check the Office WebView loopback exemption.
              </Caption1>
            )}
          </MessageBarBody>
        </MessageBar>
      )}
    </div>
  );
}

function ActiveProviderButton({
  isActive,
  onSetActive,
}: {
  isActive: boolean;
  onSetActive: () => void;
}) {
  return (
    <Button
      size="small"
      appearance={isActive ? 'primary' : 'secondary'}
      onClick={onSetActive}
      disabled={isActive}
      style={{ width: '100%' }}
    >
      {isActive ? 'Active provider' : 'Set as active'}
    </Button>
  );
}
