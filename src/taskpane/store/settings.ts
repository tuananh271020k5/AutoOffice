export const LOCAL_PROVIDER_IDS = new Set(['ollama', 'lmstudio']);

export interface McpServerConfig {
  name: string;
  url: string;
  transport: 'streamable-http' | 'sse';
  enabled: boolean;
}

export interface ProviderConfig {
  id: string;
  name: string;
  apiKey: string;
  baseUrl?: string; // For OpenAI-compatible local models
}

export interface AppSettings {
  selectedProviderId: string;
  selectedModel: string;
  providers: ProviderConfig[];
  autoApprove: boolean;
  /** Word writes always show a preview and keep a recovery copy when enabled. */
  previewWordChanges?: boolean;
  mcpServers: McpServerConfig[];
  maxRetries: number;
  executionTimeout: number;
}

function getEnv(key: string): string {
  try {
    const metaEnv = (import.meta as unknown as { env?: Record<string, string | undefined> }).env;
    return (metaEnv?.[key] ?? '').trim();
  } catch {
    return '';
  }
}

export function buildDefaultSettings(): AppSettings {
  const defaultProviderId = getEnv('VITE_DEFAULT_PROVIDER') || 'openai';
  const defaultModel = getEnv('VITE_DEFAULT_MODEL') || 'gpt-4o';
  const defaultApiKey = getEnv('VITE_DEFAULT_API_KEY');
  const defaultBaseUrl = getEnv('VITE_DEFAULT_BASE_URL');
  const autoApproveEnv = getEnv('VITE_AUTO_APPROVE').toLowerCase();
  const autoApprove = autoApproveEnv ? autoApproveEnv !== 'false' && autoApproveEnv !== '0' : true;
  const maxRetries = parseInt(getEnv('VITE_MAX_RETRIES') || '3', 10) || 3;
  const timeoutSec = parseInt(getEnv('VITE_EXECUTION_TIMEOUT_SECONDS') || '30', 10) || 30;

  const baseProviders: ProviderConfig[] = [
    { id: 'anthropic', name: 'Anthropic', apiKey: getEnv('VITE_ANTHROPIC_API_KEY') },
    { id: 'openai', name: 'OpenAI', apiKey: getEnv('VITE_OPENAI_API_KEY') },
    { id: 'google', name: 'Google', apiKey: getEnv('VITE_GOOGLE_API_KEY') },
    { id: 'groq', name: 'Groq', apiKey: getEnv('VITE_GROQ_API_KEY') },
    { id: 'xai', name: 'xAI', apiKey: getEnv('VITE_XAI_API_KEY') },
    { id: 'deepseek', name: 'DeepSeek', apiKey: getEnv('VITE_DEEPSEEK_API_KEY') },
    { id: 'gateway', name: 'Vercel AI Gateway', apiKey: getEnv('VITE_GATEWAY_API_KEY') },
    { id: 'openrouter', name: 'OpenRouter', apiKey: getEnv('VITE_OPENROUTER_API_KEY'), baseUrl: getEnv('VITE_OPENROUTER_BASE_URL') },
    { id: 'ollama', name: 'Ollama', apiKey: '', baseUrl: getEnv('VITE_OLLAMA_BASE_URL') },
    { id: 'lmstudio', name: 'LM Studio', apiKey: '', baseUrl: getEnv('VITE_LMSTUDIO_BASE_URL') || 'http://localhost:1234/v1' },
    { id: 'openai-compatible', name: 'OpenAI-Compatible', apiKey: getEnv('VITE_OPENAI_COMPATIBLE_API_KEY'), baseUrl: getEnv('VITE_OPENAI_COMPATIBLE_BASE_URL') },
  ];

  const providers = baseProviders.map(p => {
    if (p.id === defaultProviderId) {
      return {
        ...p,
        apiKey: p.apiKey || defaultApiKey,
        ...(defaultBaseUrl ? { baseUrl: defaultBaseUrl } : {}),
      };
    }
    return p;
  });

  return {
    selectedProviderId: defaultProviderId,
    selectedModel: defaultModel,
    providers,
    autoApprove,
    previewWordChanges: true,
    mcpServers: [],
    maxRetries,
    executionTimeout: timeoutSec * 1000,
  };
}

const STORAGE_KEY = 'autooffice_settings';

function mergeSettings(saved: Partial<AppSettings>): AppSettings {
  const defaults = buildDefaultSettings();
  const savedProviders = saved.providers ?? [];
  const providers = defaults.providers.map(def => {
    const existing = savedProviders.find(p => p.id === def.id);
    if (!existing) return def;
    return {
      ...def,
      ...existing,
      apiKey: existing.apiKey ?? def.apiKey,
      baseUrl: existing.baseUrl || def.baseUrl,
    };
  });

  const selectedProviderId = saved.selectedProviderId || defaults.selectedProviderId || 'openai';
  const selectedModel = saved.selectedModel || defaults.selectedModel || 'gpt-4o';

  return {
    ...defaults,
    ...saved,
    selectedProviderId,
    selectedModel,
    providers,
  };
}

/**
 * Loads settings strictly from the current user's local machine (`localStorage`).
 * Migrates any legacy `Office.context.roamingSettings` entry into `localStorage` once
 * so API keys are never synced across shared Office accounts or baked into shared builds.
 */
export function loadSettings(): AppSettings {
  const defaults = buildDefaultSettings();
  try {
    const localRaw = localStorage.getItem(STORAGE_KEY);
    if (localRaw) {
      return mergeSettings(JSON.parse(localRaw));
    }
    if (typeof Office !== 'undefined' && Office.context?.roamingSettings) {
      const roamingRaw = Office.context.roamingSettings.get(STORAGE_KEY);
      if (roamingRaw) {
        const merged = mergeSettings(JSON.parse(roamingRaw));
        localStorage.setItem(STORAGE_KEY, JSON.stringify(merged));
        return merged;
      }
    }
  } catch {
    // Fall through to defaults
  }
  return defaults;
}

export function saveSettings(settings: AppSettings): void {
  const json = JSON.stringify(settings);
  try {
    localStorage.setItem(STORAGE_KEY, json);
  } catch {
    // Silent failure — settings will be lost on next load
  }
}

export function hasActiveProviderApiKey(settings: AppSettings): boolean {
  const prov = settings.providers.find(p => p.id === settings.selectedProviderId);
  if (!prov) return false;
  if (LOCAL_PROVIDER_IDS.has(prov.id)) return true;
  return Boolean(prov.apiKey && prov.apiKey.trim().length > 0);
}
