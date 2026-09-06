// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

/**
 * Model Settings Store
 * Manages AI model selection and reasoning effort preferences
 * Supports per-feature model overrides for optimal results
 * Persists to localStorage for cross-session consistency
 */

import { useState, useEffect, useCallback } from 'react';

export type ModelType = 'gpt-6-astra' | 'gpt-5.1' | 'gpt-5.2' | 'gpt-5.4' | 'gpt-5.4-mini' | 'gpt-5.6-sol' | 'gpt-5.6-terra' | 'gpt-5.6-luna' | 'claude-opus-5' | 'deepseek-v3.2-speciale' | 'deepseek-v4-pro' | 'grok-4.1-fast' | 'grok-4.3' | 'mistral-large-3' | 'kimi-k2-5' | 'kimi-k2-7-code';
export const REASONING_EFFORT_OPTIONS = [
  { value: 'none', label: 'None' },
  { value: 'minimal', label: 'Minimal' },
  { value: 'low', label: 'Low' },
  { value: 'medium', label: 'Medium' },
  { value: 'high', label: 'High' },
  { value: 'xhigh', label: 'Extra High' },
  { value: 'max', label: 'Max' },
] as const;
export type ReasoningEffort = (typeof REASONING_EFFORT_OPTIONS)[number]['value'];

const STANDARD_REASONING_EFFORTS = ['none', 'low', 'medium', 'high'] as const;
const EXTENDED_REASONING_EFFORTS = [...STANDARD_REASONING_EFFORTS, 'xhigh'] as const;
const FRONTIER_REASONING_EFFORTS = [...EXTENDED_REASONING_EFFORTS, 'max'] as const;
const CLAUDE_REASONING_EFFORTS = ['low', 'medium', 'high', 'max'] as const;

/**
 * Feature types that can have independent model settings
 */
export type FeatureType = 'architectureGeneration' | 'validation' | 'deploymentGuide' | 'blueprint';

/**
 * Per-feature model override settings
 * When undefined, the feature uses the default model settings
 */
export interface FeatureModelOverride {
  model: ModelType;
  reasoningEffort?: ReasoningEffort; // Only used for reasoning models
}

export interface ModelSettings {
  model: ModelType;
  reasoningEffort: ReasoningEffort;
  // Per-feature overrides (optional)
  featureOverrides?: Partial<Record<FeatureType, FeatureModelOverride>>;
}

const STORAGE_KEY = 'azure-diagrams-model-settings';
const STORAGE_VERSION = 3;
const ASTRA_MIGRATION_VERSION = 1;

const DEFAULT_SETTINGS: ModelSettings = {
  model: 'gpt-6-astra',
  reasoningEffort: 'low',
  featureOverrides: {}
};

/**
 * Feature display configuration
 */
export const FEATURE_CONFIG: Record<FeatureType, {
  displayName: string;
  description: string;
  recommendedModel: ModelType;
  recommendedReasoning?: ReasoningEffort;
}> = {
  architectureGeneration: {
    displayName: 'Architecture Generation',
    description: 'Creating Azure architecture diagrams',
    recommendedModel: 'gpt-6-astra',
    recommendedReasoning: 'low'
  },
  validation: {
    displayName: 'Architecture Validation',
    description: 'WAF validation and security analysis',
    recommendedModel: 'gpt-6-astra',
    recommendedReasoning: 'low'
  },
  deploymentGuide: {
    displayName: 'Deployment Guide & Bicep',
    description: 'Generating deployment guides and IaC templates',
    recommendedModel: 'gpt-6-astra',
    recommendedReasoning: 'low'
  },
  blueprint: {
    displayName: 'Blueprint Diagrams',
    description: 'Whiteboard-style blueprint sketches (fast, cost-efficient)',
    recommendedModel: 'gpt-6-astra',
    recommendedReasoning: 'low'
  }
};

const LEGACY_FEATURE_MODELS: Record<FeatureType, ModelType> = {
  architectureGeneration: 'gpt-5.6-sol',
  validation: 'gpt-5.6-terra',
  deploymentGuide: 'gpt-5.6-terra',
  blueprint: 'gpt-5.6-luna',
};

/**
 * Model configuration including deployment names and parameters
 */
export const MODEL_CONFIG: Record<ModelType, {
  displayName: string;
  deploymentEnvVar: string;
  isReasoning: boolean;
  maxCompletionTokens: number;
  description: string;
  recommendedUse?: string;
  defaultReasoningEffort?: ReasoningEffort;
  supportedReasoningEfforts?: readonly ReasoningEffort[];
  apiFormat?: 'responses' | 'chat-completions' | 'anthropic-messages'; // defaults to 'responses'
  supportsVision?: boolean; // defaults to true
}> = {
  'gpt-6-astra': {
    displayName: 'GPT-6 Astra',
    deploymentEnvVar: 'VITE_AZURE_OPENAI_DEPLOYMENT_GPT6ASTRA',
    isReasoning: true,
    maxCompletionTokens: 32000,
    description: 'Azure OpenAI model for architecture design, validation, and deployment guidance',
    recommendedUse: 'Recommended for all features',
    defaultReasoningEffort: 'low',
    supportedReasoningEfforts: FRONTIER_REASONING_EFFORTS,
    apiFormat: 'responses',
    supportsVision: true,
  },
  'gpt-5.1': {
    displayName: 'GPT-5.1',
    deploymentEnvVar: 'VITE_AZURE_OPENAI_DEPLOYMENT_GPT51',
    isReasoning: true,
    maxCompletionTokens: 32000,
    description: 'Versatile model - fast by default, optional reasoning when needed',
    defaultReasoningEffort: 'none',
    supportedReasoningEfforts: STANDARD_REASONING_EFFORTS,
  },
  'gpt-5.2': {
    displayName: 'GPT-5.2',
    deploymentEnvVar: 'VITE_AZURE_OPENAI_DEPLOYMENT_GPT52',
    isReasoning: true,
    maxCompletionTokens: 32000,
    description: 'Most capable reasoning model - best for complex architectures',
    supportedReasoningEfforts: EXTENDED_REASONING_EFFORTS,
  },
  'gpt-5.4': {
    displayName: 'GPT-5.4',
    deploymentEnvVar: 'VITE_AZURE_OPENAI_DEPLOYMENT_GPT54',
    isReasoning: true,
    maxCompletionTokens: 32000,
    description: 'Most capable frontier model - best knowledge work, coding, and tool use',
    supportedReasoningEfforts: EXTENDED_REASONING_EFFORTS,
  },
  'gpt-5.4-mini': {
    displayName: 'GPT-5.4 Mini',
    deploymentEnvVar: 'VITE_AZURE_OPENAI_DEPLOYMENT_GPT54MINI',
    isReasoning: true,
    maxCompletionTokens: 32000,
    description: 'Compact frontier model - fast and cost-efficient with strong reasoning',
    defaultReasoningEffort: 'low',
    supportedReasoningEfforts: EXTENDED_REASONING_EFFORTS,
  },
  'gpt-5.6-sol': {
    displayName: 'GPT-5.6 Sol',
    deploymentEnvVar: 'VITE_AZURE_OPENAI_DEPLOYMENT_GPT56SOL',
    isReasoning: true,
    maxCompletionTokens: 32000,
    description: 'Previous-generation reasoning model for complex architectures',
    recommendedUse: 'Alternative model',
    defaultReasoningEffort: 'low',
    supportedReasoningEfforts: FRONTIER_REASONING_EFFORTS,
  },
  'gpt-5.6-terra': {
    displayName: 'GPT-5.6 Terra',
    deploymentEnvVar: 'VITE_AZURE_OPENAI_DEPLOYMENT_GPT56TERRA',
    isReasoning: true,
    maxCompletionTokens: 32000,
    description: 'Frontier reasoning model - grounded, thorough analysis for complex architectures',
    recommendedUse: 'Validation + deployment',
    defaultReasoningEffort: 'low',
    supportedReasoningEfforts: FRONTIER_REASONING_EFFORTS,
  },
  'gpt-5.6-luna': {
    displayName: 'GPT-5.6 Luna',
    deploymentEnvVar: 'VITE_AZURE_OPENAI_DEPLOYMENT_GPT56LUNA',
    isReasoning: true,
    maxCompletionTokens: 32000,
    description: 'Frontier reasoning model - fast, creative reasoning for architecture design',
    recommendedUse: 'Fast blueprints',
    defaultReasoningEffort: 'low',
    supportedReasoningEfforts: FRONTIER_REASONING_EFFORTS,
  },
  'claude-opus-5': {
    displayName: 'Claude Opus 5',
    deploymentEnvVar: 'VITE_AZURE_FOUNDRY_DEPLOYMENT_CLAUDE_OPUS5',
    isReasoning: true,
    maxCompletionTokens: 32000,
    description: 'Anthropic frontier model hosted in Microsoft Foundry - deep analysis and strong structured output',
    recommendedUse: 'Alternative frontier',
    defaultReasoningEffort: 'low',
    supportedReasoningEfforts: CLAUDE_REASONING_EFFORTS,
    apiFormat: 'anthropic-messages',
    supportsVision: true,
  },
  'deepseek-v3.2-speciale': {
    displayName: 'DeepSeek V3.2 Speciale',
    deploymentEnvVar: 'VITE_AZURE_OPENAI_DEPLOYMENT_DEEPSEEK',
    isReasoning: false,
    maxCompletionTokens: 16000,
    description: 'Strong structured JSON output at lower cost - third-party model',
    apiFormat: 'chat-completions',
    supportsVision: false,
  },
  'deepseek-v4-pro': {
    displayName: 'DeepSeek V4 Pro',
    deploymentEnvVar: 'VITE_AZURE_OPENAI_DEPLOYMENT_DEEPSEEK_V4_PRO',
    isReasoning: false,
    maxCompletionTokens: 16000,
    description: 'Flagship DeepSeek V4 - top-tier quality at third-party pricing',
    apiFormat: 'chat-completions',
    supportsVision: false,
  },
  'grok-4.1-fast': {
    displayName: 'Grok 4.1 Fast',
    deploymentEnvVar: 'VITE_AZURE_OPENAI_DEPLOYMENT_GROK4FAST',
    isReasoning: false,
    maxCompletionTokens: 16000,
    description: 'Fast non-reasoning model from xAI - diversified provider',
    apiFormat: 'chat-completions',
    supportsVision: false,
  },
  'grok-4.3': {
    displayName: 'Grok 4.3',
    deploymentEnvVar: 'VITE_AZURE_OPENAI_DEPLOYMENT_GROK43',
    isReasoning: false,
    maxCompletionTokens: 16000,
    description: 'Frontier xAI model - top-tier quality, broad knowledge',
    apiFormat: 'chat-completions',
    supportsVision: false,
  },
  'mistral-large-3': {
    displayName: 'Mistral Large 3',
    deploymentEnvVar: 'VITE_AZURE_OPENAI_DEPLOYMENT_MISTRALLARGE3',
    isReasoning: false,
    maxCompletionTokens: 16000,
    description: 'Mistral flagship - strong reasoning and multilingual',
    apiFormat: 'chat-completions',
    supportsVision: false,
  },
  'kimi-k2-5': {
    displayName: 'Kimi K2.5',
    deploymentEnvVar: 'VITE_AZURE_OPENAI_DEPLOYMENT_KIMIK25',
    isReasoning: false,
    maxCompletionTokens: 16000,
    description: 'MoonshotAI trillion-param MoE - strong JSON / long context',
    apiFormat: 'chat-completions',
    supportsVision: false,
  },
  'kimi-k2-7-code': {
    displayName: 'Kimi K2.7 Code',
    deploymentEnvVar: 'VITE_AZURE_OPENAI_DEPLOYMENT_KIMIK27CODE',
    isReasoning: false,
    // Kimi K2.7 Code emits an internal reasoning trace (reasoning_content) that
    // consumes the completion budget before any answer content is produced. A
    // 16k budget is frequently exhausted by reasoning + large JSON on complex
    // architectures, truncating (finish_reason=length) with empty content. Give
    // it a larger budget so reasoning and the JSON answer both fit.
    maxCompletionTokens: 32000,
    description: 'MoonshotAI Kimi K2.7 - optimized for code and structured output',
    apiFormat: 'chat-completions',
    supportsVision: false,
  },
};

export function isReasoningEffort(value: unknown): value is ReasoningEffort {
  return REASONING_EFFORT_OPTIONS.some(option => option.value === value);
}

function isModelType(value: unknown): value is ModelType {
  return typeof value === 'string' && Object.prototype.hasOwnProperty.call(MODEL_CONFIG, value);
}

export function getSupportedReasoningEfforts(model: ModelType): readonly ReasoningEffort[] {
  const config = MODEL_CONFIG[model];
  return config.isReasoning
    ? (config.supportedReasoningEfforts ?? STANDARD_REASONING_EFFORTS)
    : [];
}

export function getCommonSupportedReasoningEfforts(models: Iterable<ModelType>): ReasoningEffort[] {
  const reasoningModels = [...models].filter(model => MODEL_CONFIG[model].isReasoning);
  if (reasoningModels.length === 0) return [];
  return REASONING_EFFORT_OPTIONS
    .map(option => option.value)
    .filter(effort => reasoningModels.every(model => getSupportedReasoningEfforts(model).includes(effort)));
}

export function getReasoningEffortLabel(effort: ReasoningEffort): (typeof REASONING_EFFORT_OPTIONS)[number]['label'] {
  return REASONING_EFFORT_OPTIONS.find(option => option.value === effort)?.label ?? 'Medium';
}

export function normalizeReasoningEffort(model: ModelType, effort: unknown): ReasoningEffort {
  const config = MODEL_CONFIG[model];
  if (!config.isReasoning) {
    return isReasoningEffort(effort) ? effort : DEFAULT_SETTINGS.reasoningEffort;
  }

  const supported = getSupportedReasoningEfforts(model);
  if (isReasoningEffort(effort) && supported.includes(effort)) return effort;
  if (config.defaultReasoningEffort && supported.includes(config.defaultReasoningEffort)) {
    return config.defaultReasoningEffort;
  }
  return supported.includes('medium') ? 'medium' : supported[0];
}

function normalizeFeatureOverrides(
  value: unknown,
  migrateToAstra = false,
): Partial<Record<FeatureType, FeatureModelOverride>> {
  if (!value || typeof value !== 'object') return {};

  const normalized: Partial<Record<FeatureType, FeatureModelOverride>> = {};
  for (const [feature, rawOverride] of Object.entries(value)) {
    if (!Object.prototype.hasOwnProperty.call(FEATURE_CONFIG, feature)) continue;
    if (!rawOverride || typeof rawOverride !== 'object' || !('model' in rawOverride)) continue;
    if (!isModelType(rawOverride.model)) continue;
    const model = migrateToAstra && rawOverride.model.startsWith('gpt-5.6-')
      ? 'gpt-6-astra'
      : rawOverride.model;
    if (!isModelAvailable(model)) continue;

    const rawEffort = 'reasoningEffort' in rawOverride ? rawOverride.reasoningEffort : undefined;
    normalized[feature as FeatureType] = {
      model,
      reasoningEffort: rawEffort === undefined
        ? undefined
        : normalizeReasoningEffort(model, rawEffort),
    };
  }
  return normalized;
}

/**
 * Static map of deployment names per model.
 *
 * SECURITY: These MUST be accessed with literal `import.meta.env.VITE_...` keys.
 * Using a dynamic/computed key (e.g. `import.meta.env[someVar]`) forces Vite to
 * inline the ENTIRE env object into the client bundle — which leaks every VITE_
 * variable, including the Azure OpenAI API key. Deployment names themselves are
 * not secrets, so embedding them is fine.
 *
 * The literals are read lazily inside a function so this module can also be
 * imported by non-Vite runtimes (unit tests, scripts) where `import.meta.env`
 * does not exist.
 */
let deploymentNamesCache: Record<ModelType, string | undefined> | null = null;

export function getDeploymentNames(): Record<ModelType, string | undefined> {
  if (deploymentNamesCache) {
    return deploymentNamesCache;
  }
  try {
    deploymentNamesCache = {
      'gpt-6-astra': import.meta.env.VITE_AZURE_OPENAI_DEPLOYMENT_GPT6ASTRA,
      'gpt-5.1': import.meta.env.VITE_AZURE_OPENAI_DEPLOYMENT_GPT51,
      'gpt-5.2': import.meta.env.VITE_AZURE_OPENAI_DEPLOYMENT_GPT52,
      'gpt-5.4': import.meta.env.VITE_AZURE_OPENAI_DEPLOYMENT_GPT54,
      'gpt-5.4-mini': import.meta.env.VITE_AZURE_OPENAI_DEPLOYMENT_GPT54MINI,
      'gpt-5.6-sol': import.meta.env.VITE_AZURE_OPENAI_DEPLOYMENT_GPT56SOL,
      'gpt-5.6-terra': import.meta.env.VITE_AZURE_OPENAI_DEPLOYMENT_GPT56TERRA,
      'gpt-5.6-luna': import.meta.env.VITE_AZURE_OPENAI_DEPLOYMENT_GPT56LUNA,
      'claude-opus-5': import.meta.env.VITE_AZURE_FOUNDRY_DEPLOYMENT_CLAUDE_OPUS5,
      'deepseek-v3.2-speciale': import.meta.env.VITE_AZURE_OPENAI_DEPLOYMENT_DEEPSEEK,
      'deepseek-v4-pro': import.meta.env.VITE_AZURE_OPENAI_DEPLOYMENT_DEEPSEEK_V4_PRO,
      'grok-4.1-fast': import.meta.env.VITE_AZURE_OPENAI_DEPLOYMENT_GROK4FAST,
      'grok-4.3': import.meta.env.VITE_AZURE_OPENAI_DEPLOYMENT_GROK43,
      'mistral-large-3': import.meta.env.VITE_AZURE_OPENAI_DEPLOYMENT_MISTRALLARGE3,
      'kimi-k2-5': import.meta.env.VITE_AZURE_OPENAI_DEPLOYMENT_KIMIK25,
      'kimi-k2-7-code': import.meta.env.VITE_AZURE_OPENAI_DEPLOYMENT_KIMIK27CODE,
    };
  } catch {
    deploymentNamesCache = Object.fromEntries(
      (Object.keys(MODEL_CONFIG) as ModelType[]).map(model => [model, undefined]),
    ) as Record<ModelType, string | undefined>;
  }
  return deploymentNamesCache;
}

/**
 * Get deployment name for a specific model
 * Each model requires its own deployment env var to be set
 */
export function getDeploymentName(model: ModelType): string {
  const config = MODEL_CONFIG[model];

  // Static lookup (see the deployment-name note above — do not use a dynamic key).
  const specificDeployment = getDeploymentNames()[model];
  if (specificDeployment) {
    return specificDeployment;
  }
  
  // No fallback - each model needs its own deployment configured
  throw new Error(`No deployment configured for ${config.displayName}. Set ${config.deploymentEnvVar} in your .env file.`);
}

/**
 * Build the recommended application portfolio from models that are actually deployed.
 * Prefer Astra throughout; retain the existing portfolio on installations without it.
 */
export function getRecommendedModelSettings(): ModelSettings {
  const availableModels = getAvailableModels();
  const recommendationFor = (feature: FeatureType) => {
    const recommendation = FEATURE_CONFIG[feature];
    return {
      ...recommendation,
      recommendedModel: availableModels.includes(recommendation.recommendedModel)
        ? recommendation.recommendedModel
        : LEGACY_FEATURE_MODELS[feature],
    };
  };
  const architectureRecommendation = recommendationFor('architectureGeneration');
  const defaultModel = availableModels.includes(architectureRecommendation.recommendedModel)
    ? architectureRecommendation.recommendedModel
    : (availableModels[0] || DEFAULT_SETTINGS.model);
  const defaultReasoning = normalizeReasoningEffort(
    defaultModel,
    architectureRecommendation.recommendedReasoning
      || MODEL_CONFIG[defaultModel].defaultReasoningEffort
      || DEFAULT_SETTINGS.reasoningEffort,
  );
  const featureOverrides: Partial<Record<FeatureType, FeatureModelOverride>> = {};

  (Object.keys(FEATURE_CONFIG) as FeatureType[]).forEach((feature) => {
    const recommendation = recommendationFor(feature);
    if (!availableModels.includes(recommendation.recommendedModel)) return;

    const reasoningEffort = normalizeReasoningEffort(
      recommendation.recommendedModel,
      recommendation.recommendedReasoning
        || MODEL_CONFIG[recommendation.recommendedModel].defaultReasoningEffort
        || defaultReasoning,
    );

    if (
      recommendation.recommendedModel !== defaultModel
      || reasoningEffort !== defaultReasoning
    ) {
      featureOverrides[feature] = {
        model: recommendation.recommendedModel,
        reasoningEffort,
      };
    }
  });

  return {
    model: defaultModel,
    reasoningEffort: defaultReasoning,
    featureOverrides,
  };
}

/**
 * Load settings from localStorage
 */
function loadSettings(): ModelSettings {
  const availableModels = getAvailableModels();
  const fallbackModel = availableModels.includes(DEFAULT_SETTINGS.model)
    ? DEFAULT_SETTINGS.model
    : (availableModels[0] || DEFAULT_SETTINGS.model);

  try {
    const stored = typeof localStorage === 'undefined' ? null : localStorage.getItem(STORAGE_KEY);
    if (stored) {
      const parsed: unknown = JSON.parse(stored);
      if (parsed && typeof parsed === 'object' && 'model' in parsed && isModelType(parsed.model)) {
        const storedVersion = 'version' in parsed && Number.isInteger(parsed.version)
          ? Number(parsed.version)
          : 1;
        // Record adoption only when Astra is deployed, so later configuration changes
        // still migrate old selections without overriding subsequent explicit choices.
        const astraMigrationApplied = 'astraMigrationVersion' in parsed
          && Number.isInteger(parsed.astraMigrationVersion)
          && Number(parsed.astraMigrationVersion) >= ASTRA_MIGRATION_VERSION;
        const migrateToAstra = isModelAvailable('gpt-6-astra')
          && storedVersion <= STORAGE_VERSION && !astraMigrationApplied;
        const storedModel = migrateToAstra && parsed.model.startsWith('gpt-5.6-')
          ? 'gpt-6-astra'
          : parsed.model;
        const selectedModel = isModelAvailable(storedModel) ? storedModel : fallbackModel;
        const reasoningEffort = normalizeReasoningEffort(
          selectedModel, 'reasoningEffort' in parsed ? parsed.reasoningEffort : undefined,
        );
        const featureOverrides = normalizeFeatureOverrides(
          'featureOverrides' in parsed ? parsed.featureOverrides : undefined, migrateToAstra,
        );
        const recommended = getRecommendedModelSettings();
        let settings: ModelSettings = {
          model: selectedModel,
          reasoningEffort,
          featureOverrides,
        };

        if (
          storedVersion < 2
          && selectedModel === recommended.model
          && reasoningEffort === recommended.reasoningEffort
          && Object.keys(featureOverrides).length === 0
        ) {
          settings = recommended;
        }

        if (migrateToAstra) saveSettings(settings);
        return settings;
      }
    }
  } catch (e) {
    console.warn('Failed to load model settings:', e);
  }
  return getRecommendedModelSettings();
}

/**
 * Save settings to localStorage
 */
function saveSettings(settings: ModelSettings): void {
  if (typeof localStorage === 'undefined') return;
  try {
    localStorage.setItem(STORAGE_KEY, JSON.stringify({
      version: STORAGE_VERSION,
      ...(isModelAvailable('gpt-6-astra') ? { astraMigrationVersion: ASTRA_MIGRATION_VERSION } : {}),
      ...settings,
    }));
  } catch (e) {
    console.warn('Failed to save model settings:', e);
  }
}

// Global state for non-hook access
let currentSettings: ModelSettings = loadSettings();
const listeners: Set<(settings: ModelSettings) => void> = new Set();

function notifyListeners() {
  listeners.forEach(listener => listener(currentSettings));
}

/**
 * Get current model settings (non-hook version for services)
 */
export function getModelSettings(): ModelSettings {
  return { ...currentSettings };
}

/**
 * Get model settings for a specific feature
 * Returns the feature-specific override if set, otherwise returns default settings
 */
export function getModelSettingsForFeature(feature: FeatureType): { model: ModelType; reasoningEffort: ReasoningEffort } {
  const settings = getModelSettings();
  const override = settings.featureOverrides?.[feature];
  
  if (override) {
    const config = MODEL_CONFIG[override.model];
    return {
      model: override.model,
      // For reasoning models, use override reasoning or fall back to default
      // For non-reasoning models, reasoning effort doesn't matter but include it for consistency
      reasoningEffort: config.isReasoning 
        ? normalizeReasoningEffort(override.model, override.reasoningEffort || settings.reasoningEffort)
        : settings.reasoningEffort
    };
  }
  
  // No override, use default settings
  return {
    model: settings.model,
    reasoningEffort: settings.reasoningEffort
  };
}

/**
 * Update feature-specific model override
 */
export function updateFeatureOverride(feature: FeatureType, override: FeatureModelOverride | null): void {
  const newOverrides = { ...currentSettings.featureOverrides };
  
  if (override === null) {
    delete newOverrides[feature];
  } else {
    newOverrides[feature] = {
      ...override,
      reasoningEffort: override.reasoningEffort === undefined
        ? undefined
        : normalizeReasoningEffort(override.model, override.reasoningEffort),
    };
  }
  
  updateModelSettings({ featureOverrides: newOverrides });
}

/**
 * Check if a feature has a custom override set
 */
export function hasFeatureOverride(feature: FeatureType): boolean {
  return !!currentSettings.featureOverrides?.[feature];
}

/**
 * Update model settings (non-hook version for services)
 */
export function updateModelSettings(updates: Partial<ModelSettings>): void {
  const nextSettings = { ...currentSettings, ...updates };
  currentSettings = {
    ...nextSettings,
    reasoningEffort: normalizeReasoningEffort(nextSettings.model, nextSettings.reasoningEffort),
    featureOverrides: normalizeFeatureOverrides(nextSettings.featureOverrides),
  };
  saveSettings(currentSettings);
  notifyListeners();
}

/**
 * React hook for model settings
 * Provides reactive updates when settings change
 */
export function useModelSettings(): [ModelSettings, (updates: Partial<ModelSettings>) => void] {
  const [settings, setSettings] = useState<ModelSettings>(currentSettings);

  useEffect(() => {
    const listener = (newSettings: ModelSettings) => {
      setSettings({ ...newSettings });
    };
    listeners.add(listener);
    return () => {
      listeners.delete(listener);
    };
  }, []);

  const update = useCallback((updates: Partial<ModelSettings>) => {
    updateModelSettings(updates);
  }, []);

  return [settings, update];
}

/**
 * Check if a model is available (has deployment configured)
 */
export function isModelAvailable(model: ModelType): boolean {
  try {
    getDeploymentName(model);
    if (
      MODEL_CONFIG[model].apiFormat === 'anthropic-messages'
      && !import.meta.env.VITE_AZURE_FOUNDRY_ENDPOINT
    ) {
      return false;
    }
    return true;
  } catch {
    return false;
  }
}

/**
 * Get list of available models
 */
export function getAvailableModels(): ModelType[] {
  return (Object.keys(MODEL_CONFIG) as ModelType[]).filter(isModelAvailable);
}
