// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

/**
 * Model Settings Store
 * Manages AI model selection and reasoning effort preferences
 * Supports per-feature model overrides for optimal results
 * Persists to localStorage for cross-session consistency
 */

import { useState, useEffect, useCallback } from 'react';

export type ModelType = 'gpt-5.1' | 'gpt-5.2' | 'gpt-5.4' | 'gpt-5.4-mini' | 'gpt-5.6-sol' | 'gpt-5.6-terra' | 'gpt-5.6-luna' | 'claude-opus-5' | 'deepseek-v3.2-speciale' | 'deepseek-v4-pro' | 'grok-4.1-fast' | 'grok-4.3' | 'mistral-large-3' | 'kimi-k2-5' | 'kimi-k2-7-code';
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
const GPT_56_REASONING_EFFORTS = [...EXTENDED_REASONING_EFFORTS, 'max'] as const;
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
const STORAGE_VERSION = 2;

const DEFAULT_SETTINGS: ModelSettings = {
  model: 'gpt-5.6-sol',
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
    recommendedModel: 'gpt-5.6-sol',
    recommendedReasoning: 'low'
  },
  validation: {
    displayName: 'Architecture Validation',
    description: 'WAF validation and security analysis',
    recommendedModel: 'gpt-5.6-terra',
    recommendedReasoning: 'low'
  },
  deploymentGuide: {
    displayName: 'Deployment Guide & Bicep',
    description: 'Generating deployment guides and IaC templates',
    recommendedModel: 'gpt-5.6-terra',
    recommendedReasoning: 'low'
  },
  blueprint: {
    displayName: 'Blueprint Diagrams',
    description: 'Whiteboard-style blueprint sketches (fast, cost-efficient)',
    recommendedModel: 'gpt-5.6-luna',
    recommendedReasoning: 'low'
  }
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
    description: 'Newest frontier reasoning model - top-tier quality for complex architectures',
    recommendedUse: 'Highest quality',
    defaultReasoningEffort: 'low',
    supportedReasoningEfforts: GPT_56_REASONING_EFFORTS,
  },
  'gpt-5.6-terra': {
    displayName: 'GPT-5.6 Terra',
    deploymentEnvVar: 'VITE_AZURE_OPENAI_DEPLOYMENT_GPT56TERRA',
    isReasoning: true,
    maxCompletionTokens: 32000,
    description: 'Frontier reasoning model - grounded, thorough analysis for complex architectures',
    recommendedUse: 'Validation + deployment',
    defaultReasoningEffort: 'low',
    supportedReasoningEfforts: GPT_56_REASONING_EFFORTS,
  },
  'gpt-5.6-luna': {
    displayName: 'GPT-5.6 Luna',
    deploymentEnvVar: 'VITE_AZURE_OPENAI_DEPLOYMENT_GPT56LUNA',
    isReasoning: true,
    maxCompletionTokens: 32000,
    description: 'Frontier reasoning model - fast, creative reasoning for architecture design',
    recommendedUse: 'Fast blueprints',
    defaultReasoningEffort: 'low',
    supportedReasoningEfforts: GPT_56_REASONING_EFFORTS,
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
): Partial<Record<FeatureType, FeatureModelOverride>> {
  if (!value || typeof value !== 'object') return {};

  const normalized: Partial<Record<FeatureType, FeatureModelOverride>> = {};
  for (const [feature, rawOverride] of Object.entries(value)) {
    if (!rawOverride || typeof rawOverride !== 'object' || !('model' in rawOverride)) continue;
    const model = (rawOverride as FeatureModelOverride).model;
    if (!MODEL_CONFIG[model] || !isModelAvailable(model)) continue;

    const rawEffort = (rawOverride as FeatureModelOverride).reasoningEffort;
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
 */
export const DEPLOYMENT_NAMES: Record<ModelType, string | undefined> = {
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

/**
 * Get deployment name for a specific model
 * Each model requires its own deployment env var to be set
 */
export function getDeploymentName(model: ModelType): string {
  const config = MODEL_CONFIG[model];

  // Static lookup (see DEPLOYMENT_NAMES note above — do not use a dynamic key).
  const specificDeployment = DEPLOYMENT_NAMES[model];
  if (specificDeployment) {
    return specificDeployment;
  }
  
  // No fallback - each model needs its own deployment configured
  throw new Error(`No deployment configured for ${config.displayName}. Set ${config.deploymentEnvVar} in your .env file.`);
}

/**
 * Build the recommended application portfolio from models that are actually deployed.
 * Sol remains the default, while Terra and Luna are assigned to their strongest features.
 */
export function getRecommendedModelSettings(): ModelSettings {
  const availableModels = getAvailableModels();
  const architectureRecommendation = FEATURE_CONFIG.architectureGeneration;
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
    const recommendation = FEATURE_CONFIG[feature];
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
    const stored = localStorage.getItem(STORAGE_KEY);
    if (stored) {
      const parsed = JSON.parse(stored);
      // Validate model type
      if (parsed.model && MODEL_CONFIG[parsed.model as ModelType]) {
        const storedModel = parsed.model as ModelType;
        const selectedModel = isModelAvailable(storedModel) ? storedModel : fallbackModel;
        const reasoningEffort = normalizeReasoningEffort(selectedModel, parsed.reasoningEffort);
        const featureOverrides = normalizeFeatureOverrides(parsed.featureOverrides);
        const storedVersion = Number.isInteger(parsed.version) ? parsed.version : 1;

        if (
          storedVersion < STORAGE_VERSION
          && selectedModel === DEFAULT_SETTINGS.model
          && reasoningEffort === DEFAULT_SETTINGS.reasoningEffort
          && Object.keys(featureOverrides).length === 0
        ) {
          return getRecommendedModelSettings();
        }

        return {
          model: selectedModel,
          reasoningEffort,
          featureOverrides,
        };
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
  try {
    localStorage.setItem(STORAGE_KEY, JSON.stringify({
      version: STORAGE_VERSION,
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
