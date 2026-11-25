// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

import { useState, useEffect, useCallback } from 'react';

/** Execution identifiers only. Historical diagram/review model names are strings. */
export type ModelType = 'gpt-6.2-astra' | 'gpt-6.2-sol';
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
const GPT62_REASONING_EFFORTS = Object.freeze(['low', 'medium', 'high', 'xhigh', 'max'] as const);

export type FeatureType = 'architectureGeneration' | 'validation' | 'deploymentGuide' | 'blueprint';
export interface FeatureModelOverride {
  model: ModelType;
  reasoningEffort?: ReasoningEffort;
}
export interface ModelSettings {
  model: ModelType;
  reasoningEffort: ReasoningEffort;
  featureOverrides?: Partial<Record<FeatureType, FeatureModelOverride>>;
}

const STORAGE_KEY = 'azure-diagrams-model-settings';
const STORAGE_VERSION = 6;

export class AIModelConfigurationError extends Error {
  readonly source = 'client';
  readonly retryable = false;

  constructor(public readonly code: string, message: string) {
    super(message);
    this.name = 'AIModelConfigurationError';
  }
}

export function isModelType(value: unknown): value is ModelType {
  return value === 'gpt-6.2-astra' || value === 'gpt-6.2-sol';
}

function migrateModelType(value: unknown): ModelType {
  if (isModelType(value)) return value;
  if (value === 'gpt-6-astra') return 'gpt-6.2-astra';
  return 'gpt-6.2-sol';
}

export function assertSupportedModel(value: unknown): asserts value is ModelType {
  if (!isModelType(value)) {
    throw new AIModelConfigurationError('unsupported_ai_model', 'Select managed GPT-6.2 Astra or GPT-6.2 Sol.');
  }
}

export const FEATURE_CONFIG: Readonly<Record<FeatureType, {
  displayName: string;
  description: string;
  recommendedModel: ModelType;
  recommendedReasoning?: ReasoningEffort;
}>> = {
  architectureGeneration: {
    displayName: 'Architecture Generation', description: 'Creating Azure architecture diagrams',
    recommendedModel: 'gpt-6.2-sol', recommendedReasoning: 'low',
  },
  validation: {
    displayName: 'Architecture Validation', description: 'WAF validation and security analysis',
    recommendedModel: 'gpt-6.2-sol', recommendedReasoning: 'low',
  },
  deploymentGuide: {
    displayName: 'Deployment Guide & Bicep', description: 'Generating deployment guides and IaC templates',
    recommendedModel: 'gpt-6.2-sol', recommendedReasoning: 'low',
  },
  blueprint: {
    displayName: 'Blueprint Diagrams', description: 'Whiteboard-style blueprint sketches',
    recommendedModel: 'gpt-6.2-sol', recommendedReasoning: 'low',
  },
};

export const MODEL_CONFIG = Object.freeze({
  'gpt-6.2-astra': Object.freeze({
    displayName: 'GPT-6.2 Astra',
    deploymentEnvVar: 'VITE_AZURE_OPENAI_DEPLOYMENT_GPT62ASTRA',
    isReasoning: true,
    maxCompletionTokens: 32000,
    maxEffortCompletionTokens: 128000,
    description: 'Azure OpenAI model for architecture design, validation, and deployment guidance',
    recommendedUse: 'Selectable for all AI features',
    defaultReasoningEffort: 'low',
    supportedReasoningEfforts: GPT62_REASONING_EFFORTS,
    apiFormat: 'responses',
    supportsVision: true,
  } as const),
  'gpt-6.2-sol': Object.freeze({
    displayName: 'GPT-6.2 Sol',
    deploymentEnvVar: 'VITE_AZURE_OPENAI_DEPLOYMENT_GPT62SOL',
    isReasoning: true,
    maxCompletionTokens: 32000,
    maxEffortCompletionTokens: 128000,
    description: 'Azure OpenAI model for architecture design, validation, and deployment guidance',
    recommendedUse: 'Default for all AI features',
    defaultReasoningEffort: 'low',
    supportedReasoningEfforts: GPT62_REASONING_EFFORTS,
    apiFormat: 'responses',
    supportsVision: true,
  } as const),
});

export function getManagedOutputTokenLimit(model: ModelType, effort: ReasoningEffort): number {
  const config = MODEL_CONFIG[model];
  return effort === 'max' ? config.maxEffortCompletionTokens : config.maxCompletionTokens;
}

export function isReasoningEffort(value: unknown): value is ReasoningEffort {
  return REASONING_EFFORT_OPTIONS.some(option => option.value === value);
}

export function getSupportedReasoningEfforts(model: ModelType): readonly ReasoningEffort[] {
  assertSupportedModel(model);
  return MODEL_CONFIG[model].supportedReasoningEfforts;
}

export function getReasoningEffortLabel(effort: ReasoningEffort): (typeof REASONING_EFFORT_OPTIONS)[number]['label'] {
  return REASONING_EFFORT_OPTIONS.find(option => option.value === effort)?.label ?? 'Medium';
}

export function normalizeReasoningEffort(model: ModelType, effort: unknown): ReasoningEffort {
  const supported = getSupportedReasoningEfforts(model);
  return isReasoningEffort(effort) && supported.includes(effort)
    ? effort : MODEL_CONFIG[model].defaultReasoningEffort;
}

function record(value: unknown): value is Record<string, unknown> {
  return value !== null && typeof value === 'object' && !Array.isArray(value);
}

function normalizeFeatureOverrides(value: unknown, migrate = false): Partial<Record<FeatureType, FeatureModelOverride>> {
  if (!record(value)) return {};
  const normalized: Partial<Record<FeatureType, FeatureModelOverride>> = {};
  for (const feature of Object.keys(FEATURE_CONFIG) as FeatureType[]) {
    const raw = value[feature];
    if (!record(raw)) continue;
    if (!migrate) assertSupportedModel(raw.model);
    const model = migrate ? migrateModelType(raw.model) : raw.model as ModelType;
    normalized[feature] = {
      model,
      ...(raw.reasoningEffort === undefined ? {} : {
        reasoningEffort: normalizeReasoningEffort(model, raw.reasoningEffort),
      }),
    };
  }
  return normalized;
}

/** Keep literal env access: computed keys would expose unrelated VITE values. */
export function getDeploymentNames(): Record<ModelType, string | undefined> {
  let astra: string | undefined;
  let sol: string | undefined;
  try {
    astra = import.meta.env.VITE_AZURE_OPENAI_DEPLOYMENT_GPT62ASTRA;
    sol = import.meta.env.VITE_AZURE_OPENAI_DEPLOYMENT_GPT62SOL;
  } catch {
    // Non-Vite consumers have no import.meta.env; availability stays closed.
  }
  return { 'gpt-6.2-astra': astra?.trim() || undefined, 'gpt-6.2-sol': sol?.trim() || undefined };
}

export function getDeploymentName(model: ModelType): string {
  assertSupportedModel(model);
  const deployment = getDeploymentNames()[model];
  if (!deployment) {
    throw new AIModelConfigurationError(
      'managed_model_not_configured',
      `No deployment configured for ${MODEL_CONFIG[model].displayName}. Set ${MODEL_CONFIG[model].deploymentEnvVar}.`,
    );
  }
  return deployment;
}

export function isModelAvailable(model: ModelType): boolean {
  if (!isModelType(model)) return false;
  try {
    return Boolean(import.meta.env.VITE_AZURE_OPENAI_ENDPOINT?.trim()) && Boolean(getDeploymentNames()[model]);
  } catch {
    return false;
  }
}

export function getAvailableModels(): ModelType[] {
  return (Object.keys(MODEL_CONFIG) as ModelType[]).filter(isModelAvailable);
}

export function getRecommendedModelSettings(): ModelSettings {
  return { model: 'gpt-6.2-sol', reasoningEffort: 'low', featureOverrides: {} };
}

function saveSettings(settings: ModelSettings): void {
  if (typeof localStorage === 'undefined') return;
  try {
    localStorage.setItem(STORAGE_KEY, JSON.stringify({
      version: STORAGE_VERSION, ...settings,
    }));
  } catch (error) {
    console.warn('Failed to save model settings:', error);
  }
}

function loadSettings(): ModelSettings {
  try {
    const stored = typeof localStorage === 'undefined' ? null : localStorage.getItem(STORAGE_KEY);
    if (stored) {
      const parsed: unknown = JSON.parse(stored);
      if (record(parsed)) {
        // Storage is preference migration, never authority for execution.
        // Retired execution preferences migrate, historical diagrams do not.
        const model = migrateModelType(parsed.model);
        const settings: ModelSettings = {
          model,
          reasoningEffort: normalizeReasoningEffort(model, parsed.reasoningEffort),
          featureOverrides: normalizeFeatureOverrides(parsed.featureOverrides, true),
        };
        if (parsed.version !== STORAGE_VERSION
          || parsed.model !== settings.model || parsed.reasoningEffort !== settings.reasoningEffort
          || JSON.stringify(parsed.featureOverrides) !== JSON.stringify(settings.featureOverrides)) {
          saveSettings(settings);
        }
        return settings;
      }
    }
  } catch (error) {
    console.warn('Failed to load model settings:', error);
  }
  return getRecommendedModelSettings();
}

let currentSettings = loadSettings();
let selectionRevision = 0;
export function getManagedAISelectionRevision(): number { return selectionRevision; }
const listeners = new Set<(settings: ModelSettings) => void>();

export function getModelSettings(): ModelSettings {
  return {
    ...currentSettings,
    featureOverrides: Object.fromEntries(Object.entries(currentSettings.featureOverrides ?? {})
      .map(([feature, override]) => [feature, { ...override }])),
  };
}

export function getModelSettingsForFeature(feature: FeatureType): { model: ModelType; reasoningEffort: ReasoningEffort } {
  const override = currentSettings.featureOverrides?.[feature];
  return {
    model: override?.model ?? currentSettings.model,
    reasoningEffort: normalizeReasoningEffort(override?.model ?? currentSettings.model,
      override?.reasoningEffort ?? currentSettings.reasoningEffort),
  };
}

export function updateFeatureOverride(feature: FeatureType, override: FeatureModelOverride | null): void {
  if (override !== null) assertSupportedModel(override.model);
  const overrides = { ...currentSettings.featureOverrides };
  if (override === null) delete overrides[feature];
  else overrides[feature] = { ...override };
  updateModelSettings({ featureOverrides: overrides });
}

export function hasFeatureOverride(feature: FeatureType): boolean {
  return !!currentSettings.featureOverrides?.[feature];
}

export function updateModelSettings(updates: Partial<ModelSettings>): void {
  if (Object.prototype.hasOwnProperty.call(updates, 'model')) assertSupportedModel(updates.model);
  const next = { ...currentSettings, ...updates };
  const normalized: ModelSettings = {
    model: next.model,
    reasoningEffort: normalizeReasoningEffort(next.model, next.reasoningEffort),
    featureOverrides: normalizeFeatureOverrides(next.featureOverrides),
  };
  currentSettings = normalized;
  selectionRevision++;
  saveSettings(currentSettings);
  listeners.forEach(listener => listener(getModelSettings()));
}

export function useModelSettings(): [ModelSettings, (updates: Partial<ModelSettings>) => void] {
  const [settings, setSettings] = useState<ModelSettings>(getModelSettings);
  useEffect(() => {
    const listener = (next: ModelSettings) => setSettings(next);
    listeners.add(listener);
    return () => { listeners.delete(listener); };
  }, []);
  const update = useCallback((updates: Partial<ModelSettings>) => updateModelSettings(updates), []);
  return [settings, update];
}
