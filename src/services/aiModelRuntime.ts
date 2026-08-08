// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

import {
  getDeploymentName,
  getModelSettingsForFeature,
  MODEL_CONFIG,
  type FeatureType,
  type ModelType,
  type ReasoningEffort,
} from '../stores/modelSettingsStore';
import {
  getBYOAIApiKey,
  getBYOAIProviderLabel,
  getBYOAISettings,
  isBYOAIReady,
  validateBYOAISettings,
} from '../stores/byoAISettingsStore';
import {
  isAiBackendConfigured,
  type ApiFormat,
  type BYOAIProxyConfig,
} from './apiHelper';
import {
  getRuntimeConfigSnapshot,
  isBYOAIEnabledOnServer,
} from './runtimeConfig';

export interface RuntimeModelOverride {
  model: ModelType;
  reasoningEffort: ReasoningEffort;
  forceManaged?: boolean;
}

export interface AIModelRuntime {
  source: 'managed' | 'bring-your-own';
  deployment: string;
  displayName: string;
  telemetryModel: string;
  apiFormat: ApiFormat;
  isReasoning: boolean;
  supportsVision: boolean;
  maxCompletionTokens: number;
  reasoningEffort: ReasoningEffort;
  byo?: BYOAIProxyConfig;
}

export function resolveAIModelRuntime(
  feature: FeatureType,
  override?: RuntimeModelOverride,
): AIModelRuntime {
  const featureSettings = override ?? getModelSettingsForFeature(feature);
  const byoSettings = getBYOAISettings();

  if (byoSettings.enabled && !override?.forceManaged) {
    const runtimeConfig = getRuntimeConfigSnapshot();
    if (runtimeConfig.status !== 'ready') {
      throw new Error(
        'Bring-your-own AI availability has not been confirmed by the application server.',
      );
    }
    if (!isBYOAIEnabledOnServer()) {
      throw new Error('Bring-your-own AI is disabled by the application administrator.');
    }
    const apiKey = getBYOAIApiKey();
    const validation = validateBYOAISettings(byoSettings, apiKey);
    if (!validation.valid || !validation.settings) {
      const reason = Object.values(validation.errors)[0] || 'Complete the custom AI settings.';
      throw new Error(`Bring-your-own AI is not ready. ${reason}`);
    }
    const normalized = validation.settings;
    const providerLabel = getBYOAIProviderLabel(normalized.provider);
    return {
      source: 'bring-your-own',
      deployment: normalized.model,
      displayName: `BYO ${providerLabel} · ${normalized.model}`,
      telemetryModel: `BYO ${providerLabel}`,
      apiFormat: normalized.apiFormat,
      isReasoning: normalized.isReasoning,
      supportsVision: normalized.supportsVision,
      maxCompletionTokens: 32_000,
      reasoningEffort: normalized.isReasoning ? normalized.reasoningEffort : 'none',
      byo: {
        provider: normalized.provider,
        endpoint: normalized.endpoint,
        apiKey,
      },
    };
  }

  const modelConfig = MODEL_CONFIG[featureSettings.model];
  const apiFormat = modelConfig.apiFormat || 'responses';
  if (!isAiBackendConfigured(apiFormat)) {
    throw new Error(
      apiFormat === 'anthropic-messages'
        ? 'Microsoft Foundry is not configured. Please check your environment.'
        : 'Azure OpenAI is not configured. Please check your environment.',
    );
  }

  let deployment: string;
  try {
    deployment = getDeploymentName(featureSettings.model);
  } catch {
    throw new Error(
      `No deployment configured for ${modelConfig.displayName}. Please check your environment.`,
    );
  }

  return {
    source: 'managed',
    deployment,
    displayName: modelConfig.displayName,
    telemetryModel: modelConfig.displayName,
    apiFormat,
    isReasoning: modelConfig.isReasoning,
    supportsVision: modelConfig.supportsVision !== false,
    maxCompletionTokens: modelConfig.maxCompletionTokens,
    reasoningEffort: featureSettings.reasoningEffort,
  };
}

export function isManagedAIModelConfigured(): boolean {
  return (Object.keys(MODEL_CONFIG) as ModelType[]).some(model => (
    isAiBackendConfigured(MODEL_CONFIG[model].apiFormat || 'responses')
  ));
}

export function isAnyAIModelConfigured(): boolean {
  if (getBYOAISettings().enabled) {
    return isBYOAIEnabledOnServer() && isBYOAIReady();
  }
  return isManagedAIModelConfigured();
}
