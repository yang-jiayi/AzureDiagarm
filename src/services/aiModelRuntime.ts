// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

import {
  AIModelConfigurationError,
  assertSupportedModel,
  getDeploymentName,
  getModelSettingsForFeature,
  getSupportedReasoningEfforts,
  isModelAvailable,
  isReasoningEffort,
  MODEL_CONFIG,
  type FeatureType,
  type ModelType,
  type ReasoningEffort,
} from '../stores/modelSettingsStore';
import {
  getBYOAIConnectionState,
  getBYOAIProviderLabel,
  getBYOAISelectionRevision,
  getBYOAISettings,
  getBYOAIStorageError,
  BYOAIStorageError,
  validateBYOAIProfile,
  type BYOAIAPIFormat,
  type BYOAIProfile,
} from '../stores/byoAISettingsStore';
import { getRuntimeConfigSnapshot, isBYOAIEnabledOnServer } from './runtimeConfig';

export interface RuntimeModelOverride {
  model: ModelType;
  reasoningEffort: ReasoningEffort;
  /** Opaque, nonsecret submission snapshot. Preserve this object when adding signal/callbacks. */
  connection?: CapturedAIConnection;
}

export interface EffectiveAIModelInfo {
  source: 'managed' | 'bring-your-own';
  model: string;
  displayName: string;
  profileId?: string;
  apiFormat: BYOAIAPIFormat;
  isReasoning: boolean;
  supportsVision: boolean;
  maxCompletionTokens: number;
  reasoningEffort: ReasoningEffort;
  ready: boolean;
  code?: string;
}

export interface CapturedAIConnection extends Readonly<Omit<EffectiveAIModelInfo, 'ready' | 'code'>> {
  readonly feature: FeatureType;
  readonly selectionRevision: number;
  readonly revision: number;
  readonly deployment: string;
}

export interface AIModelRuntime extends Omit<EffectiveAIModelInfo, 'ready' | 'code'> {
  deployment: string;
  telemetryModel: string;
  connection: CapturedAIConnection;
}

const captures = new WeakSet<object>();
const STALE_MESSAGE = 'The selected AI connection changed before this request was sent. Review the connection and submit again.';

function configurationError(code: string): AIModelConfigurationError {
  if (code === 'byo_settings_save_failed' || code === 'byo_settings_read_failed') return new BYOAIStorageError(code);
  const messages: Record<string, string> = {
    byo_profile_missing: 'The selected AI connection is missing. Select another profile or managed GPT-6 Astra.',
    byo_key_required: 'Enter this profile’s API key and test the connection in this browser tab.',
    byo_unverified: 'Test this AI connection successfully before using it.',
    byo_test_failed: 'This AI connection failed its test. Check its settings and key, then test again.',
    byo_not_enabled: 'Bring-your-own AI is disabled by the application administrator.',
    byo_availability_unknown: 'Bring-your-own AI availability has not been confirmed by the application server.',
    byo_invalid_profile: 'The selected AI connection has invalid settings. Edit the profile and test again.',
    astra_not_configured: 'Azure OpenAI is not configured. Please check your environment.',
    stale_ai_configuration: STALE_MESSAGE,
  };
  return new AIModelConfigurationError(code, messages[code] || 'The AI connection is not ready.');
}

function profileInfo(profile: BYOAIProfile): Omit<EffectiveAIModelInfo, 'ready' | 'code'> {
  return {
    source: 'bring-your-own',
    profileId: profile.id,
    model: profile.model,
    displayName: `BYO ${getBYOAIProviderLabel(profile.provider)} · ${profile.name} · ${profile.model}`,
    apiFormat: profile.apiFormat,
    isReasoning: profile.isReasoning,
    supportsVision: profile.supportsVision,
    maxCompletionTokens: profile.maxCompletionTokens,
    reasoningEffort: profile.reasoningEffort,
  };
}

/** Safe display metadata; never resolves or returns a key, and never substitutes managed for BYO. */
export function getEffectiveAIModelInfo(feature: FeatureType): EffectiveAIModelInfo {
  const settings = getBYOAISettings();
  const storageError = getBYOAIStorageError();
  if (settings.activeProfileId !== null) {
    const profile = settings.profiles.find(item => item.id === settings.activeProfileId);
    if (!profile) return {
      source: 'bring-your-own', profileId: settings.activeProfileId, model: '',
      displayName: 'Missing BYO connection', apiFormat: 'responses', isReasoning: false,
      supportsVision: false, maxCompletionTokens: 0, reasoningEffort: 'none',
      ready: false, code: storageError?.code ?? 'byo_profile_missing',
    };
    const state = getBYOAIConnectionState(profile.id);
    const capability = getRuntimeConfigSnapshot();
    const code = storageError?.code ?? (capability.status !== 'ready' ? 'byo_availability_unknown'
      : !capability.bringYourOwnAI ? 'byo_not_enabled'
      : !validateBYOAIProfile(profile).valid ? 'byo_invalid_profile'
      : !state.hasApiKey ? 'byo_key_required'
      : state.status === 'failed' ? 'byo_test_failed'
      : !state.verified ? 'byo_unverified' : undefined);
    return { ...profileInfo(profile), ready: code === undefined, ...(code ? { code } : {}) };
  }
  const model = getModelSettingsForFeature(feature);
  const config = MODEL_CONFIG['gpt-6-astra'];
  const ready = !storageError && isModelAvailable('gpt-6-astra');
  return {
    source: 'managed', model: 'gpt-6-astra', displayName: config.displayName,
    apiFormat: 'responses', isReasoning: true, supportsVision: true,
    maxCompletionTokens: config.maxCompletionTokens, reasoningEffort: model.reasoningEffort,
    ready, ...(!ready ? { code: storageError?.code ?? 'astra_not_configured' } : {}),
  };
}

function validateOverride(override: RuntimeModelOverride): void {
  assertSupportedModel(override?.model);
  if (!isReasoningEffort(override.reasoningEffort)) {
    throw new AIModelConfigurationError('unsupported_reasoning_effort', 'The requested reasoning effort is not supported.');
  }
  if (Object.keys(override).some(key => !['model', 'reasoningEffort', 'connection', 'signal', 'onRetryWait'].includes(key))) {
    throw new AIModelConfigurationError('unsupported_ai_provider', 'Select an AI connection through the connection settings.');
  }
}

/** Capture before asynchronous preparation/admission. Spreading the override keeps the opaque connection intact. */
export function captureRuntimeModelOverride(feature: FeatureType, override?: RuntimeModelOverride): RuntimeModelOverride {
  if (override !== undefined) {
    validateOverride(override);
    if (override.connection !== undefined) {
      assertCapturedAIConnectionCurrent(override.connection);
      if (override.reasoningEffort !== override.connection.reasoningEffort) throw configurationError('stale_ai_configuration');
      return { ...override };
    }
  }
  const info = getEffectiveAIModelInfo(feature);
  if (!info.ready) throw configurationError(info.code!);
  const reasoningEffort = info.source === 'managed'
    ? (override?.reasoningEffort ?? info.reasoningEffort) : info.reasoningEffort;
  if (info.source === 'managed' && !getSupportedReasoningEfforts('gpt-6-astra').includes(reasoningEffort)) {
    throw new AIModelConfigurationError('unsupported_reasoning_effort', 'The requested reasoning effort is not supported by GPT-6 Astra.');
  }
  const { ready: _ready, code: _code, ...publicInfo } = info;
  const connection: CapturedAIConnection = Object.freeze({
    ...publicInfo, reasoningEffort, feature,
    selectionRevision: getBYOAISelectionRevision(),
    revision: info.profileId ? getBYOAIConnectionState(info.profileId).revision : 0,
    deployment: info.source === 'managed' ? getDeploymentName('gpt-6-astra') : info.model,
  });
  captures.add(connection);
  return { ...override, model: 'gpt-6-astra', reasoningEffort, connection };
}

/** Rechecked immediately before *each* dispatch, including a throttled retry. */
export function assertCapturedAIConnectionCurrent(connection: CapturedAIConnection): void {
  const storageError = getBYOAIStorageError();
  if (storageError) throw configurationError(storageError.code);
  if (!connection || typeof connection !== 'object' || !captures.has(connection)
    || connection.selectionRevision !== getBYOAISelectionRevision()) throw configurationError('stale_ai_configuration');
  const settings = getBYOAISettings();
  if (connection.source === 'bring-your-own') {
    if (!connection.profileId || settings.activeProfileId !== connection.profileId
      || getBYOAIConnectionState(connection.profileId).revision !== connection.revision) {
      throw configurationError('stale_ai_configuration');
    }
    if (!isBYOAIEnabledOnServer()) throw configurationError(
      getRuntimeConfigSnapshot().status === 'ready' ? 'byo_not_enabled' : 'byo_availability_unknown',
    );
    if (!getBYOAIConnectionState(connection.profileId).verified) throw configurationError('byo_unverified');
  } else {
    if (settings.activeProfileId !== null) throw configurationError('stale_ai_configuration');
    if (!isModelAvailable('gpt-6-astra')) throw configurationError('astra_not_configured');
    if (connection.deployment !== getDeploymentName('gpt-6-astra')) throw configurationError('stale_ai_configuration');
  }
}

export function resolveAIModelRuntime(feature: FeatureType, override?: RuntimeModelOverride): AIModelRuntime {
  const capture = captureRuntimeModelOverride(feature, override);
  const connection = capture.connection!;
  const profile = connection.profileId
    ? getBYOAISettings().profiles.find(item => item.id === connection.profileId) : undefined;
  return {
    source: connection.source, model: connection.model, displayName: connection.displayName,
    ...(connection.profileId ? { profileId: connection.profileId } : {}),
    deployment: connection.deployment, apiFormat: connection.apiFormat,
    isReasoning: connection.isReasoning, supportsVision: connection.supportsVision,
    maxCompletionTokens: connection.maxCompletionTokens, reasoningEffort: connection.reasoningEffort,
    telemetryModel: profile ? `BYO ${getBYOAIProviderLabel(profile.provider)}` : MODEL_CONFIG['gpt-6-astra'].displayName,
    connection,
  };
}

export function isManagedAIModelConfigured(): boolean {
  return isModelAvailable('gpt-6-astra');
}

export function isAnyAIModelConfigured(): boolean {
  return getEffectiveAIModelInfo('architectureGeneration').ready;
}
