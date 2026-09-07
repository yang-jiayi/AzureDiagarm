// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

import { useEffect, useState } from 'react';
import { AIModelConfigurationError, isReasoningEffort, type ReasoningEffort } from './modelSettingsStore';
import {
  clearBYOAIConnectionSessions,
  getBYOAIConnectionSessionState,
  invalidateBYOAIConnection,
  isValidBYOAIApiKey,
  removeBYOAIConnectionSession,
  setBYOAIConnectionSecret,
  subscribeBYOAIConnectionSessions,
  type BYOAIConnectionState,
} from '../services/byoAIConnectionSession';
import { isBYOAIEnabledOnServer } from '../services/runtimeConfig';

export type { BYOAIConnectionState } from '../services/byoAIConnectionSession';
export type BYOAIProvider = 'azure-openai' | 'openai';
export type BYOAIAPIFormat = 'responses' | 'chat-completions';
export const MAX_BYO_AI_PROFILES = 10;
export const BYOAI_REASONING_EFFORTS = [
  'none', 'minimal', 'low', 'medium', 'high', 'xhigh', 'max',
] as const satisfies readonly ReasoningEffort[];

export interface BYOAIProfile {
  id: string;
  name: string;
  provider: BYOAIProvider;
  endpoint: string;
  model: string;
  apiFormat: BYOAIAPIFormat;
  reasoningEffort: ReasoningEffort;
  isReasoning: boolean;
  supportsVision: boolean;
  maxCompletionTokens: number;
}

export interface BYOAISettings {
  profiles: BYOAIProfile[];
  /** A deleted/malformed selected profile remains selected until an explicit switch. */
  activeProfileId: string | null;
}

export interface BYOAIProfileValidationResult {
  valid: boolean;
  profile?: BYOAIProfile;
  errors: Partial<Record<keyof BYOAIProfile, string>>;
}

export const BYOAI_STORAGE_ERROR_MESSAGES = Object.freeze({
  byo_settings_save_failed: 'AI connection settings could not be saved. Allow site storage or free browser storage, then try again.',
  byo_settings_read_failed: 'AI connection settings could not be loaded safely. Allow site storage, then reload or explicitly save a connection.',
});
export type BYOAIStorageErrorCode = keyof typeof BYOAI_STORAGE_ERROR_MESSAGES;
export interface BYOAIStorageFailure {
  code: BYOAIStorageErrorCode;
  source: 'client';
  message: string;
}

export class BYOAIStorageError extends AIModelConfigurationError {
  constructor(code: BYOAIStorageErrorCode) {
    super(code, BYOAI_STORAGE_ERROR_MESSAGES[code]);
    this.name = 'BYOAIStorageError';
  }
}

export const DEFAULT_BYO_AI_PROFILE: Readonly<Omit<BYOAIProfile, 'id'>> = Object.freeze({
  name: 'My AI connection',
  provider: 'azure-openai',
  endpoint: '',
  model: 'gpt-6-astra',
  apiFormat: 'responses',
  reasoningEffort: 'low',
  isReasoning: true,
  supportsVision: true,
  maxCompletionTokens: 32000,
});

const STORAGE_KEY = 'azure-diagrams-byo-ai-settings';
const STORAGE_VERSION = 3;
const OPENAI_ENDPOINT = 'https://api.openai.com';
const ID_RE = /^[A-Za-z0-9][A-Za-z0-9_-]{0,79}$/;
const MODEL_RE = /^(?=.{1,128}$)(?=.*[A-Za-z0-9])[A-Za-z0-9._:-]+$/;
const AZURE_SUFFIXES = [
  '.openai.azure.com', '.openai.azure.us', '.openai.azure.cn',
  '.cognitiveservices.azure.com', '.cognitiveservices.azure.us', '.cognitiveservices.azure.cn',
  '.services.ai.azure.com', '.services.ai.azure.us', '.services.ai.azure.cn',
];
const record = (value: unknown): value is Record<string, unknown> =>
  value !== null && typeof value === 'object' && !Array.isArray(value);
const text = (value: unknown, fallback = '') => typeof value === 'string' ? value.trim() : fallback;

export function normalizeBYOAIEndpoint(provider: BYOAIProvider, value: unknown): string | null {
  if (provider !== 'azure-openai' && provider !== 'openai') return null;
  if (typeof value !== 'string') return null;
  const endpoint = value.trim();
  // Check the original authority too: URL normalizes away empty ?, # and @.
  if (!/^https:\/\/[a-z0-9.-]+\/?$/i.test(endpoint)) return null;
  let url: URL;
  try { url = new URL(endpoint); } catch { return null; }
  if (url.protocol !== 'https:' || url.port || url.username || url.password
    || url.search || url.hash || url.pathname !== '/') return null;
  if (provider === 'openai') return url.origin === OPENAI_ENDPOINT ? OPENAI_ENDPOINT : null;
  const trusted = AZURE_SUFFIXES.some(suffix => {
    if (!url.hostname.endsWith(suffix)) return false;
    const resource = url.hostname.slice(0, -suffix.length);
    return /^[a-z0-9](?:[a-z0-9-]{0,61}[a-z0-9])?$/.test(resource);
  });
  return trusted ? url.origin : null;
}

function publicProfile(raw: Record<string, unknown>, id: string): BYOAIProfile {
  const provider = raw.provider === 'openai' ? 'openai' : 'azure-openai';
  return {
    id,
    name: text(raw.name, DEFAULT_BYO_AI_PROFILE.name).slice(0, 80),
    provider,
    endpoint: normalizeBYOAIEndpoint(provider, text(raw.endpoint, provider === 'openai' ? OPENAI_ENDPOINT : '')) ?? '',
    model: text(raw.model),
    apiFormat: raw.apiFormat === 'chat-completions' ? 'chat-completions' : 'responses',
    reasoningEffort: isReasoningEffort(raw.reasoningEffort) ? raw.reasoningEffort : 'low',
    isReasoning: raw.isReasoning !== false,
    supportsVision: raw.supportsVision !== false,
    maxCompletionTokens: typeof raw.maxCompletionTokens === 'number'
      ? raw.maxCompletionTokens : DEFAULT_BYO_AI_PROFILE.maxCompletionTokens,
  };
}

export function validateBYOAIProfile(value: unknown): BYOAIProfileValidationResult {
  const raw = record(value) ? value : {};
  const errors: BYOAIProfileValidationResult['errors'] = {};
  const id = text(raw.id);
  if (raw.id !== undefined && !ID_RE.test(id)) errors.id = 'Use a valid connection profile ID.';
  if (typeof raw.name !== 'string' || !raw.name.trim() || raw.name.trim().length > 80
    || /[\u0000-\u001f\u007f]/.test(raw.name)) errors.name = 'Enter a connection name of 1–80 characters.';
  if (raw.provider !== 'azure-openai' && raw.provider !== 'openai') errors.provider = 'Select Azure OpenAI or OpenAI.';
  if (raw.apiFormat !== 'responses' && raw.apiFormat !== 'chat-completions') errors.apiFormat = 'Select Responses or Chat Completions.';
  if (typeof raw.model !== 'string' || !MODEL_RE.test(raw.model.trim())) {
    errors.model = 'Enter a deployment or model identifier containing 1–128 letters, numbers, dots, hyphens, underscores, or colons.';
  }
  if (!isReasoningEffort(raw.reasoningEffort)) errors.reasoningEffort = 'Select a supported reasoning effort.';
  if (typeof raw.isReasoning !== 'boolean') errors.isReasoning = 'Specify whether the model supports reasoning.';
  if (typeof raw.supportsVision !== 'boolean') errors.supportsVision = 'Specify whether the model supports images.';
  if (typeof raw.maxCompletionTokens !== 'number' || !Number.isSafeInteger(raw.maxCompletionTokens)
    || raw.maxCompletionTokens < 1 || raw.maxCompletionTokens > 32768) {
    errors.maxCompletionTokens = 'Enter a whole-number output limit from 1 to 32768 tokens.';
  }
  const endpoint = normalizeBYOAIEndpoint(raw.provider as BYOAIProvider, raw.endpoint);
  if (!endpoint) errors.endpoint = 'Use an Azure resource HTTPS origin or https://api.openai.com, without credentials, paths, ports, query, or fragment.';
  const valid = Object.keys(errors).length === 0;
  return { valid, ...(valid ? { profile: { ...publicProfile(raw, id), endpoint: endpoint! } } : {}), errors };
}

/** Storage migration is an allowlist, not an object spread (including v1/v2). */
export function normalizeBYOAISettings(value: unknown): BYOAISettings {
  if (!record(value)) return { profiles: [], activeProfileId: null };
  if (!Array.isArray(value.profiles)) {
    if (!('provider' in value || 'model' in value || 'enabled' in value)) {
      return { profiles: [], activeProfileId: value.activeProfileId === null || value.activeProfileId === undefined
        ? null : 'missing-byo-profile' };
    }
    const profile = publicProfile(value, 'byo-migrated');
    profile.name = text(value.name, 'Migrated AI connection').slice(0, 80) || 'Migrated AI connection';
    profile.endpoint = profile.provider === 'openai' ? OPENAI_ENDPOINT
      : normalizeBYOAIEndpoint(profile.provider, profile.endpoint) ?? '';
    return { profiles: [profile], activeProfileId: value.enabled === true ? profile.id : null };
  }
  const profiles: BYOAIProfile[] = [];
  const ids = new Set<string>();
  for (const raw of value.profiles) {
    if (!record(raw)) continue;
    const id = text(raw.id);
    if (!ID_RE.test(id) || ids.has(id) || profiles.length >= MAX_BYO_AI_PROFILES) continue;
    // Never turn an unknown saved provider/format into a runnable one.
    if ((raw.provider !== 'openai' && raw.provider !== 'azure-openai')
      || (raw.apiFormat !== 'responses' && raw.apiFormat !== 'chat-completions')) continue;
    const profile = publicProfile(raw, id);
    profile.endpoint = normalizeBYOAIEndpoint(profile.provider, profile.endpoint) ?? '';
    profiles.push(profile);
    ids.add(id);
  }
  return {
    profiles,
    activeProfileId: value.activeProfileId === null || value.activeProfileId === undefined ? null
      : ID_RE.test(text(value.activeProfileId)) ? text(value.activeProfileId) : 'missing-byo-profile',
  };
}

function serializeSettings(settings: BYOAISettings): string {
  return JSON.stringify({
    version: STORAGE_VERSION,
    profiles: settings.profiles.map(profile => publicProfile(profile as unknown as Record<string, unknown>, profile.id)),
    activeProfileId: settings.activeProfileId,
  });
}

function persist(settings: BYOAISettings): void {
  try {
    if (typeof localStorage === 'undefined') throw new BYOAIStorageError('byo_settings_save_failed');
    localStorage.setItem(STORAGE_KEY, serializeSettings(settings));
  } catch {
    throw new BYOAIStorageError('byo_settings_save_failed');
  }
}

function loadSettings(): { settings: BYOAISettings; failure: BYOAIStorageErrorCode | null } {
  let stored: string | null;
  try {
    if (typeof localStorage === 'undefined') {
      // Non-browser imports have no saved selection. Browser storage failures
      // must instead remain blocked, never an implicit managed selection.
      if (typeof window !== 'undefined') throw new BYOAIStorageError('byo_settings_read_failed');
      return { settings: { profiles: [], activeProfileId: null }, failure: null };
    }
    stored = localStorage.getItem(STORAGE_KEY);
  } catch {
    return { settings: { profiles: [], activeProfileId: 'missing-byo-profile' }, failure: 'byo_settings_read_failed' };
  }
  if (stored === null) return { settings: { profiles: [], activeProfileId: null }, failure: null };
  let settings: BYOAISettings;
  try {
    const parsed: unknown = JSON.parse(stored);
    settings = record(parsed) ? normalizeBYOAISettings(parsed)
      : { profiles: [], activeProfileId: 'missing-byo-profile' };
  } catch {
    settings = { profiles: [], activeProfileId: 'missing-byo-profile' };
  }
  try {
    if (stored !== serializeSettings(settings)) persist(settings);
    return { settings, failure: null };
  } catch {
    // Keep only sanitized recovery fields. A successful test cannot make a
    // failed migration runnable until the public selection is explicitly saved.
    return { settings, failure: 'byo_settings_save_failed' };
  }
}

const initial = loadSettings();
let currentSettings = initial.settings;
let storageFailure = initial.failure;
let selectionRevision = 0;
const listeners = new Set<() => void>();
const notify = () => listeners.forEach(listener => listener());
subscribeBYOAIConnectionSessions(notify);

export function getBYOAISettings(): BYOAISettings {
  return { profiles: currentSettings.profiles.map(profile => ({ ...profile })), activeProfileId: currentSettings.activeProfileId };
}

export function getBYOAISelectionRevision(): number { return selectionRevision; }

export function getBYOAIStorageError(): BYOAIStorageFailure | null {
  return storageFailure ? { code: storageFailure, source: 'client', message: BYOAI_STORAGE_ERROR_MESSAGES[storageFailure] } : null;
}

function commitSettings(settings: BYOAISettings): void {
  // Web Storage writes are synchronous: no observable settings/session changes
  // may precede this required write, which can throw (quota, privacy policy).
  persist(settings);
  currentSettings = settings;
  storageFailure = null;
}

function requestSignature(profile: BYOAIProfile): string {
  const { id: _id, name: _name, ...request } = profile;
  return JSON.stringify(request);
}

export function upsertBYOAIProfile(value: Omit<BYOAIProfile, 'id'> & { id?: string }): BYOAIProfile {
  const result = validateBYOAIProfile(value);
  if (!result.valid || !result.profile) {
    throw new AIModelConfigurationError('byo_invalid_profile', Object.values(result.errors)[0] || 'Complete the connection profile.');
  }
  const profile = { ...result.profile, id: result.profile.id || crypto.randomUUID() };
  const index = currentSettings.profiles.findIndex(item => item.id === profile.id);
  if (index < 0 && currentSettings.profiles.length >= MAX_BYO_AI_PROFILES) {
    throw new AIModelConfigurationError('byo_profile_limit', 'You can save up to 10 AI connection profiles.');
  }
  const previous = currentSettings.profiles[index];
  const nextSettings = {
    ...currentSettings,
    profiles: index < 0 ? [...currentSettings.profiles, profile]
      : currentSettings.profiles.map((item, at) => at === index ? profile : item),
  };
  commitSettings(nextSettings);
  if (!previous || requestSignature(previous) !== requestSignature(profile)) invalidateBYOAIConnection(profile.id);
  notify();
  return { ...profile };
}

export function removeBYOAIProfile(profileId: string): void {
  commitSettings({ ...currentSettings, profiles: currentSettings.profiles.filter(profile => profile.id !== profileId) });
  removeBYOAIConnectionSession(profileId);
  notify();
}

export function getBYOAIConnectionState(profileId: string): BYOAIConnectionState {
  return currentSettings.profiles.some(profile => profile.id === profileId)
    ? getBYOAIConnectionSessionState(profileId)
    : { status: 'missing-profile', hasApiKey: false, verified: false, revision: 0 };
}

/** Invalidate request-affecting draft edits without saving the draft or discarding the existing key. */
export function invalidateBYOAIProfile(profileId: string): void {
  if (!currentSettings.profiles.some(profile => profile.id === profileId)) {
    throw new AIModelConfigurationError('byo_profile_missing', 'The selected AI connection is missing. Select another profile or managed GPT-6 Astra.');
  }
  invalidateBYOAIConnection(profileId);
}

export function setBYOAIApiKey(profileId: string, key: string): void {
  if (!currentSettings.profiles.some(profile => profile.id === profileId)) {
    throw new AIModelConfigurationError('byo_profile_missing', 'The selected AI connection is missing. Select another profile or managed GPT-6 Astra.');
  }
  if (typeof key !== 'string' || (key.trim() && !isValidBYOAIApiKey(key.trim()))) {
    throw new AIModelConfigurationError('byo_invalid_api_key', 'Enter a valid API key. Keys stay only in this browser tab memory.');
  }
  setBYOAIConnectionSecret(profileId, key.trim());
}

export function selectBYOAIProfile(profileId: string | null): void {
  if (profileId !== null) {
    const state = getBYOAIConnectionState(profileId);
    if (state.status === 'missing-profile') {
      throw new AIModelConfigurationError('byo_profile_missing', 'The selected AI connection is missing. Select another profile or managed GPT-6 Astra.');
    }
    if (!isBYOAIEnabledOnServer()) {
      throw new AIModelConfigurationError('byo_not_enabled', 'Bring-your-own AI availability must be confirmed by the application server.');
    }
    if (!state.verified) {
      throw new AIModelConfigurationError(state.hasApiKey ? 'byo_unverified' : 'byo_key_required',
        'Enter this profile’s API key and successfully test the connection before using it.');
    }
  }
  if (currentSettings.activeProfileId === profileId && !storageFailure) return;
  commitSettings({ ...currentSettings, activeProfileId: profileId });
  selectionRevision++;
  notify();
}

export function reloadBYOAISettings(): void {
  const loaded = loadSettings();
  currentSettings = loaded.settings;
  storageFailure = loaded.failure;
  selectionRevision++;
  clearBYOAIConnectionSessions();
  notify();
  if (loaded.failure) throw new BYOAIStorageError(loaded.failure);
}

export function getBYOAIProviderLabel(provider: BYOAIProvider): string {
  return provider === 'openai' ? 'OpenAI' : 'Azure OpenAI';
}

function snapshot() {
  const settings = getBYOAISettings();
  return { settings, storageError: getBYOAIStorageError(), connectionStates: Object.fromEntries(settings.profiles.map(profile =>
    [profile.id, getBYOAIConnectionState(profile.id)])) as Record<string, BYOAIConnectionState> };
}

export function useBYOAISettings(): ReturnType<typeof snapshot> {
  const [value, setValue] = useState(snapshot);
  useEffect(() => {
    const listener = () => setValue(snapshot());
    listeners.add(listener);
    listener();
    return () => { listeners.delete(listener); };
  }, []);
  return value;
}
