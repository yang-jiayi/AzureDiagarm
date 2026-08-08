// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

import { useEffect, useState } from 'react';
import type { ReasoningEffort } from './modelSettingsStore';

export type BYOAIProvider = 'azure-openai' | 'openai';
export type BYOAIAPIFormat = 'responses' | 'chat-completions';
export type BYOAICapabilityMode = 'auto' | 'manual';
export type BYOAIConnectionState = 'disabled' | 'key-required' | 'unverified' | 'verified';
export const BYOAI_REASONING_EFFORTS = [
  'none',
  'minimal',
  'low',
  'medium',
  'high',
  'xhigh',
] as const satisfies readonly ReasoningEffort[];

export interface BYOAISettings {
  enabled: boolean;
  provider: BYOAIProvider;
  endpoint: string;
  model: string;
  apiFormat: BYOAIAPIFormat;
  reasoningEffort: ReasoningEffort;
  capabilityMode: BYOAICapabilityMode;
  isReasoning: boolean;
  supportsVision: boolean;
}

export interface BYOAIValidationResult {
  valid: boolean;
  settings?: BYOAISettings;
  errors: Partial<Record<'endpoint' | 'model' | 'apiKey', string>>;
}

const STORAGE_KEY = 'azure-diagrams-byo-ai-settings';
const STORAGE_VERSION = 2;
const OFFICIAL_OPENAI_ENDPOINT = 'https://api.openai.com';
const MODEL_NAME_RE = /^(?=.{1,128}$)(?=.*[A-Za-z0-9])[A-Za-z0-9._:-]+$/;
const API_KEY_RE = /^[^\s\r\n]{8,512}$/;
const AZURE_OPENAI_HOST_SUFFIXES = [
  '.openai.azure.com',
  '.openai.azure.us',
  '.openai.azure.cn',
  '.cognitiveservices.azure.com',
  '.cognitiveservices.azure.us',
  '.cognitiveservices.azure.cn',
  '.services.ai.azure.com',
  '.services.ai.azure.us',
  '.services.ai.azure.cn',
];

export const DEFAULT_BYO_AI_SETTINGS: BYOAISettings = {
  enabled: false,
  provider: 'azure-openai',
  endpoint: '',
  model: 'gpt-5.6-sol',
  apiFormat: 'responses',
  reasoningEffort: 'low',
  capabilityMode: 'auto',
  isReasoning: true,
  supportsVision: true,
};

function normalizeAzureEndpoint(value: string): string | null {
  let url: URL;
  try {
    url = new URL(value);
  } catch {
    return null;
  }
  const hostname = url.hostname.toLowerCase();
  const trustedHost = AZURE_OPENAI_HOST_SUFFIXES.some(suffix => (
    hostname.endsWith(suffix) && hostname.length > suffix.length
  ));
  if (
    url.protocol !== 'https:'
    || url.username
    || url.password
    || url.port
    || url.search
    || url.hash
    || (url.pathname !== '/' && url.pathname !== '')
    || !trustedHost
  ) {
    return null;
  }
  return url.origin;
}

function normalizeProvider(value: unknown): BYOAIProvider {
  return value === 'openai' ? 'openai' : 'azure-openai';
}

function normalizeApiFormat(value: unknown): BYOAIAPIFormat {
  return value === 'chat-completions' ? 'chat-completions' : 'responses';
}

function normalizeCapabilityMode(value: unknown, storageVersion: number): BYOAICapabilityMode {
  if (value === 'manual') return 'manual';
  if (value === 'auto') return 'auto';
  return storageVersion < STORAGE_VERSION ? 'manual' : 'auto';
}

function normalizeReasoningEffort(value: unknown): ReasoningEffort {
  return typeof value === 'string' && (BYOAI_REASONING_EFFORTS as readonly string[]).includes(value)
    ? value as ReasoningEffort
    : DEFAULT_BYO_AI_SETTINGS.reasoningEffort;
}

export function inferBYOAICapabilities(model: string): {
  isReasoning: boolean;
  supportsVision: boolean;
} {
  const normalized = model.trim().toLowerCase();
  const reasoningModel = /(?:^|[^a-z0-9])(?:o[134]|gpt[-_.]?5)(?:[^a-z0-9]|$)/.test(normalized);
  const visionModel = reasoningModel
    || /(?:^|[^a-z0-9])(?:gpt[-_.]?4o|gpt[-_.]?4[._-]1|vision)(?:[^a-z0-9]|$)/
      .test(normalized);
  return {
    isReasoning: reasoningModel,
    supportsVision: visionModel,
  };
}

export function applyBYOAIAutomaticCapabilities(settings: BYOAISettings): BYOAISettings {
  if (settings.capabilityMode !== 'auto') return settings;
  return {
    ...settings,
    ...inferBYOAICapabilities(settings.model),
  };
}

export function normalizeBYOAISettings(value: unknown): BYOAISettings {
  const raw = value && typeof value === 'object'
    ? value as Partial<BYOAISettings>
    : {};
  const provider = normalizeProvider(raw.provider);
  const storageVersion = Number.isInteger((raw as { version?: unknown }).version)
    ? Number((raw as { version?: unknown }).version)
    : STORAGE_VERSION;
  return applyBYOAIAutomaticCapabilities({
    enabled: raw.enabled === true,
    provider,
    endpoint: provider === 'openai'
      ? OFFICIAL_OPENAI_ENDPOINT
      : String(raw.endpoint ?? '').trim(),
    model: String(raw.model ?? DEFAULT_BYO_AI_SETTINGS.model).trim(),
    apiFormat: normalizeApiFormat(raw.apiFormat),
    reasoningEffort: normalizeReasoningEffort(raw.reasoningEffort),
    capabilityMode: normalizeCapabilityMode(raw.capabilityMode, storageVersion),
    isReasoning: raw.isReasoning !== false,
    supportsVision: raw.supportsVision !== false,
  });
}

export function validateBYOAISettings(
  value: unknown,
  apiKey: string,
  requireApiKey = true,
): BYOAIValidationResult {
  const settings = normalizeBYOAISettings(value);
  const errors: BYOAIValidationResult['errors'] = {};

  if (settings.provider === 'azure-openai') {
    const endpoint = normalizeAzureEndpoint(settings.endpoint);
    if (!endpoint) {
      errors.endpoint = 'Enter a trusted Azure OpenAI or Microsoft Foundry HTTPS endpoint without an API path.';
    } else {
      settings.endpoint = endpoint;
    }
  } else {
    settings.endpoint = OFFICIAL_OPENAI_ENDPOINT;
  }

  if (!MODEL_NAME_RE.test(settings.model)) {
    errors.model = 'Use a deployment or model name containing at least one letter or number.';
  }

  if (requireApiKey && !API_KEY_RE.test(apiKey.trim())) {
    errors.apiKey = 'Enter a valid API key. It is kept only in this browser tab memory.';
  }

  return {
    valid: Object.keys(errors).length === 0,
    ...(Object.keys(errors).length === 0 ? { settings } : {}),
    errors,
  };
}

function loadSettings(): BYOAISettings {
  if (typeof localStorage === 'undefined') return DEFAULT_BYO_AI_SETTINGS;
  try {
    const stored = localStorage.getItem(STORAGE_KEY);
    if (!stored) return DEFAULT_BYO_AI_SETTINGS;
    const parsed = JSON.parse(stored);
    return normalizeBYOAISettings(parsed);
  } catch (error) {
    console.warn('Failed to load bring-your-own AI settings:', error);
    return DEFAULT_BYO_AI_SETTINGS;
  }
}

function saveSettings(settings: BYOAISettings): void {
  if (typeof localStorage === 'undefined') return;
  try {
    localStorage.setItem(STORAGE_KEY, JSON.stringify({
      version: STORAGE_VERSION,
      ...settings,
    }));
  } catch (error) {
    console.warn('Failed to save bring-your-own AI settings:', error);
  }
}

let currentSettings = loadSettings();
let currentApiKey = '';
let currentConfigurationVerified = false;
const listeners = new Set<() => void>();

function notifyListeners(): void {
  listeners.forEach(listener => listener());
}

export function getBYOAISettings(): BYOAISettings {
  return { ...currentSettings };
}

export function getBYOAIApiKey(): string {
  return currentApiKey;
}

export function hasBYOAIApiKey(): boolean {
  return API_KEY_RE.test(currentApiKey);
}

export function saveBYOAIConfiguration(
  settings: BYOAISettings,
  apiKey?: string,
  options: { verified?: boolean } = {},
): void {
  currentSettings = normalizeBYOAISettings(settings);
  if (apiKey !== undefined) currentApiKey = apiKey.trim();
  currentConfigurationVerified = options.verified === true;
  saveSettings(currentSettings);
  notifyListeners();
}

export function disableBYOAI(): void {
  currentSettings = { ...currentSettings, enabled: false };
  currentApiKey = '';
  currentConfigurationVerified = false;
  saveSettings(currentSettings);
  notifyListeners();
}

export function reloadBYOAISettings(): void {
  currentSettings = loadSettings();
  currentApiKey = '';
  currentConfigurationVerified = false;
  notifyListeners();
}

export function isBYOAIReady(): boolean {
  if (!currentSettings.enabled) return false;
  return validateBYOAISettings(currentSettings, currentApiKey).valid;
}

export function isBYOAIVerified(): boolean {
  return isBYOAIReady() && currentConfigurationVerified;
}

export function getBYOAIConnectionState(): BYOAIConnectionState {
  if (!currentSettings.enabled) return 'disabled';
  if (!hasBYOAIApiKey()) return 'key-required';
  return isBYOAIVerified() ? 'verified' : 'unverified';
}

export function getBYOAIProviderLabel(provider = currentSettings.provider): string {
  return provider === 'openai' ? 'OpenAI' : 'Azure OpenAI';
}

export function getBYOAIModelLabel(): string {
  return `BYO ${getBYOAIProviderLabel()} · ${currentSettings.model}`;
}

export function useBYOAISettings(): {
  settings: BYOAISettings;
  hasApiKey: boolean;
  connectionState: BYOAIConnectionState;
  verified: boolean;
} {
  const [snapshot, setSnapshot] = useState(() => ({
    settings: getBYOAISettings(),
    hasApiKey: hasBYOAIApiKey(),
    connectionState: getBYOAIConnectionState(),
    verified: isBYOAIVerified(),
  }));

  useEffect(() => {
    const listener = () => {
      setSnapshot({
        settings: getBYOAISettings(),
        hasApiKey: hasBYOAIApiKey(),
        connectionState: getBYOAIConnectionState(),
        verified: isBYOAIVerified(),
      });
    };
    listeners.add(listener);
    return () => {
      listeners.delete(listener);
    };
  }, []);

  return snapshot;
}
