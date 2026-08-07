// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

import { useEffect, useState } from 'react';

export type BYOAIProvider = 'azure-openai' | 'openai';
export type BYOAIAPIFormat = 'responses' | 'chat-completions';

export interface BYOAISettings {
  enabled: boolean;
  provider: BYOAIProvider;
  endpoint: string;
  model: string;
  apiFormat: BYOAIAPIFormat;
  apiVersion: string;
  isReasoning: boolean;
  supportsVision: boolean;
}

export interface BYOAIValidationResult {
  valid: boolean;
  settings?: BYOAISettings;
  errors: Partial<Record<'endpoint' | 'model' | 'apiVersion' | 'apiKey', string>>;
}

const STORAGE_KEY = 'azure-diagrams-byo-ai-settings';
const STORAGE_VERSION = 1;
const DEFAULT_AZURE_API_VERSION = '2024-05-01-preview';
const OFFICIAL_OPENAI_ENDPOINT = 'https://api.openai.com';
const MODEL_NAME_RE = /^[A-Za-z0-9._:-]{1,128}$/;
const API_VERSION_RE = /^\d{4}-\d{2}-\d{2}(?:-preview)?$/;
const API_KEY_RE = /^[^\s\r\n]{8,512}$/;
const AZURE_OPENAI_HOST_SUFFIXES = [
  '.openai.azure.com',
  '.openai.azure.us',
  '.openai.azure.cn',
  '.cognitiveservices.azure.com',
  '.cognitiveservices.azure.us',
  '.cognitiveservices.azure.cn',
];

export const DEFAULT_BYO_AI_SETTINGS: BYOAISettings = {
  enabled: false,
  provider: 'azure-openai',
  endpoint: '',
  model: 'gpt-5.6-sol',
  apiFormat: 'responses',
  apiVersion: DEFAULT_AZURE_API_VERSION,
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

export function normalizeBYOAISettings(value: unknown): BYOAISettings {
  const raw = value && typeof value === 'object'
    ? value as Partial<BYOAISettings>
    : {};
  const provider = normalizeProvider(raw.provider);
  return {
    enabled: raw.enabled === true,
    provider,
    endpoint: provider === 'openai'
      ? OFFICIAL_OPENAI_ENDPOINT
      : String(raw.endpoint ?? '').trim(),
    model: String(raw.model ?? DEFAULT_BYO_AI_SETTINGS.model).trim(),
    apiFormat: normalizeApiFormat(raw.apiFormat),
    apiVersion: String(raw.apiVersion ?? DEFAULT_AZURE_API_VERSION).trim(),
    isReasoning: raw.isReasoning !== false,
    supportsVision: raw.supportsVision !== false,
  };
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
      errors.endpoint = 'Enter a trusted Azure OpenAI HTTPS endpoint without an API path.';
    } else {
      settings.endpoint = endpoint;
    }
    if (!API_VERSION_RE.test(settings.apiVersion)) {
      errors.apiVersion = 'Use an Azure API version such as 2024-05-01-preview.';
    }
  } else {
    settings.endpoint = OFFICIAL_OPENAI_ENDPOINT;
  }

  if (!MODEL_NAME_RE.test(settings.model)) {
    errors.model = 'Use a valid deployment or model name (letters, numbers, ., _, -, or :).';
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

export function saveBYOAIConfiguration(settings: BYOAISettings, apiKey?: string): void {
  currentSettings = normalizeBYOAISettings(settings);
  if (apiKey !== undefined) currentApiKey = apiKey.trim();
  saveSettings(currentSettings);
  notifyListeners();
}

export function disableBYOAI(): void {
  currentSettings = { ...currentSettings, enabled: false };
  currentApiKey = '';
  saveSettings(currentSettings);
  notifyListeners();
}

export function isBYOAIReady(): boolean {
  if (!currentSettings.enabled) return false;
  return validateBYOAISettings(currentSettings, currentApiKey).valid;
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
} {
  const [snapshot, setSnapshot] = useState(() => ({
    settings: getBYOAISettings(),
    hasApiKey: hasBYOAIApiKey(),
  }));

  useEffect(() => {
    const listener = () => {
      setSnapshot({
        settings: getBYOAISettings(),
        hasApiKey: hasBYOAIApiKey(),
      });
    };
    listeners.add(listener);
    return () => {
      listeners.delete(listener);
    };
  }, []);

  return snapshot;
}
