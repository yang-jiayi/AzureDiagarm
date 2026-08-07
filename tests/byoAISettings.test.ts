import test, { afterEach, beforeEach } from 'node:test';
import assert from 'node:assert/strict';
import {
  DEFAULT_BYO_AI_SETTINGS,
  disableBYOAI,
  getBYOAIApiKey,
  getBYOAISettings,
  isBYOAIReady,
  saveBYOAIConfiguration,
  validateBYOAISettings,
} from '../src/stores/byoAISettingsStore.ts';

class MemoryStorage implements Storage {
  private readonly values = new Map<string, string>();

  get length(): number {
    return this.values.size;
  }

  clear(): void {
    this.values.clear();
  }

  getItem(key: string): string | null {
    return this.values.get(key) ?? null;
  }

  key(index: number): string | null {
    return [...this.values.keys()][index] ?? null;
  }

  removeItem(key: string): void {
    this.values.delete(key);
  }

  setItem(key: string, value: string): void {
    this.values.set(key, value);
  }
}

const storage = new MemoryStorage();
const originalLocalStorage = globalThis.localStorage;

beforeEach(() => {
  Object.defineProperty(globalThis, 'localStorage', {
    configurable: true,
    value: storage,
  });
  storage.clear();
  saveBYOAIConfiguration(DEFAULT_BYO_AI_SETTINGS, '');
});

afterEach(() => {
  disableBYOAI();
  Object.defineProperty(globalThis, 'localStorage', {
    configurable: true,
    value: originalLocalStorage,
  });
});

test('BYO settings persist public configuration but never persist the API key', () => {
  const apiKey = 'sk-test-secret-value-123456';
  saveBYOAIConfiguration({
    enabled: true,
    provider: 'azure-openai',
    endpoint: 'https://contoso.openai.azure.com/',
    model: 'architecture-deployment',
    apiFormat: 'responses',
    apiVersion: '2024-12-01-preview',
    isReasoning: true,
    supportsVision: true,
  }, apiKey);

  const persisted = storage.getItem('azure-diagrams-byo-ai-settings');
  assert.ok(persisted);
  assert.equal(persisted.includes(apiKey), false);
  assert.equal(persisted.includes('apiKey'), false);
  assert.deepEqual(getBYOAISettings(), {
    enabled: true,
    provider: 'azure-openai',
    endpoint: 'https://contoso.openai.azure.com/',
    model: 'architecture-deployment',
    apiFormat: 'responses',
    apiVersion: '2024-12-01-preview',
    isReasoning: true,
    supportsVision: true,
  });
  assert.equal(getBYOAIApiKey(), apiKey);
  assert.equal(isBYOAIReady(), true);

  disableBYOAI();
  assert.equal(getBYOAIApiKey(), '');
  assert.equal(getBYOAISettings().enabled, false);
  assert.equal(isBYOAIReady(), false);
});

test('BYO validation restricts endpoints to Azure OpenAI or official OpenAI', () => {
  const azure = validateBYOAISettings({
    enabled: true,
    provider: 'azure-openai',
    endpoint: 'https://contoso.openai.azure.com',
    model: 'architecture-deployment',
    apiFormat: 'responses',
    apiVersion: '2024-12-01-preview',
    isReasoning: true,
    supportsVision: true,
  }, 'azure-secret-key-value');
  assert.equal(azure.valid, true);

  const officialOpenAI = validateBYOAISettings({
    enabled: true,
    provider: 'openai',
    endpoint: 'https://api.openai.com',
    model: 'gpt-5',
    apiFormat: 'responses',
    apiVersion: '',
    isReasoning: true,
    supportsVision: true,
  }, 'sk-openai-secret-value');
  assert.equal(officialOpenAI.valid, true);

  const untrusted = validateBYOAISettings({
    enabled: true,
    provider: 'azure-openai',
    endpoint: 'https://example.com/openai',
    model: 'gpt-5',
    apiFormat: 'responses',
    apiVersion: '2024-12-01-preview',
    isReasoning: true,
    supportsVision: true,
  }, 'untrusted-secret-value');
  assert.equal(untrusted.valid, false);
  assert.match(untrusted.errors.endpoint || '', /trusted Azure OpenAI HTTPS endpoint/);
});
