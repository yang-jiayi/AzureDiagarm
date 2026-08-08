import test, { afterEach, beforeEach } from 'node:test';
import assert from 'node:assert/strict';
import {
  DEFAULT_BYO_AI_SETTINGS,
  disableBYOAI,
  getBYOAIApiKey,
  getBYOAIConnectionState,
  getBYOAISettings,
  inferBYOAICapabilities,
  isBYOAIReady,
  isBYOAIVerified,
  reloadBYOAISettings,
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

test('BYO settings persist public configuration and require key re-entry after reload', () => {
  const apiKey = 'sk-test-secret-value-123456';
  saveBYOAIConfiguration({
    enabled: true,
    provider: 'azure-openai',
    endpoint: 'https://contoso.openai.azure.com/',
    model: 'gpt-5-architecture',
    apiFormat: 'responses',
    reasoningEffort: 'high',
    capabilityMode: 'manual',
    isReasoning: true,
    supportsVision: true,
  }, apiKey, { verified: true });

  const persisted = storage.getItem('azure-diagrams-byo-ai-settings');
  assert.ok(persisted);
  assert.equal(persisted.includes(apiKey), false);
  assert.equal(persisted.includes('apiKey'), false);
  assert.equal(getBYOAIApiKey(), apiKey);
  assert.equal(isBYOAIReady(), true);
  assert.equal(isBYOAIVerified(), true);
  assert.equal(getBYOAIConnectionState(), 'verified');

  reloadBYOAISettings();

  assert.deepEqual(getBYOAISettings(), {
    enabled: true,
    provider: 'azure-openai',
    endpoint: 'https://contoso.openai.azure.com/',
    model: 'gpt-5-architecture',
    apiFormat: 'responses',
    reasoningEffort: 'high',
    capabilityMode: 'manual',
    isReasoning: true,
    supportsVision: true,
  });
  assert.equal(getBYOAIApiKey(), '');
  assert.equal(isBYOAIReady(), false);
  assert.equal(isBYOAIVerified(), false);
  assert.equal(getBYOAIConnectionState(), 'key-required');

  disableBYOAI();
  assert.equal(getBYOAISettings().enabled, false);
  assert.equal(getBYOAIConnectionState(), 'disabled');
});

test('BYO validation accepts official Azure and Foundry origins and rejects unsafe names', () => {
  const azure = validateBYOAISettings({
    enabled: true,
    provider: 'azure-openai',
    endpoint: 'https://contoso.openai.azure.com',
    model: 'architecture-deployment',
    apiFormat: 'responses',
    reasoningEffort: 'low',
    capabilityMode: 'manual',
    isReasoning: true,
    supportsVision: true,
  }, 'azure-secret-key-value');
  assert.equal(azure.valid, true);

  const foundry = validateBYOAISettings({
    enabled: true,
    provider: 'azure-openai',
    endpoint: 'https://contoso.services.ai.azure.com',
    model: 'gpt-5-production',
    apiFormat: 'chat-completions',
    reasoningEffort: 'minimal',
    capabilityMode: 'auto',
    isReasoning: false,
    supportsVision: false,
  }, 'foundry-secret-key-value');
  assert.equal(foundry.valid, true);
  assert.equal(foundry.settings?.isReasoning, true);

  const officialOpenAI = validateBYOAISettings({
    enabled: true,
    provider: 'openai',
    endpoint: 'https://untrusted.example',
    model: 'gpt-5',
    apiFormat: 'responses',
    reasoningEffort: 'low',
    capabilityMode: 'auto',
    isReasoning: false,
    supportsVision: false,
  }, 'sk-openai-secret-value');
  assert.equal(officialOpenAI.valid, true);
  assert.equal(officialOpenAI.settings?.endpoint, 'https://api.openai.com');

  const untrusted = validateBYOAISettings({
    enabled: true,
    provider: 'azure-openai',
    endpoint: 'https://example.com/openai',
    model: '..',
    apiFormat: 'responses',
    reasoningEffort: 'low',
    capabilityMode: 'manual',
    isReasoning: true,
    supportsVision: true,
  }, 'untrusted-secret-value');
  assert.equal(untrusted.valid, false);
  assert.match(untrusted.errors.endpoint || '', /trusted Azure OpenAI or Microsoft Foundry/);
  assert.match(untrusted.errors.model || '', /at least one letter or number/);
});

test('automatic BYO capability detection is conservative for custom deployment aliases', () => {
  assert.deepEqual(inferBYOAICapabilities('gpt-5.6-sol'), {
    isReasoning: true,
    supportsVision: true,
  });
  assert.deepEqual(inferBYOAICapabilities('gpt-4o'), {
    isReasoning: false,
    supportsVision: true,
  });
  assert.deepEqual(inferBYOAICapabilities('architecture-production'), {
    isReasoning: false,
    supportsVision: false,
  });
});
