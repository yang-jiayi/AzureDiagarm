import test from 'node:test';
import assert from 'node:assert/strict';
import { build } from 'esbuild';
import type { FeatureType, ModelType } from '../src/stores/modelSettingsStore';

const STORAGE_KEY = 'azure-diagrams-model-settings';
const env = {
  VITE_AZURE_OPENAI_ENDPOINT: 'https://test.openai.azure.com/',
  VITE_AZURE_OPENAI_DEPLOYMENT_GPT61SOL: 'actual-sol-alias',
  VITE_AZURE_FOUNDRY_ENDPOINT: 'https://test.services.ai.azure.com/',
  VITE_AZURE_FOUNDRY_DEPLOYMENT_OPUS55: 'actual-opus-alias',
};
const features: FeatureType[] = ['architectureGeneration', 'validation', 'deploymentGuide', 'blueprint'];
const bundles = new Map<string, Promise<string>>();
async function loadStore(environment = env, stored?: object, entries = new Map<string, string>()) {
  if (stored) entries.set(STORAGE_KEY, JSON.stringify(stored));
  const key = JSON.stringify(environment);
  if (!bundles.has(key)) bundles.set(key, build({
    stdin: {
      contents: `export * from './src/stores/modelSettingsStore'; export * from './src/services/aiModelRuntime';`,
      resolveDir: process.cwd(), loader: 'ts',
    },
    bundle: true, write: false, platform: 'node', format: 'cjs', logLevel: 'silent',
    define: { 'import.meta.env': key },
  }).then(result => result.outputFiles[0].text));
  const module = { exports: {} };
  new Function('module', 'exports', 'localStorage', await bundles.get(key)!)(
    module, module.exports, {
      getItem: (key: string) => entries.get(key) ?? null,
      setItem: (key: string, value: string) => entries.set(key, value),
    },
  );
  return { store: module.exports as typeof import('../src/stores/modelSettingsStore')
    & typeof import('../src/services/aiModelRuntime'), entries };
}

test('both managed models resolve exact aliases and provider formats for every feature', async () => {
  const { store } = await loadStore();
  assert.deepEqual(store.getAvailableModels(), ['gpt-6.1-sol', 'claude-opus-5-5']);
  assert.equal(store.getModelSettings().model, 'gpt-6.1-sol');
  for (const model of store.getAvailableModels()) {
    store.updateModelSettings({ model, reasoningEffort: 'max' });
    for (const feature of features) {
      const runtime = store.resolveAIModelRuntime(feature);
      assert.equal(runtime.model, model);
      assert.equal(runtime.deployment, model === 'gpt-6.1-sol' ? 'actual-sol-alias' : 'actual-opus-alias');
      assert.equal(runtime.apiFormat, model === 'gpt-6.1-sol' ? 'responses' : 'anthropic-messages');
      assert.equal(runtime.maxCompletionTokens, 32000);
      assert.equal(runtime.reasoningEffort, 'max');
      assert.equal(runtime.supportsVision, true);
    }
  }
});

for (const version of [1, 2, 3, 4]) {
  test(`version ${version} retired preferences migrate to Sol without touching diagram history`, async () => {
    const history = JSON.stringify({ model: 'gpt-6-astra', diagram: { services: [] } });
    const entries = new Map([['azure-diagrams-history', history]]);
    const { store } = await loadStore(env, {
      version, model: 'gpt-6-astra', reasoningEffort: 'max',
      featureOverrides: { blueprint: { model: 'gpt-6-astra', reasoningEffort: 'high' } },
    }, entries);
    assert.equal(store.getModelSettings().model, 'gpt-6.1-sol');
    assert.equal(store.getModelSettings().reasoningEffort, 'max');
    assert.deepEqual(store.getModelSettingsForFeature('blueprint'), { model: 'gpt-6.1-sol', reasoningEffort: 'high' });
    assert.equal(JSON.parse(entries.get(STORAGE_KEY)!).version, 5);
    assert.equal(entries.get('azure-diagrams-history'), history);
  });
}

test('explicit Opus selection and feature overrides survive reload independently', async () => {
  const { store, entries } = await loadStore();
  store.updateModelSettings({ model: 'claude-opus-5-5', reasoningEffort: 'none' });
  assert.equal(store.getModelSettings().reasoningEffort, 'medium');
  store.updateFeatureOverride('validation', { model: 'gpt-6.1-sol', reasoningEffort: 'none' });
  store.updateFeatureOverride('blueprint', { model: 'claude-opus-5-5', reasoningEffort: 'max' });
  const snapshot = store.getModelSettings();
  snapshot.featureOverrides!.blueprint!.reasoningEffort = 'low';
  const reloaded = await loadStore(env, undefined, entries);
  assert.deepEqual(reloaded.store.getModelSettings(), store.getModelSettings());
  assert.equal(reloaded.store.getModelSettingsForFeature('blueprint').reasoningEffort, 'max');
  assert.equal(reloaded.store.getModelSettingsForFeature('validation').model, 'gpt-6.1-sol');
});

test('missing selected model fails closed, never falling back to Sol or retired Astra', async () => {
  const { store } = await loadStore({ ...env, VITE_AZURE_FOUNDRY_DEPLOYMENT_OPUS55: '' });
  assert.deepEqual(store.getAvailableModels(), ['gpt-6.1-sol']);
  store.updateModelSettings({ model: 'claude-opus-5-5' });
  assert.equal(store.isAnyAIModelConfigured(), false);
  assert.throws(() => store.resolveAIModelRuntime('blueprint'), { code: 'managed_model_not_configured' });
  assert.throws(() => store.getDeploymentName('gpt-6-astra' as ModelType), { code: 'unsupported_ai_model' });
});

test('managed capture identity cannot be forged, changed or redirected after selection edits', async () => {
  const { store } = await loadStore();
  const capture = store.captureRuntimeModelOverride('validation');
  assert.equal(store.resolveAIModelRuntime('validation', { ...capture }).connection, capture.connection);
  assert.throws(() => store.resolveAIModelRuntime('validation', { ...capture, connection: { ...capture.connection! } }), { code: 'stale_ai_configuration' });
  assert.throws(() => store.resolveAIModelRuntime('validation', { ...capture, model: 'claude-opus-5-5' }), { code: 'stale_ai_configuration' });
  store.updateModelSettings({ model: 'claude-opus-5-5' });
  store.updateModelSettings({ model: 'gpt-6.1-sol' });
  assert.throws(() => store.resolveAIModelRuntime('validation', capture), { code: 'stale_ai_configuration' });
});

test('explicit overrides use the requested model, with model-specific effort validation', async () => {
  const { store } = await loadStore();
  assert.equal(store.resolveAIModelRuntime('blueprint', { model: 'claude-opus-5-5', reasoningEffort: 'high' }).apiFormat, 'anthropic-messages');
  assert.throws(() => store.resolveAIModelRuntime('blueprint', { model: 'claude-opus-5-5', reasoningEffort: 'none' }), { code: 'unsupported_reasoning_effort' });
  for (const model of ['gpt-6-astra', 'gpt-5.6-sol', '__proto__', 'constructor', undefined]) {
    assert.throws(() => Reflect.apply(store.updateModelSettings, undefined, [{ model }]), { code: 'unsupported_ai_model' });
  }
});
