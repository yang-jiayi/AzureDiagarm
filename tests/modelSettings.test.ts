import test from 'node:test';
import assert from 'node:assert/strict';
import { build } from 'esbuild';
import type { FeatureType } from '../src/stores/modelSettingsStore';

const STORAGE_KEY = 'azure-diagrams-model-settings';
const env = {
  VITE_AZURE_OPENAI_ENDPOINT: 'https://test.openai.azure.com/',
  VITE_AZURE_OPENAI_DEPLOYMENT_GPT62ASTRA: 'actual-astra-alias',
  VITE_AZURE_OPENAI_DEPLOYMENT_GPT62SOL: 'actual-sol-alias',
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

test('both managed models resolve exact aliases and Responses format for every feature', async () => {
  const { store } = await loadStore();
  assert.deepEqual(store.getAvailableModels(), ['gpt-6.2-astra', 'gpt-6.2-sol']);
  assert.equal(store.getModelSettings().model, 'gpt-6.2-sol');
  for (const model of store.getAvailableModels()) {
    store.updateModelSettings({ model, reasoningEffort: 'max' });
    for (const feature of features) {
      const runtime = store.resolveAIModelRuntime(feature);
      assert.equal(runtime.model, model);
      assert.equal(runtime.deployment, model === 'gpt-6.2-astra' ? 'actual-astra-alias' : 'actual-sol-alias');
      assert.equal(runtime.apiFormat, 'responses');
      assert.equal(runtime.maxCompletionTokens, 32000);
      assert.equal(runtime.reasoningEffort, 'max');
      assert.equal(runtime.supportsVision, true);
    }
  }
});

for (const [retiredModel, expected] of [
  ['gpt-6-astra', 'gpt-6.2-astra'],
  ['gpt-6.1-sol', 'gpt-6.2-sol'],
  ['claude-opus-5-5', 'gpt-6.2-sol'],
]) {
  test(`saved ${retiredModel} preference migrates safely without touching diagram history`, async () => {
    const history = JSON.stringify({ model: retiredModel, diagram: { services: [] } });
    const entries = new Map([['azure-diagrams-history', history]]);
    const { store } = await loadStore(env, {
      version: 5, model: retiredModel, reasoningEffort: 'high',
      featureOverrides: { blueprint: { model: retiredModel, reasoningEffort: 'medium' } },
    }, entries);
    assert.equal(store.getModelSettings().model, expected);
    assert.equal(store.getModelSettingsForFeature('blueprint').model, expected);
    assert.equal(JSON.parse(entries.get(STORAGE_KEY)!).version, 6);
    assert.equal(entries.get('azure-diagrams-history'), history);
  });
}

test('global and feature selections persist independently', async () => {
  const { store, entries } = await loadStore();
  store.updateModelSettings({ model: 'gpt-6.2-astra', reasoningEffort: 'high' });
  store.updateFeatureOverride('validation', { model: 'gpt-6.2-sol', reasoningEffort: 'medium' });
  store.updateFeatureOverride('blueprint', { model: 'gpt-6.2-astra', reasoningEffort: 'max' });
  const snapshot = store.getModelSettings();
  snapshot.featureOverrides!.blueprint!.reasoningEffort = 'low';
  const reloaded = await loadStore(env, undefined, entries);
  assert.deepEqual(reloaded.store.getModelSettings(), store.getModelSettings());
  assert.equal(reloaded.store.getModelSettingsForFeature('blueprint').reasoningEffort, 'max');
  assert.equal(reloaded.store.getModelSettingsForFeature('validation').model, 'gpt-6.2-sol');
  assert.equal(reloaded.store.getModelSettingsForFeature('validation').reasoningEffort, 'medium');
});

test('unsupported saved reasoning normalizes to low globally and per feature without changing history', async () => {
  const history = JSON.stringify({ model: 'gpt-6.1-sol', reasoningEffort: 'none' });
  const entries = new Map([['azure-diagrams-history', history]]);
  const { store } = await loadStore(env, {
    version: 5, model: 'gpt-6.1-sol', reasoningEffort: 'none',
    featureOverrides: { blueprint: { model: 'gpt-6.2-astra', reasoningEffort: 'none' } },
  }, entries);
  assert.equal(store.getModelSettings().reasoningEffort, 'low');
  assert.equal(store.getModelSettingsForFeature('blueprint').reasoningEffort, 'low');
  assert.equal(store.resolveAIModelRuntime('blueprint').reasoningEffort, 'low');
  const saved = JSON.parse(entries.get(STORAGE_KEY)!);
  assert.equal(saved.reasoningEffort, 'low');
  assert.equal(saved.featureOverrides.blueprint.reasoningEffort, 'low');
  assert.equal(entries.get('azure-diagrams-history'), history);
});

test('missing selected model fails closed, never falling back to Sol', async () => {
  const { store } = await loadStore({ ...env, VITE_AZURE_OPENAI_DEPLOYMENT_GPT62ASTRA: '' });
  assert.deepEqual(store.getAvailableModels(), ['gpt-6.2-sol']);
  store.updateModelSettings({ model: 'gpt-6.2-astra' });
  assert.equal(store.isAnyAIModelConfigured(), false);
  assert.throws(() => store.resolveAIModelRuntime('blueprint'), { code: 'managed_model_not_configured' });
  assert.throws(() => store.getDeploymentName('gpt-6-astra' as never), { code: 'unsupported_ai_model' });
});

test('managed capture identity cannot be forged, changed or redirected after selection edits', async () => {
  const { store } = await loadStore();
  const capture = store.captureRuntimeModelOverride('validation');
  assert.equal(store.resolveAIModelRuntime('validation', { ...capture }).connection, capture.connection);
  assert.throws(() => store.resolveAIModelRuntime('validation', { ...capture, connection: { ...capture.connection! } }), { code: 'stale_ai_configuration' });
  assert.throws(() => store.resolveAIModelRuntime('validation', { ...capture, model: 'gpt-6.2-astra' }), { code: 'stale_ai_configuration' });
  store.updateModelSettings({ model: 'gpt-6.2-astra' });
  store.updateModelSettings({ model: 'gpt-6.2-sol' });
  assert.throws(() => store.resolveAIModelRuntime('validation', capture), { code: 'stale_ai_configuration' });
});

test('explicit overrides use a requested model and reject unsupported reasoning', async () => {
  const { store } = await loadStore();
  assert.equal(store.resolveAIModelRuntime('blueprint', { model: 'gpt-6.2-astra', reasoningEffort: 'high' }).apiFormat, 'responses');
  for (const model of ['gpt-6.2-astra', 'gpt-6.2-sol'] as const) {
    assert.throws(() => store.resolveAIModelRuntime('blueprint', { model, reasoningEffort: 'none' }), { code: 'unsupported_reasoning_effort' });
  }
  for (const model of ['gpt-6-astra', 'gpt-5.6-sol', '__proto__', 'constructor', undefined]) {
    assert.throws(() => Reflect.apply(store.updateModelSettings, undefined, [{ model }]), { code: 'unsupported_ai_model' });
  }
});
