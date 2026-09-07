import test from 'node:test';
import assert from 'node:assert/strict';
import { build } from 'esbuild';
import type { FeatureType } from '../src/stores/modelSettingsStore';

const STORAGE_KEY = 'azure-diagrams-model-settings';
const legacyEnv = {
  VITE_AZURE_OPENAI_ENDPOINT: 'https://test.openai.azure.com/',
  VITE_AZURE_OPENAI_DEPLOYMENT_GPT56SOL: 'legacy-sol',
  VITE_AZURE_OPENAI_DEPLOYMENT_GPT56TERRA: 'legacy-terra',
  VITE_AZURE_OPENAI_DEPLOYMENT_GPT56LUNA: 'legacy-luna',
  VITE_AZURE_FOUNDRY_ENDPOINT: 'https://test.services.ai.azure.com/',
  VITE_AZURE_FOUNDRY_DEPLOYMENT_CLAUDE_OPUS5: 'legacy-claude',
};
const astraEnv = { ...legacyEnv, VITE_AZURE_OPENAI_DEPLOYMENT_GPT6ASTRA: 'actual-astra-deployment' };
const features: FeatureType[] = ['architectureGeneration', 'validation', 'deploymentGuide', 'blueprint'];
const legacyModels = [
  'gpt-5.1', 'gpt-5.2', 'gpt-5.4', 'gpt-5.4-mini',
  'gpt-5.6-sol', 'gpt-5.6-terra', 'gpt-5.6-luna', 'claude-opus-5',
  'deepseek-v3.2-speciale', 'deepseek-v4-pro', 'grok-4.1-fast', 'grok-4.3',
  'mistral-large-3', 'kimi-k2-5', 'kimi-k2-7-code', 'custom-deployment',
];
const bundles = new Map<string, Promise<string>>();

async function loadStore(env: Record<string, string>, stored?: object, entries = new Map<string, string>()) {
  if (stored) entries.set(STORAGE_KEY, JSON.stringify(stored));
  const key = JSON.stringify(env);
  if (!bundles.has(key)) bundles.set(key, build({
    stdin: {
      contents: `
        export * from './src/stores/modelSettingsStore';
        export * from './src/services/aiModelRuntime';
        export { buildRequestBody, callAzureOpenAIProxy } from './src/services/apiHelper';
      `,
      resolveDir: process.cwd(), loader: 'ts',
    },
    bundle: true, write: false, platform: 'node', format: 'cjs', logLevel: 'silent',
    define: { 'import.meta.env': JSON.stringify(env) },
  }).then(result => result.outputFiles[0].text));
  const bundled = await bundles.get(key)!;
  const storage = {
    getItem: (key: string) => entries.get(key) ?? null,
    setItem: (key: string, value: string) => { entries.set(key, value); },
  };
  const module = { exports: {} };
  new Function('module', 'exports', 'localStorage', bundled)(
    module, module.exports, storage,
  );
  return {
    store: module.exports as typeof import('../src/stores/modelSettingsStore')
      & typeof import('../src/services/aiModelRuntime')
      & Pick<typeof import('../src/services/apiHelper'), 'buildRequestBody' | 'callAzureOpenAIProxy'>,
    entries,
  };
}

test('singleton Astra registry and actual deployment-alias dispatch preserve MAX/32K', async t => {
  const envelopes: unknown[] = [];
  t.mock.method(globalThis, 'fetch', async (url: unknown, options: RequestInit) => {
    assert.equal(url, '/api/openai');
    envelopes.push(JSON.parse(String(options.body)));
    return new Response(JSON.stringify({ output_text: '{"services":[]}' }),
      { headers: { 'Content-Type': 'application/json' } });
  });
  const { store } = await loadStore(astraEnv);
  assert.deepEqual(Object.keys(store.MODEL_CONFIG), ['gpt-6-astra']);
  assert.deepEqual(store.getAvailableModels(), ['gpt-6-astra']);
  assert.deepEqual(store.getDeploymentNames(), { 'gpt-6-astra': 'actual-astra-deployment' });
  store.updateModelSettings({ model: 'gpt-6-astra', reasoningEffort: 'max' });
  for (const feature of features) {
    assert.equal(store.FEATURE_CONFIG[feature].recommendedModel, 'gpt-6-astra');
    const runtime = store.resolveAIModelRuntime(feature);
    assert.equal(runtime.source, 'managed');
    assert.equal(runtime.deployment, 'actual-astra-deployment');
    assert.equal(runtime.apiFormat, 'responses');
    assert.equal(runtime.reasoningEffort, 'max');
    assert.equal(runtime.maxCompletionTokens, 32000);
    assert.equal('byo' in runtime, false);
    const request = store.buildRequestBody({
      ...runtime, messages: [{ role: 'user', content: 'Offline request' }], maxTokens: runtime.maxCompletionTokens,
    });
    assert.equal(request.model, 'actual-astra-deployment');
    assert.equal(request.max_output_tokens, 32000);
    assert.deepEqual(request.reasoning, { effort: 'max' });
    assert.equal(request.store, false);
    const result = await store.callAzureOpenAIProxy({
      apiFormat: runtime.apiFormat, deployment: runtime.deployment, body: request,
    });
    assert.equal(result.ok, true);
    assert.deepEqual(envelopes.at(-1), {
      apiFormat: 'responses', deployment: 'actual-astra-deployment', body: request,
    });
  }
  assert.equal(envelopes.length, features.length);
});

for (const [name, env] of [
  ['only legacy configuration', legacyEnv],
  ['only endpoint', { VITE_AZURE_OPENAI_ENDPOINT: 'https://test.openai.azure.com/' }],
  ['only Astra deployment', { VITE_AZURE_OPENAI_DEPLOYMENT_GPT6ASTRA: 'astra' }],
  ['blank Astra deployment', { ...astraEnv, VITE_AZURE_OPENAI_DEPLOYMENT_GPT6ASTRA: ' ' }],
] as const) {
  test(`missing-Astra fails closed: ${name}`, async t => {
    const fetch = t.mock.method(globalThis, 'fetch', async () => { throw new Error('Unexpected HTTP'); });
    const { store } = await loadStore(env);
    assert.deepEqual(store.getAvailableModels(), []);
    assert.equal(store.isAnyAIModelConfigured(), false);
    assert.equal(store.isManagedAIModelConfigured(), false);
    assert.equal(store.getModelSettings().model, 'gpt-6-astra');
    for (const feature of features) assert.throws(() => store.resolveAIModelRuntime(feature));
    await assert.rejects(store.callAzureOpenAIProxy({
      apiFormat: 'responses', deployment: 'gpt-6-astra', body: { model: 'gpt-6-astra' },
    }), { name: 'AIModelConfigurationError', code: 'astra_not_configured' });
    assert.equal(fetch.mock.callCount(), 0);
  });
}

for (const version of [1, 2, 3, 4]) {
  test(`all legacy version ${version} selections migrate even after previous Astra migration`, async () => {
    // Bundle once; each execution creates a fresh store with this persisted fixture.
    for (const model of legacyModels) {
      const { store, entries } = await loadStore(astraEnv, {
        version, astraMigrationVersion: 1, model, reasoningEffort: 'max',
        featureOverrides: {
          architectureGeneration: { model, reasoningEffort: 'xhigh' },
          validation: { model, reasoningEffort: 'high' },
          deploymentGuide: { model },
          blueprint: { model, reasoningEffort: 'none' },
        },
      });
      assert.equal(store.getModelSettings().model, 'gpt-6-astra');
      assert.equal(store.getModelSettings().reasoningEffort, 'max');
      for (const feature of features) assert.equal(store.getModelSettingsForFeature(feature).model, 'gpt-6-astra');
      assert.equal(store.getModelSettingsForFeature('architectureGeneration').reasoningEffort, 'xhigh');
      assert.equal(store.getModelSettingsForFeature('validation').reasoningEffort, 'high');
      assert.equal(store.getModelSettingsForFeature('deploymentGuide').reasoningEffort, 'max');
      assert.equal(store.getModelSettingsForFeature('blueprint').reasoningEffort, 'none');
      const saved = JSON.parse(entries.get(STORAGE_KEY)!);
      assert.equal(saved.version, 4);
      assert.equal(saved.astraOnlyVersion, 1);
      assert.equal(saved.model, 'gpt-6-astra');
      assert.ok(Object.values(saved.featureOverrides).every((value) => (
        typeof value === 'object' && value !== null && 'model' in value && value.model === 'gpt-6-astra'
      )));
    }
  });
}

test('migration does not depend on Astra availability and never touches historical records', async () => {
  const history = JSON.stringify({ model: 'gpt-5.6-terra', reasoningEffort: 'max', diagram: { services: [] } });
  const entries = new Map([['azure-diagrams-history', history], ['review-history', history]]);
  const { store } = await loadStore(legacyEnv, {
    version: 3, astraMigrationVersion: 1, model: 'gpt-5.6-terra', reasoningEffort: 'max',
  }, entries);
  assert.equal(store.getModelSettings().model, 'gpt-6-astra');
  assert.equal(store.getModelSettings().reasoningEffort, 'max');
  assert.equal(JSON.parse(entries.get(STORAGE_KEY)!).model, 'gpt-6-astra');
  assert.equal(entries.get('azure-diagrams-history'), history);
  assert.equal(entries.get('review-history'), history);
});

test('current Astra preferences and feature effort survive reload and detached settings snapshots', async () => {
  const initial = await loadStore(astraEnv);
  initial.store.updateModelSettings({ model: 'gpt-6-astra', reasoningEffort: 'max' });
  initial.store.updateFeatureOverride('blueprint', { model: 'gpt-6-astra', reasoningEffort: 'high' });
  const snapshot = initial.store.getModelSettings();
  snapshot.featureOverrides!.blueprint!.reasoningEffort = 'none';
  assert.equal(initial.store.getModelSettingsForFeature('blueprint').reasoningEffort, 'high');
  const reloaded = await loadStore(astraEnv, undefined, initial.entries);
  assert.deepEqual(reloaded.store.getModelSettings(), initial.store.getModelSettings());
});

test('explicit legacy and forged overrides are rejected, not migrated at runtime', async () => {
  const { store } = await loadStore(astraEnv);
  for (const model of [...legacyModels, 'constructor', '__proto__', '', undefined]) {
    for (const feature of features) {
      assert.throws(() => Reflect.apply(store.resolveAIModelRuntime, undefined, [
        feature, { model, reasoningEffort: 'max', forceManaged: true },
      ]), { name: 'AIModelConfigurationError', code: 'unsupported_ai_model' });
    }
    assert.throws(() => Reflect.apply(store.getDeploymentName, undefined, [model]));
    assert.throws(() => Reflect.apply(store.updateModelSettings, undefined, [{ model }]));
  }
  assert.throws(() => Reflect.apply(store.updateFeatureOverride, undefined, [
    'validation', { model: 'gpt-5.6-terra', reasoningEffort: 'max' },
  ]));
  assert.equal(store.getModelSettings().model, 'gpt-6-astra');
});

test('unsupported runtime effort fails explicitly; invalid saved effort uses Astra default', async () => {
  const { store } = await loadStore(astraEnv, {
    version: 3, model: 'gpt-5.6-sol', reasoningEffort: 'minimal',
    featureOverrides: { blueprint: { model: 'gpt-5.6-luna', reasoningEffort: 'invalid' }, unknown: { model: 'gpt-6-astra' } },
  });
  assert.equal(store.getModelSettings().reasoningEffort, 'low');
  assert.deepEqual(store.getModelSettings().featureOverrides, { blueprint: { model: 'gpt-6-astra', reasoningEffort: 'low' } });
  assert.throws(() => Reflect.apply(store.resolveAIModelRuntime, undefined, [
    'architectureGeneration', { model: 'gpt-6-astra', reasoningEffort: 'minimal' },
  ]), { name: 'AIModelConfigurationError', code: 'unsupported_reasoning_effort' });
});
