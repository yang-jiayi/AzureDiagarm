import test from 'node:test';
import assert from 'node:assert/strict';
import { build } from 'esbuild';
import type { FeatureType, ModelType } from '../src/stores/modelSettingsStore';

const STORAGE_KEY = 'azure-diagrams-model-settings';
const legacyEnv = {
  VITE_AZURE_OPENAI_ENDPOINT: 'https://test.openai.azure.com/',
  VITE_AZURE_OPENAI_DEPLOYMENT_GPT56SOL: 'legacy-sol',
  VITE_AZURE_OPENAI_DEPLOYMENT_GPT56TERRA: 'legacy-terra',
  VITE_AZURE_OPENAI_DEPLOYMENT_GPT56LUNA: 'legacy-luna',
};
const astraEnv = {
  ...legacyEnv,
  VITE_AZURE_OPENAI_DEPLOYMENT_GPT6ASTRA: 'actual-astra-deployment',
};
const features: FeatureType[] = ['architectureGeneration', 'validation', 'deploymentGuide', 'blueprint'];

async function loadStore(env: Record<string, string>, stored?: object, entries = new Map<string, string>()) {
  if (stored) entries.set(STORAGE_KEY, JSON.stringify(stored));
  const result = await build({
    stdin: {
      contents: `
        export * from './src/stores/modelSettingsStore';
        export { resolveAIModelRuntime } from './src/services/aiModelRuntime';
        export { buildRequestBody } from './src/services/apiHelper';
      `,
      resolveDir: process.cwd(), loader: 'ts',
    },
    bundle: true, write: false, platform: 'node', format: 'cjs', logLevel: 'silent',
    define: { 'import.meta.env': JSON.stringify(env) },
  });
  const storage = {
    getItem: (key: string) => entries.get(key) ?? null,
    setItem: (key: string, value: string) => { entries.set(key, value); },
  };
  const module = { exports: {} };
  new Function('module', 'exports', 'localStorage', result.outputFiles[0].text)(
    module, module.exports, storage,
  );
  return {
    store: module.exports as typeof import('../src/stores/modelSettingsStore')
      & Pick<typeof import('../src/services/aiModelRuntime'), 'resolveAIModelRuntime'>
      & Pick<typeof import('../src/services/apiHelper'), 'buildRequestBody'>,
    entries,
  };
}

test('Astra is the default and recommendation for every feature, mapped to its own deployment', async () => {
  const { store } = await loadStore(astraEnv);
  assert.equal(store.getModelSettings().model, 'gpt-6-astra');
  assert.equal(store.getAvailableModels()[0], 'gpt-6-astra');
  assert.equal(store.getDeploymentName('gpt-6-astra'), 'actual-astra-deployment');
  assert.equal(store.MODEL_CONFIG['gpt-6-astra'].displayName, 'GPT-6 Astra');
  assert.equal(store.MODEL_CONFIG['gpt-6-astra'].apiFormat, 'responses');
  assert.deepEqual(store.getRecommendedModelSettings(), {
    model: 'gpt-6-astra', reasoningEffort: 'low', featureOverrides: {},
  });
  for (const feature of features) {
    assert.equal(store.FEATURE_CONFIG[feature].recommendedModel, 'gpt-6-astra');
    assert.deepEqual(store.getModelSettingsForFeature(feature), {
      model: 'gpt-6-astra', reasoningEffort: 'low',
    });
    const runtime = store.resolveAIModelRuntime(feature);
    assert.equal(runtime.source, 'managed');
    assert.equal(runtime.deployment, 'actual-astra-deployment');
    const request = store.buildRequestBody({
      ...runtime,
      messages: [{ role: 'user', content: 'Return a synthetic JSON architecture.' }],
      maxTokens: runtime.maxCompletionTokens,
    });
    assert.equal(request.model, 'actual-astra-deployment');
    assert.deepEqual(request.reasoning, { effort: 'low' });
    assert.deepEqual(request.text, { format: { type: 'json_object' } });
    assert.equal(request.store, false);
  }
});

for (const version of [1, 2]) {
  for (const model of ['gpt-5.6-sol', 'gpt-5.6-terra', 'gpt-5.6-luna'] as const) {
    test(`version ${version} ${model} selections and feature overrides migrate persistently`, async () => {
      const { store, entries } = await loadStore(astraEnv, {
        version, model, reasoningEffort: 'high',
        featureOverrides: {
          architectureGeneration: { model: 'gpt-5.6-sol', reasoningEffort: 'low' },
          validation: { model: 'gpt-5.6-terra', reasoningEffort: 'high' },
          deploymentGuide: { model: 'gpt-5.6-terra' },
          blueprint: { model: 'gpt-5.6-luna', reasoningEffort: 'medium' },
        },
      });
      assert.equal(store.getModelSettings().model, 'gpt-6-astra');
      for (const feature of features) {
        assert.equal(store.getModelSettingsForFeature(feature).model, 'gpt-6-astra');
        assert.equal(store.resolveAIModelRuntime(feature).deployment, 'actual-astra-deployment');
      }
      assert.equal(store.getModelSettingsForFeature('validation').reasoningEffort, 'high');
      assert.equal(store.getModelSettingsForFeature('deploymentGuide').reasoningEffort, 'high');
      assert.equal(store.getModelSettingsForFeature('blueprint').reasoningEffort, 'medium');
      const saved = JSON.parse(entries.get(STORAGE_KEY)!);
      assert.equal(saved.version, 3);
      assert.equal(saved.astraMigrationVersion, 1);
      assert.equal(saved.model, 'gpt-6-astra');
      const reloaded = await loadStore(astraEnv, undefined, entries);
      assert.deepEqual(reloaded.store.getModelSettings(), store.getModelSettings());
    });
  }
}

test('migration retains feature effort even when the old deployment is no longer configured', async () => {
  const { store } = await loadStore({
    VITE_AZURE_OPENAI_DEPLOYMENT_GPT6ASTRA: 'actual-astra-deployment',
  }, {
    version: 2, model: 'gpt-5.6-sol', reasoningEffort: 'high',
    featureOverrides: { validation: { model: 'gpt-5.6-terra', reasoningEffort: 'medium' } },
  });
  assert.deepEqual(store.getModelSettingsForFeature('validation'), {
    model: 'gpt-6-astra', reasoningEffort: 'medium',
  });
  assert.equal(store.getModelSettings().model, 'gpt-6-astra');
});

test('Astra is never aliased to an old model on an installation without Astra', async () => {
  const settings = {
    version: 2, model: 'gpt-5.6-sol', reasoningEffort: 'low',
    featureOverrides: { validation: { model: 'gpt-5.6-terra', reasoningEffort: 'medium' } },
  };
  const { store, entries } = await loadStore(legacyEnv, settings);
  assert.equal(store.getModelSettings().model, 'gpt-5.6-sol');
  assert.equal(store.getModelSettingsForFeature('validation').model, 'gpt-5.6-terra');
  assert.equal(store.isModelAvailable('gpt-6-astra'), false);
  assert.throws(() => store.getDeploymentName('gpt-6-astra'), /VITE_AZURE_OPENAI_DEPLOYMENT_GPT6ASTRA/);
  assert.equal(entries.get(STORAGE_KEY), JSON.stringify(settings));
  assert.deepEqual(store.getRecommendedModelSettings().featureOverrides, {
    validation: { model: 'gpt-5.6-terra', reasoningEffort: 'low' },
    deploymentGuide: { model: 'gpt-5.6-terra', reasoningEffort: 'low' },
    blueprint: { model: 'gpt-5.6-luna', reasoningEffort: 'low' },
  });
});

test('a later Astra deployment still migrates settings saved before Astra was configured', async () => {
  const old = await loadStore(legacyEnv);
  old.store.updateModelSettings({ model: 'gpt-5.6-luna', reasoningEffort: 'medium' });
  assert.equal(JSON.parse(old.entries.get(STORAGE_KEY)!).astraMigrationVersion, undefined);
  const current = await loadStore(astraEnv, undefined, old.entries);
  assert.equal(current.store.getModelSettings().model, 'gpt-6-astra');
  assert.equal(current.store.getModelSettings().reasoningEffort, 'medium');
  assert.equal(JSON.parse(current.entries.get(STORAGE_KEY)!).astraMigrationVersion, 1);
});

test('explicit alternative choices made after migration are not silently overwritten on reload', async () => {
  const initial = await loadStore(astraEnv, { version: 2, model: 'gpt-5.6-sol', reasoningEffort: 'low' });
  initial.store.updateModelSettings({ model: 'gpt-5.6-terra', reasoningEffort: 'high' });
  initial.store.updateFeatureOverride('blueprint', { model: 'gpt-5.6-luna', reasoningEffort: 'medium' });
  const reloaded = await loadStore(astraEnv, undefined, initial.entries);
  assert.equal(reloaded.store.getModelSettings().model, 'gpt-5.6-terra');
  assert.equal(reloaded.store.getModelSettingsForFeature('blueprint').model, 'gpt-5.6-luna');
});

test('older application code does not rewrite a newer persisted migration version', async () => {
  const settings = {
    version: 4, astraMigrationVersion: 2,
    model: 'gpt-5.6-sol', reasoningEffort: 'high', featureOverrides: {},
  };
  const { store, entries } = await loadStore(astraEnv, settings);
  assert.equal(store.getModelSettings().model, 'gpt-5.6-sol');
  assert.equal(entries.get(STORAGE_KEY), JSON.stringify(settings));
});

test('unrelated explicit models and separate BYO configuration are preserved', async () => {
  const byo = JSON.stringify({ enabled: true, model: 'personal-deployment', provider: 'azure-openai' });
  const entries = new Map([['azure-diagrams-byo-ai-settings', byo]]);
  const { store } = await loadStore({
    ...astraEnv,
    VITE_AZURE_OPENAI_DEPLOYMENT_GPT54MINI: 'mini',
    VITE_AZURE_OPENAI_DEPLOYMENT_DEEPSEEK: 'deepseek',
  }, {
    version: 2, model: 'gpt-5.4-mini', reasoningEffort: 'medium',
    featureOverrides: {
      validation: { model: 'gpt-5.6-terra', reasoningEffort: 'high' },
      blueprint: { model: 'deepseek-v3.2-speciale' },
    },
  }, entries);
  assert.equal(store.getModelSettings().model, 'gpt-5.4-mini');
  assert.equal(store.getModelSettingsForFeature('validation').model, 'gpt-6-astra');
  assert.equal(store.getModelSettingsForFeature('blueprint').model, 'deepseek-v3.2-speciale');
  assert.equal(entries.get('azure-diagrams-byo-ai-settings'), byo);
});

test('invalid persisted model and feature keys cannot become model overrides', async () => {
  const { store } = await loadStore(astraEnv, {
    version: 2, model: 'gpt-5.6-sol',
    featureOverrides: {
      unknownFeature: { model: 'gpt-5.6-sol' },
      validation: { model: 'constructor' },
      blueprint: { model: 'gpt-6-astra', reasoningEffort: 'unsupported' },
    },
  });
  assert.deepEqual(store.getModelSettings().featureOverrides, {
    blueprint: { model: 'gpt-6-astra' satisfies ModelType, reasoningEffort: 'low' },
  });
  const invalid = await loadStore(astraEnv, { model: 'constructor' });
  assert.equal(invalid.store.getModelSettings().model, 'gpt-6-astra');
});
