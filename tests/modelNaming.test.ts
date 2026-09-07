import test, { afterEach } from 'node:test';
import assert from 'node:assert/strict';
import * as modelNaming from '../src/utils/modelNaming';
import {
  getModelSettings,
  getSupportedReasoningEfforts,
  updateFeatureOverride,
  updateModelSettings,
} from '../src/stores/modelSettingsStore';

const initialSettings = getModelSettings();
afterEach(() => updateModelSettings(initialSettings));

test('retired comparison naming state and helpers are no longer exported', () => {
  for (const name of [
    'setSourceModel', 'clearSourceModel', 'getHistoricalModelDisplayName',
    'getModelAbbreviation', 'getReasoningEffort', 'isReasoningModel',
  ]) {
    assert.equal(name in modelNaming, false, name);
  }
});

test('explicit artifact provenance names Astra exports at their captured reasoning effort', () => {
  for (const effort of getSupportedReasoningEfforts('gpt-6-astra')) {
    updateModelSettings({ reasoningEffort: effort === 'low' ? 'max' : 'low' });
    const provenance = Object.freeze({ source: 'managed' as const, model: 'GPT-6 Astra', reasoningEffort: effort });
    assert.equal(modelNaming.getModelSuffix(provenance), `gpt6astra-${effort}`);
    for (const extension of ['pptx', 'vsdx', 'json', 'md', 'csv']) {
      assert.equal(
        modelNaming.generateModelFilename('architecture', extension, 123456, provenance),
        `architecture-123456-gpt6astra-${effort}.${extension}`,
      );
    }
  }
});

test('filenames without artifact provenance are neutral regardless of live model preferences', () => {
  updateModelSettings({ reasoningEffort: 'max' });
  updateFeatureOverride('validation', { model: 'gpt-6-astra', reasoningEffort: 'none' });
  assert.equal(modelNaming.getModelSuffix(), '');
  assert.equal(modelNaming.generateModelFilename('diagram', 'svg', 123), 'diagram-123.svg');
  updateModelSettings({ reasoningEffort: 'none' });
  assert.equal(modelNaming.getModelSuffix(), '');
  assert.equal(modelNaming.generateModelFilename('historical', 'json', 123), 'historical-123.json');
});

test('filenames use the current time only when no timestamp is supplied', t => {
  t.mock.method(Date, 'now', () => 987654);
  updateModelSettings({ reasoningEffort: 'high' });
  assert.equal(modelNaming.generateModelFilename('diagram', 'svg'), 'diagram-987654.svg');
  assert.equal(modelNaming.generateModelFilename('diagram', 'svg', 123), 'diagram-123.svg');
  assert.equal(modelNaming.generateModelFilename('diagram', 'svg', 0), 'diagram-0.svg');
});

test('BYO artifact filenames use captured source, actual deployment and reasoning rather than a display label or live selection', () => {
  const provenance = Object.freeze({
    source: 'bring-your-own' as const, model: 'BYO Azure OpenAI · Friendly label · production',
    deployment: 'customer-production-alias', reasoningEffort: 'max',
  });
  updateModelSettings({ model: 'gpt-6-astra', reasoningEffort: 'low' });
  assert.equal(modelNaming.getModelSuffix(provenance), 'byo-customer-production-alias-max');
  assert.equal(modelNaming.generateModelFilename('guide', 'md', 123, provenance), 'guide-123-byo-customer-production-alias-max.md');
  assert.deepEqual(provenance, {
    source: 'bring-your-own', model: 'BYO Azure OpenAI · Friendly label · production',
    deployment: 'customer-production-alias', reasoningEffort: 'max',
  });
  assert.equal(modelNaming.getModelSuffix({ model: 'gpt-5.6-terra', reasoningEffort: 'high' }), 'gpt-5-6-terra-high');
});
