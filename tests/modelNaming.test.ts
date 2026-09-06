import test, { afterEach } from 'node:test';
import assert from 'node:assert/strict';
import {
  clearSourceModel,
  generateModelFilename,
  getModelAbbreviation,
  getModelSuffix,
  setSourceModel,
} from '../src/utils/modelNaming';

afterEach(clearSourceModel);

test('Astra exports retain actual model provenance instead of an unknown model suffix', () => {
  setSourceModel('gpt-6-astra', 'high');
  assert.equal(getModelAbbreviation(), 'gpt6astra');
  assert.equal(getModelSuffix(), 'gpt6astra-high');
  assert.equal(
    generateModelFilename('architecture', 'pptx', 123456),
    'architecture-123456-gpt6astra-high.pptx',
  );
  assert.equal(
    generateModelFilename('architecture', 'vsdx', 123456),
    'architecture-123456-gpt6astra-high.vsdx',
  );
});

test('explicit legacy comparison exports keep their own source model', () => {
  setSourceModel('gpt-5.6-terra', 'low');
  assert.equal(getModelSuffix(), 'gpt56terra-low');
  setSourceModel('deepseek-v3.2-speciale', 'low');
  assert.equal(getModelSuffix(), 'deepseek');
});
