import assert from 'node:assert/strict';
import { existsSync } from 'node:fs';
import test from 'node:test';
import config from '../vite.config.ts';

test('development dependency discovery includes only the application HTML entries', () => {
  const entries = config.optimizeDeps?.entries;
  assert.deepEqual(entries, ['index.html', 'blueprint-preview.html']);
  assert.ok(Array.isArray(entries));
  for (const entry of entries) {
    assert.ok(existsSync(new URL(`../${entry}`, import.meta.url)), `Missing application entry: ${entry}`);
  }
});

test('scoped discovery retains normal optimization and the lazy ELK adapter', () => {
  assert.ok(config.optimizeDeps?.include?.includes('elkjs/lib/elk-api.js'));
  assert.notEqual(config.optimizeDeps?.noDiscovery, true);
});

test('pricing refresh recovery files do not enter the development file watcher', () => {
  const ignored = config.server?.watch?.ignored;
  assert.ok(ignored instanceof RegExp);
  assert.ok(ignored.test('C:\\project\\.pricing-refresh\\regions\\prices.json'));
  assert.ok(!ignored.test('C:\\project\\src\\data\\azurePricing.ts'));
});
