import assert from 'node:assert/strict';
import test from 'node:test';
import config from '../vite.config';

test('Vite pre-optimizes the lazy ELK adapter before its first browser request', () => {
  assert.ok(config.optimizeDeps?.include?.includes('elkjs/lib/elk-api.js'));
});

test('Vite ignores generated diagnostics and browser artifacts across path formats', () => {
  const watch = config.server?.watch;
  assert.ok(watch);
  const ignored = watch.ignored;
  assert.ok(ignored instanceof RegExp);
  for (const path of [
    'C:\\repo\\.azure',
    'C:\\repo\\.azure\\provider-recovery\\original-provider-diagnostic.log',
    'C:\\repo\\DONOTTRACK\\app-quality\\browser-profile\\lockfile',
    '.azure\\provider-recovery\\diagnostic.log',
    'DONOTTRACK\\capture.png',
    '/repo/.azure/provider-recovery/diagnostic.log',
    '/repo/DONOTTRACK/browser-profile/lockfile',
  ]) {
    assert.equal(ignored.test(path), true, path);
  }
});

test('Vite keeps application, tests, and similarly named source paths watched', () => {
  const watch = config.server?.watch;
  assert.ok(watch);
  const ignored = watch.ignored;
  assert.ok(ignored instanceof RegExp);
  for (const path of [
    'C:\\repo\\src\\App.tsx',
    'C:\\repo\\src\\components\\WorkflowStepper.css',
    'C:\\repo\\tests\\e2e\\editor-performance.spec.ts',
    'C:\\repo\\src\\.azure-helper.ts',
    'C:\\repo\\src\\DONOTTRACKING\\index.ts',
    '/repo/src/utils/elkLayoutRuntime.ts',
    '/repo/tests/viteWatch.test.ts',
  ]) {
    assert.equal(ignored.test(path), false, path);
  }
});
