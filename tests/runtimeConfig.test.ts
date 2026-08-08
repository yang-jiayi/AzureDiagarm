import test, { afterEach } from 'node:test';
import assert from 'node:assert/strict';
import {
  getRuntimeConfigSnapshot,
  isBYOAIEnabledOnServer,
  loadRuntimeConfig,
  resetRuntimeConfigForTests,
} from '../src/services/runtimeConfig.ts';

const originalFetch = globalThis.fetch;

afterEach(() => {
  globalThis.fetch = originalFetch;
  resetRuntimeConfigForTests();
});

test('runtime configuration exposes the BYO server kill switch', async () => {
  globalThis.fetch = (async () => new Response(JSON.stringify({
    features: { bringYourOwnAI: true },
  }), {
    status: 200,
    headers: { 'Content-Type': 'application/json' },
  })) as typeof fetch;

  const snapshot = await loadRuntimeConfig();
  assert.deepEqual(snapshot, { status: 'ready', bringYourOwnAI: true });
  assert.equal(isBYOAIEnabledOnServer(), true);
});

test('runtime configuration fails closed on malformed or unavailable responses', async () => {
  globalThis.fetch = (async () => new Response(JSON.stringify({
    features: {},
  }), {
    status: 200,
    headers: { 'Content-Type': 'application/json' },
  })) as typeof fetch;

  await loadRuntimeConfig();
  assert.equal(getRuntimeConfigSnapshot().status, 'error');
  assert.equal(isBYOAIEnabledOnServer(), false);
});
