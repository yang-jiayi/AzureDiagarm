import test, { afterEach } from 'node:test';
import assert from 'node:assert/strict';
import {
  getRuntimeConfigSnapshot, isBYOAIEnabledOnServer, loadRuntimeConfig, resetRuntimeConfigForTests,
} from '../src/services/runtimeConfig';
import { capabilityResponse, deferred, jsonResponse } from './fixtures/byoClientHarness';

afterEach(() => resetRuntimeConfigForTests());

test('capability reads are same-origin, no-store, detached and explicitly refreshable', async t => {
  const requests: RequestInit[] = [];
  const fetch = t.mock.method(globalThis, 'fetch', async (url: unknown, options?: RequestInit) => {
    assert.equal(url, '/api/runtime-config');
    requests.push(options!);
    return capabilityResponse();
  });
  const snapshot = await loadRuntimeConfig();
  assert.deepEqual(snapshot, { status: 'ready', bringYourOwnAI: true });
  assert.equal(isBYOAIEnabledOnServer(), true);
  snapshot.bringYourOwnAI = false;
  assert.equal(getRuntimeConfigSnapshot().bringYourOwnAI, true);
  await loadRuntimeConfig();
  assert.equal(fetch.mock.callCount(), 1);
  await loadRuntimeConfig(true);
  assert.equal(fetch.mock.callCount(), 2);
  assert.equal(requests[0].cache, 'no-store');
  assert.equal(requests[0].credentials, 'same-origin');
  assert.equal(requests[0].redirect, 'error');
  assert.ok(requests[0].signal instanceof AbortSignal);
});

test('failed and malformed capability reads fail closed without surfacing arbitrary errors', async t => {
  for (const response of [
    jsonResponse({ features: {} }), jsonResponse({ features: { bringYourOwnAI: 'true' } }),
    jsonResponse({ features: { bringYourOwnAI: true } }, 403),
    new Response('secret error', { status: 502 }),
    new Response('{"features":{"bringYourOwnAI":true}}', { headers: { 'Content-Type': 'application/notjson' } }),
  ]) {
    t.mock.method(globalThis, 'fetch', async () => response);
    const result = await loadRuntimeConfig(true);
    assert.equal(result.status, 'error');
    assert.equal(isBYOAIEnabledOnServer(), false);
    assert.doesNotMatch(JSON.stringify(result), /secret error/);
  }
  t.mock.method(globalThis, 'fetch', async () => { throw new Error('sk-never-surface-network-error'); });
  assert.doesNotMatch(JSON.stringify(await loadRuntimeConfig(true)), /sk-never/);
});

test('a stale enabled reply cannot overwrite a newer server-disabled result', async t => {
  const old = deferred<Response>();
  const latest = deferred<Response>();
  let calls = 0;
  t.mock.method(globalThis, 'fetch', () => ++calls === 1 ? old.promise : latest.promise);
  const first = loadRuntimeConfig(true);
  const second = loadRuntimeConfig(true);
  latest.resolve(capabilityResponse(false));
  await second;
  old.resolve(capabilityResponse(true));
  await first;
  assert.deepEqual(getRuntimeConfigSnapshot(), { status: 'ready', bringYourOwnAI: false });
  assert.equal(isBYOAIEnabledOnServer(), false);
});

test('a fresh concurrent capability result can settle without waiting for a slower independent read', async t => {
  const fast = deferred<Response>();
  const slow = deferred<Response>();
  let calls = 0;
  t.mock.method(globalThis, 'fetch', () => ++calls === 1 ? fast.promise : slow.promise);
  const first = loadRuntimeConfig(true);
  const second = loadRuntimeConfig(true);
  fast.resolve(capabilityResponse());
  assert.equal((await first).bringYourOwnAI, true);
  slow.resolve(capabilityResponse(false));
  await second;
  assert.equal(isBYOAIEnabledOnServer(), false);
});

test('cancellation is prompt even if fetch ignores AbortSignal, and late success cannot reopen capability', async t => {
  const pending = deferred<Response>();
  let signal: AbortSignal | undefined;
  t.mock.method(globalThis, 'fetch', (_url: unknown, options?: RequestInit) => {
    signal = options?.signal ?? undefined;
    return pending.promise;
  });
  const controller = new AbortController();
  const request = loadRuntimeConfig(true, { signal: controller.signal });
  const rejected = assert.rejects(request, { name: 'AbortError', userCancelled: true });
  controller.abort();
  await rejected;
  assert.equal(signal?.aborted, true);
  pending.resolve(capabilityResponse());
  await Promise.resolve();
  assert.equal(isBYOAIEnabledOnServer(), false);
});

test('capability timeout is bounded even for a transport that never settles', async t => {
  t.mock.timers.enable({ apis: ['setTimeout'] });
  let signal: AbortSignal | undefined;
  t.mock.method(globalThis, 'fetch', (_url: unknown, options?: RequestInit) => {
    signal = options?.signal ?? undefined;
    return new Promise<Response>(() => {});
  });
  const request = loadRuntimeConfig(true);
  t.mock.timers.tick(10_000);
  assert.equal((await request).status, 'error');
  assert.equal(signal?.aborted, true);
});

test('an already-cancelled capability request does no HTTP', async t => {
  const fetch = t.mock.method(globalThis, 'fetch', async () => capabilityResponse());
  const controller = new AbortController();
  controller.abort();
  await assert.rejects(loadRuntimeConfig(true, { signal: controller.signal }), { name: 'AbortError' });
  assert.equal(fetch.mock.callCount(), 0);
});
