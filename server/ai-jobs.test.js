// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.
const assert = require('node:assert/strict');
const crypto = require('node:crypto');
const { test } = require('node:test');
const express = require('express');
const { createAIJobs, createMemoryJobBackend } = require('./ai-jobs');
const { createOpenAIProxyRouter, createOpenAIExecutor } = require('./openai-proxy');
const { createBudgetManager, MemoryBudgetStore } = require('./ai-budget');
const { createAzureBlobBackend } = require('./diagram-api');

const silent = { error() {}, info() {}, warn() {} };
const envelope = () => ({
  apiFormat: 'responses', deployment: 'gpt-6-astra',
  body: { model: 'gpt-6-astra', input: 'private test prompt', reasoning: { effort: 'max' }, max_output_tokens: 32000 },
});
const success = () => new Response(JSON.stringify({ output_text: '{"services":[]}', usage: { total_tokens: 42 } }), {
  headers: { 'Content-Type': 'application/json' },
});
const deferred = () => {
  let resolve;
  const promise = new Promise(value => { resolve = value; });
  return { resolve, promise };
};
async function until(read, predicate, timeoutMs = 3000) {
  const deadline = Date.now() + timeoutMs;
  do {
    const value = await read();
    if (predicate(value)) return value;
    await new Promise(resolve => setTimeout(resolve, 10));
  } while (Date.now() < deadline);
  assert.fail('AI job did not reach the expected state.');
}

async function replica(t, { backend = createMemoryJobBackend(), budget, fetchImpl = async () => success(), ...options } = {}) {
  const jobs = createAIJobs({ backend, mode: 'public', logger: silent, heartbeatMs: 20, ...options });
  budget ||= createBudgetManager({ store: new MemoryBudgetStore(), dailyTokens: 1_000_000, concurrency: 2 });
  const app = express();
  app.use(express.json());
  app.use((req, _res, next) => {
    // Test-only identity injection on an ephemeral loopback listener.
    if (req.get('x-test-user')) req.accessPrincipal = { id: req.get('x-test-user') };
    next();
  });
  app.use('/api/openai', createOpenAIProxyRouter({
    jobs, budget, mode: 'public', logger: silent, allowByoAIEndpoints: true,
    endpoint: 'https://example.openai.azure.com/', astraDeployment: 'gpt-6-astra',
    allowedDeployments: new Set(['gpt-6-astra']), apiKey: 'test-only-managed-key', fetchImpl,
  }));
  const server = await new Promise(resolve => {
    const value = app.listen(0, '127.0.0.1', () => resolve(value));
  });
  const url = `http://127.0.0.1:${server.address().port}/api/openai`;
  t.after(async () => {
    await jobs.close();
    await new Promise(resolve => { server.close(resolve); server.closeAllConnections(); });
  });
  const request = (path, init = {}, owner = 'owner') => fetch(`${url}${path}`, {
    ...init, headers: { ...(owner ? { 'x-test-user': owner } : {}), ...init.headers },
  });
  const submit = (id = crypto.randomUUID(), body = envelope(), owner = 'owner') => request('', {
    method: 'POST', body: JSON.stringify(body),
    headers: { 'Content-Type': 'application/json', Prefer: 'respond-async', 'Idempotency-Key': id },
  }, owner);
  const status = async (id, owner) => {
    const response = await request(`/jobs/${id}`, {}, owner);
    return { status: response.status, ...(await response.json()) };
  };
  return { jobs, backend, budget, request, submit, status };
}

test('submission finishes before inference; another replica retrieves identical result and exact MAX settings', async t => {
  const backend = createMemoryJobBackend();
  const pending = deferred();
  const calls = [];
  const a = await replica(t, { backend, fetchImpl: async (url, init) => {
    calls.push({ url, body: JSON.parse(init.body) });
    return pending.promise;
  } });
  const b = await replica(t, { backend, fetchImpl: () => assert.fail('Polling dispatched inference') });
  const id = crypto.randomUUID();
  const submitted = await a.submit(id);
  assert.equal(submitted.status, 202);
  assert.equal(submitted.headers.get('cache-control'), 'no-store');
  assert.equal((await submitted.json()).job.id, id);
  await until(() => b.status(id), value => value.job.status === 'running');
  assert.equal((await b.request(`/jobs/${id}/result`)).status, 202);
  assert.equal(calls.length, 1);
  assert.deepEqual(calls[0].body.reasoning, { effort: 'max' });
  assert.equal(calls[0].body.max_output_tokens, 32000);
  assert.equal(calls[0].body.store, false);
  assert.equal(calls[0].body.stream, false);
  pending.resolve(success());
  await until(() => b.status(id), value => value.job.status === 'succeeded');
  const result = await b.request(`/jobs/${id}/result`);
  assert.equal(result.status, 200);
  assert.equal(result.headers.get('x-azurediagarm-request-id'), id);
  assert.equal((await result.json()).output_text, '{"services":[]}');
  assert.equal((await a.budget.status('owner')).usedTokens, 42);
  assert.equal((await a.budget.status('owner')).concurrentRequests, 0);
});

test('concurrent duplicate submissions across replicas dispatch exactly once; changed input conflicts', async t => {
  const backend = createMemoryJobBackend();
  const pending = deferred();
  let calls = 0;
  const options = { backend, fetchImpl: async () => { calls++; return pending.promise; } };
  const a = await replica(t, options);
  const b = await replica(t, options);
  const id = crypto.randomUUID();
  const accepted = await Promise.all(Array.from({ length: 8 }, (_, i) => (i % 2 ? a : b).submit(id)));
  assert.ok(accepted.every(response => response.status === 202));
  await until(async () => calls, value => value === 1);
  const changed = envelope();
  changed.body.input = 'different input';
  const conflict = await b.submit(id, changed);
  assert.equal(conflict.status, 409);
  assert.equal((await conflict.json()).error.code, 'ai_job_conflict');
  pending.resolve(success());
  await until(() => a.status(id), value => value.job.status === 'succeeded');
  assert.equal((await b.submit(id)).status, 202);
  assert.equal(calls, 1);
});

test('all job operations require the authenticated owner, not a caller-supplied owner ID', async t => {
  const api = await replica(t);
  const id = crypto.randomUUID();
  assert.equal((await api.submit(id)).status, 202);
  assert.equal((await api.status(id, 'other')).status, 404);
  assert.equal((await api.request(`/jobs/${id}/result`, {}, 'other')).status, 404);
  assert.equal((await api.submit(crypto.randomUUID(), envelope(), '')).status, 401);
  assert.equal((await api.status(id, '')).status, 401);
  assert.equal((await api.request(`/jobs/${id}/result`, {}, '')).status, 401);
  assert.equal((await api.request(`/jobs/${id}`, { method: 'DELETE' }, '')).status, 401);
  await api.request(`/jobs/${id}`, { method: 'DELETE' }, 'other');
  await until(() => api.status(id), value => value.job.status === 'succeeded');
});

test('remote cancellation aborts inference without refunding unknown usage', async t => {
  const backend = createMemoryJobBackend();
  let aborted = false;
  const a = await replica(t, { backend, fetchImpl: async (_url, init) => new Promise((_resolve, reject) => {
    init.signal.addEventListener('abort', () => { aborted = true; reject(init.signal.reason); }, { once: true });
  }) });
  const b = await replica(t, { backend });
  const id = crypto.randomUUID();
  await a.submit(id);
  await until(() => a.budget.status('owner'), value => value.concurrentRequests === 1);
  const charged = (await a.budget.status('owner')).usedTokens;
  const cancelled = await b.request(`/jobs/${id}`, { method: 'DELETE' });
  assert.equal(cancelled.status, 200);
  await until(() => b.status(id), value => value.job.status === 'cancelled');
  assert.equal(aborted, true);
  assert.equal((await a.budget.status('owner')).usedTokens, charged);
  assert.equal((await a.budget.status('owner')).concurrentRequests, 0);
  const result = await b.request(`/jobs/${id}/result`);
  assert.equal((await result.json()).error.code, 'ai_job_cancelled');
});

test('cancel-before-submission tombstones prevent late or lost-acknowledgment inference', async t => {
  const api = await replica(t, { fetchImpl: () => assert.fail('Cancelled request dispatched') });
  const id = crypto.randomUUID();
  assert.equal((await api.request(`/jobs/${id}`, { method: 'DELETE' })).status, 200);
  const response = await api.submit(id);
  assert.equal(response.status, 202);
  assert.equal((await response.json()).job.status, 'cancelled');
  assert.equal((await api.budget.status('owner')).usedTokens, 0);
});

test('worker loss is explicit and never replayed; result and tombstone retention is owner-scoped', async t => {
  let clock = Date.now();
  const backend = createMemoryJobBackend();
  const a = await replica(t, { backend, now: () => clock, fetchImpl: () => new Promise(() => {}) });
  const b = await replica(t, { backend, now: () => clock, fetchImpl: () => assert.fail('Interrupted job replayed') });
  const id = crypto.randomUUID();
  await a.submit(id);
  await until(() => a.budget.status('owner'), value => value.concurrentRequests === 1);
  clock += 60_001;
  assert.equal((await b.status(id)).job.status, 'failed');
  assert.equal((await (await b.request(`/jobs/${id}/result`)).json()).error.code, 'ai_job_interrupted');
  assert.equal((await b.submit(id)).status, 202);
  clock += 60 * 60_000 + 1;
  assert.equal((await b.request(`/jobs/${id}/result`)).status, 410);
  await backend.create('documents/not-an-ai-job.json', { mustRemain: true });
  clock += 24 * 60 * 60_000;
  await b.jobs.sweep();
  assert.equal((await b.status(id)).status, 404);
  assert.deepEqual((await backend.read('documents/not-an-ai-job.json')).value, { mustRemain: true });
});

test('shutdown records interruption and releases concurrency without replay or refund', async t => {
  const api = await replica(t, { fetchImpl: () => new Promise(() => {}) });
  const id = crypto.randomUUID();
  await api.submit(id);
  await until(() => api.budget.status('owner'), value => value.concurrentRequests === 1);
  const charged = (await api.budget.status('owner')).usedTokens;
  await api.jobs.close();
  assert.equal((await api.status(id)).job.status, 'failed');
  assert.equal((await (await api.request(`/jobs/${id}/result`)).json()).error.code, 'ai_job_interrupted');
  assert.equal((await api.budget.status('owner')).usedTokens, charged);
  assert.equal((await api.budget.status('owner')).concurrentRequests, 0);
  assert.equal((await api.submit()).status, 503);
});

test('job deadline is enforced even if the provider ignores abort', async t => {
  const logs = [];
  const api = await replica(t, {
    timeoutMs: 80, fetchImpl: () => new Promise(() => {}),
    logger: { error: message => logs.push(JSON.parse(message.slice(message.indexOf('{')))) },
  });
  const id = crypto.randomUUID();
  await api.submit(id);
  await until(() => api.status(id), value => value.job.status === 'failed');
  const result = await api.request(`/jobs/${id}/result`);
  assert.equal(result.status, 504);
  assert.equal((await result.json()).error.code, 'ai_job_timeout');
  const timeouts = logs.filter(entry => entry.event === 'ai_job_timeout');
  assert.equal(timeouts.length, 1);
  assert.equal(timeouts[0].requestId, id);
  assert.equal(timeouts[0].status, 'failed');
  assert.ok(timeouts[0].durationMs >= 80);
});

test('BYO input, images, keys, endpoints and raw provider errors never enter stored job diagnostics', async t => {
  const backend = createMemoryJobBackend();
  const key = 'fake-job-test-secret';
  const input = envelope();
  input.byo = { provider: 'openai', endpoint: 'https://api.openai.com/', apiKey: key };
  const logs = [];
  const api = await replica(t, {
    backend, logger: { error: message => logs.push(message) },
    fetchImpl: async () => new Response(JSON.stringify({ error: { message: `${key} private test prompt https://api.openai.com/` } }), {
      status: 401, headers: { 'Content-Type': 'application/json' },
    }),
  });
  const id = crypto.randomUUID();
  await api.submit(id, input);
  await until(() => api.status(id), value => value.job.status === 'failed');
  const result = await (await api.request(`/jobs/${id}/result`)).json();
  assert.equal(result.error.code, 'byo_authentication_failed');
  const persisted = [];
  for await (const name of backend.list('ai-jobs/')) persisted.push(await backend.read(name));
  assert.doesNotMatch(JSON.stringify([persisted, logs]), /fake-job-test-secret|private test prompt|api\.openai\.com/);
});

test('storage failures fail closed, and concurrent admissions respect the worker memory bound', async t => {
  const unavailable = await replica(t, { backend: {
    read: async () => { throw new Error('Storage unavailable'); },
  }, fetchImpl: () => assert.fail('Storage failure dispatched inference') });
  assert.equal((await unavailable.submit()).status, 503);
  const api = await replica(t, { maxActive: 2, fetchImpl: () => new Promise(() => {}) });
  const responses = await Promise.all(Array.from({ length: 10 }, () => api.submit()));
  assert.equal(responses.filter(response => response.status === 202).length, 2);
  assert.equal(responses.filter(response => response.status === 429).length, 8);
});

test('a bounded storage wait cannot hang submission or dispatch unpaid inference', async t => {
  const api = await replica(t, {
    storageTimeoutMs: 30,
    backend: { read: () => new Promise(() => {}) },
    fetchImpl: () => assert.fail('Unavailable storage dispatched'),
  });
  assert.equal((await api.submit()).status, 503);
});

test('malformed successful JSON and successful BYO credential echoes are never retained', async t => {
  for (const text of ['{malformed', '{"output_text":"fake-job-test-secret"}', '{"output_text":"\\u0066ake-job-test-secret"}']) {
    const input = envelope();
    input.byo = { provider: 'openai', endpoint: 'https://api.openai.com/', apiKey: 'fake-job-test-secret' };
    const api = await replica(t, { fetchImpl: async () => new Response(text, { headers: { 'Content-Type': 'application/json' } }) });
    const id = crypto.randomUUID();
    await api.submit(id, input);
    await until(() => api.status(id), value => value.job.status === 'failed');
    assert.doesNotMatch(await (await api.request(`/jobs/${id}/result`)).text(), /fake-job-test-secret|malformed/);
  }
});

test('lost worker storage heartbeat stops inference and persists an explicit interruption', async t => {
  const backend = createMemoryJobBackend();
  const replace = backend.replace;
  let writes = 0;
  let aborted = false;
  backend.replace = async (...args) => {
    if (++writes === 2) throw new Error('Storage heartbeat failed');
    return replace(...args);
  };
  const api = await replica(t, { backend, fetchImpl: (_url, options) => new Promise((_resolve, reject) => {
    options.signal.addEventListener('abort', () => { aborted = true; reject(options.signal.reason); }, { once: true });
  }) });
  const id = crypto.randomUUID();
  await api.submit(id);
  await until(() => api.status(id), value => value.job.status === 'failed');
  assert.equal(aborted, true);
  assert.equal((await (await api.request(`/jobs/${id}/result`)).json()).error.code, 'ai_job_interrupted');
  const status = await api.budget.status('owner');
  assert.equal(status.concurrentRequests, 0);
  assert.ok(status.usedTokens > 0);
});

test('a failed budget renewal stops a long request without an unknown-usage refund', async t => {
  t.mock.timers.enable({ apis: ['setTimeout'] });
  let dispatched = false;
  const settlements = [];
  const execute = createOpenAIExecutor({
    mode: 'public', logger: silent, endpoint: 'https://example.openai.azure.com/',
    astraDeployment: 'gpt-6-astra', allowedDeployments: new Set(['gpt-6-astra']), apiKey: 'test-only-key',
    budget: {
      reserve: async () => ({ id: 'lease' }),
      renew: async () => { throw new Error('Lease unavailable'); },
      settle: async (_owner, _lease, usage) => settlements.push(usage),
    },
    fetchImpl: async () => { dispatched = true; return new Promise(() => {}); },
  });
  const work = execute({ body: envelope(), accessPrincipal: { id: 'owner' } }, { longRunning: true, timeoutMs: 900000 });
  await new Promise(resolve => setImmediate(resolve));
  assert.equal(dispatched, true);
  t.mock.timers.tick(5000);
  const result = await work;
  assert.equal(result.body.error.code, 'ai_job_interrupted');
  assert.deepEqual(settlements, [undefined]);
});

test('version-aware cleanup removes only prior job versions, including version-only orphan records', async t => {
  const backend = createMemoryJobBackend();
  const name = `ai-jobs/${'a'.repeat(64)}/${crypto.randomUUID()}.json`;
  const removed = [];
  backend.listJobBlobs = async function* () {
    yield { name, versionId: 'old', isCurrentVersion: false };
    yield { name, versionId: 'current', isCurrentVersion: true };
    yield { name, versionId: 'soft-deleted', isCurrentVersion: false, deleted: true };
  };
  backend.removeJobVersion = async (...args) => { removed.push(args); };
  await backend.create(name, { id: name.split('/').at(-1).slice(0, -5), status: 'running', workerExpiresAt: Date.now() + 60000, deadlineAt: Date.now() + 900000 });
  const api = await replica(t, { backend });
  await api.jobs.sweep();
  assert.deepEqual(removed.map(value => value.slice(0, 2)), [[name, 'old']]);
  assert.ok((await backend.read(name)).value);
  await backend.remove(name);
  await api.jobs.sweep();
  assert.equal(removed.length, 2, 'old versions remain discoverable without a current blob');
});

test('version cleanup errors surface, remain retryable, and never overlap a second sweep', async t => {
  const backend = createMemoryJobBackend();
  const name = `ai-jobs/${'a'.repeat(64)}/${crypto.randomUUID()}.json`;
  let listings = 0;
  let fail = true;
  backend.listJobBlobs = async function* () {
    listings++;
    yield { name, versionId: 'old', isCurrentVersion: false };
  };
  backend.removeJobVersion = async () => { if (fail) throw new Error('Version delete failed'); };
  const api = await replica(t, { backend });
  const first = api.jobs.sweep();
  await api.jobs.sweep();
  await assert.rejects(first, /Version delete failed/);
  assert.equal(listings, 1);
  fail = false;
  await api.jobs.sweep();
  assert.equal(listings, 2);
});

test('Azure version enumeration and deletion cannot escape the private job namespace', async () => {
  const name = `ai-jobs/${'a'.repeat(64)}/${crypto.randomUUID()}.json`;
  const removed = [];
  const backend = createAzureBlobBackend({ containerClient: {
    async *listBlobsFlat(options) {
      assert.deepEqual(options, { prefix: 'ai-jobs/', includeVersions: true });
      yield { name, versionId: 'old', isCurrentVersion: false };
      yield { name: 'owners/one/documents/one/current.json', versionId: 'protected', isCurrentVersion: false };
    },
    getBlockBlobClient: blob => ({
      withVersion: version => ({
        deleteIfExists: async options => { removed.push({ blob, version, options }); return { succeeded: true }; },
      }),
    }),
  } });
  const entries = [];
  for await (const entry of backend.listJobBlobs()) entries.push(entry);
  assert.equal(entries.length, 1);
  const controller = new AbortController();
  await backend.removeJobVersion(name, 'old', { abortSignal: controller.signal });
  assert.equal(removed[0].options.abortSignal, controller.signal);
  await assert.rejects(backend.removeJobVersion('owners/one/documents/one/current.json', 'protected'), /restricted/);
  await assert.rejects(backend.removeJobVersion(name, ''), /restricted/);
  assert.equal(removed.length, 1);
});
