// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.
const assert = require('node:assert/strict');
const { test } = require('node:test');
const express = require('express');
const { MemoryBudgetStore, createBudgetManager } = require('./ai-budget');
const { createOpenAIProxyRouter, isJsonMediaType } = require('./openai-proxy');

const alias = 'architecture-astra-production';
const endpoint = 'https://example.openai.azure.com/';
const silentLogger = { info() {}, error() {}, warn() {} };
const jsonResponse = (status, body, headers = {}) => new Response(JSON.stringify(body), {
  status, headers: { 'Content-Type': 'application/json', ...headers },
});
const requestBody = (overrides = {}) => ({
  apiFormat: 'responses', deployment: alias,
  body: { model: alias, input: [{ role: 'user', content: 'test prompt' }], max_output_tokens: 99999, store: true },
  ...overrides,
});
const manager = (options = {}) => createBudgetManager({ store: new MemoryBudgetStore(), dailyTokens: 1_000_000, ...options });

async function startServer(t, options = {}) {
  const app = express();
  app.use(express.json({ limit: '12mb' }));
  app.use('/api/openai', createOpenAIProxyRouter({
    endpoint, astraDeployment: alias, allowedDeployments: new Set([alias]),
    apiKey: 'test-only-key', logger: silentLogger, ...options,
  }));
  const server = await new Promise(resolve => {
    const listener = app.listen(0, '127.0.0.1', () => resolve(listener));
  });
  t.after(() => new Promise(resolve => { server.close(resolve); server.closeAllConnections(); }));
  return (body = requestBody(), extra = {}) => fetch(`http://127.0.0.1:${server.address().port}/api/openai`, {
    method: 'POST', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify(body), ...extra,
  });
}

test('JSON media types accept JSON and vendor JSON, not deceptive media types', () => {
  for (const type of ['application/json; charset=utf-8', 'application/problem+json']) assert.equal(isJsonMediaType(type), true);
  for (const type of ['text/json', 'application/notjson']) assert.equal(isJsonMediaType(type), false);
});

for (const [label, overrides, status, code] of [
  ['retired deployment', { deployment: 'gpt-5.6-sol' }, 403, 'deployment_not_allowed'],
  ['omitted deployment', { deployment: undefined }, 403, 'deployment_not_allowed'],
  ['omitted model', { body: { input: 'test' } }, 403, 'deployment_not_allowed'],
  ['null model', { body: { model: null, input: 'test' } }, 403, 'deployment_not_allowed'],
  ['non-string model', { body: { model: 42, input: 'test' } }, 403, 'deployment_not_allowed'],
  ['forged model', { body: { model: 'gpt-5.6-sol', input: 'test' } }, 403, 'deployment_not_allowed'],
  ['literal model is not deployment alias', { body: { model: 'gpt-6-astra', input: 'test' } }, 403, 'deployment_not_allowed'],
  ['chat completions', { apiFormat: 'chat-completions' }, 400, 'invalid_api_format'],
  ['Anthropic', { apiFormat: 'anthropic-messages' }, 400, 'invalid_api_format'],
  ['unknown format', { apiFormat: 'other' }, 400, 'invalid_api_format'],
  ['forged Responses envelope around chat', { body: { model: alias, messages: [] } }, 400, 'invalid_api_format'],
  ['forged Responses envelope around Anthropic', { body: { model: alias, thinking: {} } }, 400, 'invalid_api_format'],
  ['endpoint override', { endpoint: 'https://other.openai.azure.com/' }, 403, 'byo_not_enabled'],
  ['key override', { apiKey: 'forged-secret' }, 403, 'byo_not_enabled'],
  ['nested endpoint', { body: { input: 'test', endpoint: 'https://other.openai.azure.com/' } }, 403, 'deployment_not_allowed'],
  ['missing body', { body: null }, 400, 'missing_request_body'],
  ['array body', { body: [] }, 400, 'missing_request_body'],
]) {
  test(`managed Astra proxy rejects ${label} even with BYO enabled, before credentials, rate limits, budget and dispatch`, async t => {
    const touched = [];
    const post = await startServer(t, {
      apiKey: undefined,
      allowByoAIEndpoints: true,
      foundryEndpoint: 'https://old.services.ai.azure.com/',
      allowedFoundryDeployments: new Set(['claude-opus-5']),
      credential: { getToken() { touched.push('credential'); throw new Error('forbidden'); } },
      consumeRateLimit() { touched.push('rate'); return 0; },
      budget: { reserve() { touched.push('reserve'); throw new Error('forbidden'); } },
      fetchImpl() { touched.push('fetch'); throw new Error('forbidden'); },
    });
    const response = await post(requestBody(overrides));
    assert.equal(response.status, status);
    const payload = await response.json();
    assert.equal(payload.error.code, code);
    assert.doesNotMatch(JSON.stringify(payload), /forged-secret/);
    assert.deepEqual(touched, []);
  });
}

for (const options of [
  { astraDeployment: undefined, allowedDeployments: new Set(['gpt-5.6-sol']) },
  { allowedDeployments: new Set([alias, 'gpt-5.6-sol']) },
  { allowedDeployments: new Set(['gpt-5.6-sol']) },
  { allowedDeployments: new Set() },
  { astraDeployment: '../other' },
]) {
  test('stale allowlists cannot authorize a legacy deployment or act as implicit Astra configuration', async t => {
    let touched = false;
    const post = await startServer(t, {
      ...options, credential: { getToken() { touched = true; } },
      fetchImpl() { touched = true; }, budget: { reserve() { touched = true; } },
    });
    const response = await post();
    assert.equal(response.status, 503);
    assert.equal((await response.json()).error.code, 'astra_not_configured');
    assert.equal(touched, false);
  });
}

test('managed identity dispatch uses the configured alias and fixed Responses route with bounded output', async t => {
  const budget = manager();
  let captured;
  let scope;
  const post = await startServer(t, {
    apiKey: undefined, budget,
    credential: { async getToken(value) { scope = value; return { token: 'managed-test-token' }; } },
    fetchImpl: async (url, init) => {
      captured = { url, init };
      return jsonResponse(200, { usage: { total_tokens: 12 }, output_text: '{}' });
    },
  });
  const response = await post(requestBody({ body: { model: alias, input: 'test', max_output_tokens: 99999, stream: true, store: true } }));
  assert.equal(response.status, 200);
  await response.text();
  assert.equal(scope, 'https://cognitiveservices.azure.com/.default');
  assert.equal(captured.url, `${endpoint}openai/v1/responses`);
  assert.equal(captured.init.headers.Authorization, 'Bearer managed-test-token');
  assert.equal(captured.init.redirect, 'error');
  const body = JSON.parse(captured.init.body);
  assert.equal(body.model, alias);
  assert.equal(body.max_output_tokens, 32768);
  assert.equal(body.store, false);
  assert.equal(body.stream, false);
  assert.equal((await budget.status('local-development')).usedTokens, 12);
  assert.equal((await budget.status('local-development')).concurrentRequests, 0);
});

test('managed credential acquisition failures remain distinguishable and do not reserve tokens', async t => {
  let reserved = false;
  const post = await startServer(t, {
    apiKey: undefined, credential: { getToken() { throw new Error('secret diagnostic'); } },
    budget: { reserve() { reserved = true; } },
  });
  const response = await post();
  assert.equal(response.status, 502);
  const payload = await response.json();
  assert.equal(payload.error.code, 'credential_acquisition_failed');
  assert.doesNotMatch(JSON.stringify(payload), /secret diagnostic/);
  assert.equal(reserved, false);
});

test('Responses dispatch preserves explicit reasoning, JSON output and architecture output limits', async t => {
  const body = {
    model: alias, input: [{ role: 'user', content: 'architecture request' }],
    max_output_tokens: 32000, reasoning: { effort: 'max' },
    text: { format: { type: 'json_object' } }, store: false,
  };
  let dispatched;
  const post = await startServer(t, {
    fetchImpl: async (_url, init) => {
      dispatched = JSON.parse(init.body);
      return jsonResponse(200, { output_text: '{}' });
    },
  });
  const response = await post(requestBody({ body }));
  assert.equal(response.status, 200);
  await response.text();
  assert.deepEqual(dispatched, { ...body, stream: false });
});

for (const [status, code] of [
  [401, 'azure_openai_authentication_failed'], [403, 'azure_openai_authentication_failed'],
  [404, 'deployment_not_found'], [429, 'azure_openai_rate_limited'],
  [500, 'azure_openai_unavailable'], [502, 'azure_openai_unavailable'],
  [503, 'azure_openai_unavailable'], [504, 'azure_openai_timeout'],
]) {
  test(`upstream HTTP ${status} preserves safe diagnostics and releases concurrency`, async t => {
    const budget = manager();
    const post = await startServer(t, {
      budget, fetchImpl: async () => jsonResponse(status, { error: { message: 'sensitive-provider-detail' } }, {
        'x-ms-request-id': '11111111-1111-1111-1111-111111111111', 'x-ms-retry-after-ms': '1200',
      }),
    });
    const response = await post();
    assert.equal(response.status, status);
    assert.equal(response.headers.get('Retry-After'), '2');
    assert.equal(response.headers.get('X-Upstream-Request-Id'), '11111111-1111-1111-1111-111111111111');
    const payload = await response.json();
    assert.equal(payload.error.code, code);
    assert.doesNotMatch(JSON.stringify(payload), /sensitive-provider-detail/);
    assert.equal((await budget.status('local-development')).concurrentRequests, 0);
  });
}

test('shared rate limiting happens before credentials, budget reservation and dispatch', async t => {
  const touched = [];
  const post = await startServer(t, {
    apiKey: undefined, consumeRateLimit: async () => 60,
    credential: { getToken() { touched.push('credential'); } },
    budget: { reserve() { touched.push('reserve'); } }, fetchImpl() { touched.push('fetch'); },
  });
  const response = await post();
  assert.equal(response.status, 429);
  assert.equal(response.headers.get('Retry-After'), '60');
  assert.equal((await response.json()).error.code, 'proxy_rate_limit_exceeded');
  assert.deepEqual(touched, []);
});

for (const [label, budget, expected, code] of [
  ['exhausted', manager({ dailyTokens: 100 }), 429, 'ai_daily_budget_exceeded'],
  ['unavailable', { async reserve() { throw new Error('storage unavailable'); } }, 503, 'ai_budget_unavailable'],
]) {
  test(`budget ${label} fails closed without dispatch`, async t => {
    let dispatched = false;
    const post = await startServer(t, { budget, fetchImpl() { dispatched = true; } });
    const response = await post();
    assert.equal(response.status, expected);
    assert.ok(Number(response.headers.get('Retry-After')) > 0);
    assert.equal((await response.json()).error.code, code);
    assert.equal(dispatched, false);
  });
}

test('public AI requests still require an authenticated identity and configured budget', async t => {
  for (const budget of [manager(), undefined]) {
    const post = await startServer(t, { mode: 'public', budget, fetchImpl() { throw new Error('must not dispatch'); } });
    assert.equal((await post()).status, budget ? 401 : 503);
  }
});

test('authenticated public Astra requests retain per-user accounting', async t => {
  const budget = manager();
  const post = await startServer(t, {
    mode: 'public', budget,
    fetchImpl: async () => jsonResponse(200, { usage: { total_tokens: 17 } }),
  });
  const response = await post(requestBody(), {
    headers: {
      'Content-Type': 'application/json',
      'x-ms-client-principal-name': 'allowed@example.com',
      'x-ms-client-principal-id': 'authenticated-test-user',
    },
  });
  assert.equal(response.status, 200);
  await response.text();
  assert.equal((await budget.status('authenticated-test-user')).usedTokens, 17);
  assert.equal((await budget.status('local-development')).usedTokens, 0);
  assert.equal((await budget.status('authenticated-test-user')).concurrentRequests, 0);
});

test('non-JSON responses, transport resets and body read failures release concurrency without refunding unknown usage', async t => {
  const budget = manager();
  const outcomes = [
    () => new Response('<html>bad</html>', { headers: { 'content-type': 'text/html' } }),
    () => { throw new Error('connection reset'); },
    () => ({ ok: true, status: 200, headers: new Headers(), text: async () => { throw new Error('body reset'); } }),
  ];
  const post = await startServer(t, { budget, fetchImpl: async () => outcomes.shift()() });
  for (let n = 0; n < 3; n++) {
    const response = await post();
    assert.equal(response.status, 502);
    await response.text();
    assert.equal((await budget.status('local-development')).concurrentRequests, 0);
  }
  assert.ok((await budget.status('local-development')).usedTokens > 0);
});

test('timeout and client cancellation abort upstream and release shared concurrency', async t => {
  const budget = manager();
  let observedAbort = 0;
  let markStarted;
  let started = new Promise(resolve => { markStarted = resolve; });
  const post = await startServer(t, {
    budget, timeoutMs: 100,
    fetchImpl: async (_url, init) => new Promise((_resolve, reject) => {
      markStarted();
      init.signal.addEventListener('abort', () => { observedAbort++; reject(init.signal.reason); }, { once: true });
    }),
  });
  const response = await post();
  assert.equal(response.status, 504);
  await response.text();
  assert.equal((await budget.status('local-development')).concurrentRequests, 0);
  started = new Promise(resolve => { markStarted = resolve; });
  const controller = new AbortController();
  const pending = post(requestBody(), { signal: controller.signal }).catch(error => error);
  await started;
  controller.abort();
  await pending;
  for (let count = 0; count < 30 && (await budget.status('local-development')).concurrentRequests; count++) {
    await new Promise(resolve => setTimeout(resolve, 10));
  }
  assert.equal((await budget.status('local-development')).concurrentRequests, 0);
  assert.equal(observedAbort, 2);
});

for (const body of [
  { input: 'test', tools: [{ type: 'web_search' }] },
  { input: 'test', previous_response_id: 'stored' },
  { input: 'test', background: true },
]) {
  test('stored/background/tool requests remain rejected before credentials and reservation', async t => {
    let touched = false;
    const post = await startServer(t, {
      apiKey: undefined, credential: { getToken() { touched = true; } },
      budget: { reserve() { touched = true; } },
    });
    const response = await post(requestBody({ body: { model: alias, ...body } }));
    assert.equal(response.status, 400);
    assert.equal((await response.json()).error.code, 'unsupported_request_mode');
    assert.equal(touched, false);
  });
}
