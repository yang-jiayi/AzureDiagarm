// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.
const assert = require('node:assert/strict');
const { test } = require('node:test');
const express = require('express');
const { createOpenAIProxyRouter, resolveByoRequestConfig } = require('./openai-proxy');
const { MemoryBudgetStore, createBudgetManager, reservationTokens } = require('./ai-budget');
const { createOriginGuard } = require('./deployment-security');
const { createAccessControlRouter } = require('./access-control');

const userKey = 'user-only-test-key';
const managedAlias = 'approved-astra';
const managedEndpoint = 'https://managed.openai.azure.com/';
const customEndpoint = 'https://custom.openai.azure.com';
const customModel = 'custom-model:version-1';
const origin = 'https://application.example';
const silentLogger = { info() {}, warn() {}, error() {} };
const manager = (options = {}) => createBudgetManager({ store: new MemoryBudgetStore(), dailyTokens: 1_000_000, ...options });
const jsonResponse = (status = 200, body = { output_text: 'OK', usage: { total_tokens: 8 } }, headers = {}) => (
  new Response(JSON.stringify(body), { status, headers: { 'Content-Type': 'application/json', ...headers } })
);
const request = (overrides = {}) => ({
  apiFormat: 'responses', deployment: customModel,
  body: { model: customModel, input: 'Reply OK', max_output_tokens: 8 },
  byo: { provider: 'azure-openai', endpoint: customEndpoint, apiKey: userKey },
  ...overrides,
});

async function startServer(t, options = {}) {
  const app = express();
  app.use(express.json({ limit: '12mb' }));
  if (options.mode === 'public') {
    const access = createAccessControlRouter({
      enabled: true, adminEmail: 'admin@example.com', publicAppUrl: origin,
      table: { async *listEntities() { yield { email: 'allowed@example.com' }; } },
      logger: silentLogger,
    });
    app.use('/api', access.requireAllowed, createOriginGuard({ mode: 'public', origin }));
  }
  app.use('/api/openai', createOpenAIProxyRouter({
    allowByoAIEndpoints: true, logger: silentLogger,
    credential: { getToken() { throw new Error('Managed credentials must not be requested.'); } },
    fetchImpl: () => { throw new Error('Unexpected upstream dispatch.'); },
    ...options,
  }));
  const server = await new Promise(resolve => {
    const listener = app.listen(0, '127.0.0.1', () => resolve(listener));
  });
  t.after(() => new Promise(resolve => { server.close(resolve); server.closeAllConnections(); }));
  return (body = request(), extra = {}) => fetch(`http://127.0.0.1:${server.address().port}/api/openai`, {
    method: 'POST', body: JSON.stringify(body), ...extra,
    headers: { 'Content-Type': 'application/json', ...extra.headers },
  });
}

test('explicitly opted-in BYO works with no managed configuration or credential', async t => {
  const budget = manager();
  const post = await startServer(t, {
    credential: undefined, budget, fetchImpl: async () => jsonResponse(),
  });
  const response = await post();
  assert.equal(response.status, 200);
  await response.text();
  assert.equal((await budget.status('local-development')).usedTokens, 8);
  const managed = await post(request({ byo: undefined, deployment: 'gpt-6-astra' }));
  assert.equal(managed.status, 503);
  assert.equal((await managed.json()).error.code, 'astra_not_configured');
});

for (const provider of ['azure-openai', 'openai']) {
  for (const apiFormat of ['responses', 'chat-completions']) {
    test(`${provider} ${apiFormat} preserves model, vision and profile settings using only the user's key`, async t => {
      const budget = manager();
      const image = 'data:image/png;base64,AA==';
      const body = apiFormat === 'responses' ? {
        model: customModel,
        input: [{ role: 'user', content: [{ type: 'input_text', text: 'private diagram' }, { type: 'input_image', image_url: image, detail: 'high' }] }],
        max_output_tokens: 32000, reasoning: { effort: 'high' }, text: { format: { type: 'json_object' } },
        store: true, stream: true,
      } : {
        model: customModel,
        messages: [{ role: 'user', content: [{ type: 'text', text: 'private diagram' }, { type: 'image_url', image_url: { url: image, detail: 'high' } }] }],
        max_completion_tokens: 32000, reasoning_effort: 'high', response_format: { type: 'json_object' },
        store: true, stream: true,
      };
      let captured;
      let acquired = 0;
      const post = await startServer(t, {
        endpoint: managedEndpoint, astraDeployment: managedAlias, allowedDeployments: new Set([managedAlias]),
        apiKey: 'server-only-test-key', credential: { getToken() { acquired++; } }, budget,
        fetchImpl: async (url, init) => {
          captured = { url, init };
          return jsonResponse(200, {
            usage: apiFormat === 'responses' ? { input_tokens: 11, output_tokens: 7 } : { prompt_tokens: 11, completion_tokens: 7 },
          });
        },
      });
      const endpoint = provider === 'openai' ? 'https://api.openai.com' : customEndpoint;
      const response = await post(request({ apiFormat, body, byo: { provider, endpoint, apiKey: userKey } }));
      assert.equal(response.status, 200);
      await response.text();
      assert.equal(acquired, 0);
      assert.equal(captured.url, `${endpoint}/${provider === 'azure-openai' ? 'openai/' : ''}v1/${apiFormat === 'responses' ? 'responses' : 'chat/completions'}`);
      assert.deepEqual(captured.init.headers, {
        'Content-Type': 'application/json',
        ...(provider === 'openai' ? { Authorization: `Bearer ${userKey}` } : { 'api-key': userKey }),
      });
      assert.equal(captured.init.redirect, 'error');
      assert.ok(captured.init.signal instanceof AbortSignal);
      assert.deepEqual(JSON.parse(captured.init.body), { ...body, store: false, stream: false });
      assert.ok(reservationTokens(JSON.parse(captured.init.body), apiFormat) > 163000);
      assert.equal((await budget.status('local-development')).usedTokens, 18);
      assert.equal((await budget.status('local-development')).concurrentRequests, 0);
    });
  }
}

test('Chat Completions preserves explicit max_tokens rather than inventing a reasoning-model field', async t => {
  let body;
  const post = await startServer(t, {
    fetchImpl: async (_url, init) => { body = JSON.parse(init.body); return jsonResponse(); },
  });
  assert.equal((await post(request({
    apiFormat: 'chat-completions', body: { model: customModel, messages: [{ role: 'user', content: 'Hello' }], max_tokens: 99999 },
  }))).status, 200);
  assert.equal(body.max_tokens, 32768);
  assert.equal(Object.hasOwn(body, 'max_completion_tokens'), false);
  assert.equal(Object.hasOwn(body, 'max_output_tokens'), false);
});

for (const flag of [undefined, false, 'true', 'false', 'TRUE', 1, null]) {
  test(`stale client capability cannot enable BYO when server opt-in is ${String(flag)}`, async t => {
    const touched = [];
    const post = await startServer(t, {
      allowByoAIEndpoints: flag,
      endpoint: managedEndpoint, astraDeployment: managedAlias, allowedDeployments: new Set([managedAlias]),
      credential: { getToken() { touched.push('credential'); } },
      consumeRateLimit() { touched.push('rate'); },
      budget: { reserve() { touched.push('budget'); } },
      fetchImpl() { touched.push('fetch'); },
    });
    for (const byo of [request().byo, null, {}, { provider: 'openai', endpoint: 'https://api.openai.com', apiKey: userKey }]) {
      const response = await post(request({ byo }));
      assert.equal(response.status, 403);
      assert.equal((await response.json()).error.code, 'byo_not_enabled');
    }
    assert.deepEqual(touched, []);
  });
}

const invalidCases = [
  ['null configuration', { byo: null }, 'invalid_byo_configuration'],
  ['array configuration', { byo: [] }, 'invalid_byo_configuration'],
  ['string configuration', { byo: 'secret' }, 'invalid_byo_configuration'],
  ['extra credential override', { byo: { ...request().byo, useManagedIdentity: true } }, 'invalid_byo_configuration'],
  ['retired Foundry provider', { byo: { ...request().byo, provider: 'foundry' } }, 'invalid_byo_provider'],
  ['arbitrary compatible provider', { byo: { ...request().byo, provider: 'openai-compatible' } }, 'invalid_byo_provider'],
  ...[undefined, '', 'short', ' key-with-whitespace', 'key\nline', 'key\u0000control', 'x'.repeat(513)].map(apiKey => (
    ['invalid key', { byo: { ...request().byo, apiKey } }, 'invalid_byo_api_key']
  )),
  ...[undefined, null, 1, '', '../model', 'model/name', ' model ', 'x'.repeat(129), '..'].map(deployment => (
    ['invalid model ID', { deployment }, 'invalid_deployment_name']
  )),
  ...['anthropic-messages', 'unknown', undefined].map(apiFormat => (
    ['invalid format', { apiFormat }, 'invalid_api_format']
  )),
  ['missing model', { body: { input: 'Hello' } }, 'deployment_not_allowed'],
  ['mismatched model', { body: { model: 'other', input: 'Hello' } }, 'deployment_not_allowed'],
  ['managed model substitution', { body: { model: managedAlias, input: 'Hello' } }, 'deployment_not_allowed'],
  ['nested credential', { body: { ...request().body, apiKey: 'another-key' } }, 'deployment_not_allowed'],
  ['nested endpoint', { body: { ...request().body, endpoint: managedEndpoint } }, 'deployment_not_allowed'],
  ['Chat in Responses envelope', { body: { model: customModel, messages: [], max_tokens: 8 } }, 'invalid_api_format'],
  ['Responses in Chat envelope', { apiFormat: 'chat-completions' }, 'invalid_api_format'],
  ['ambiguous Chat output caps', { apiFormat: 'chat-completions', body: { model: customModel, messages: [], max_tokens: 8, max_completion_tokens: 8 } }, 'invalid_api_format'],
];
for (const [label, overrides, code] of invalidCases) {
  test(`invalid BYO ${label} fails before rate, credentials, reservation or upstream without managed fallthrough`, async t => {
    const touched = [];
    const post = await startServer(t, {
      endpoint: managedEndpoint, astraDeployment: managedAlias, allowedDeployments: new Set([managedAlias]),
      apiKey: 'server-only-test-key',
      credential: { getToken() { touched.push('credential'); } },
      consumeRateLimit() { touched.push('rate'); },
      budget: { reserve() { touched.push('budget'); } },
      fetchImpl() { touched.push('fetch'); },
    });
    const response = await post(request(overrides));
    assert.equal(response.status, code === 'deployment_not_allowed' ? 403 : 400);
    const payload = await response.json();
    assert.equal(payload.error.code, code);
    assert.equal(payload.error.requestId, response.headers.get('X-AzureDiagarm-Request-Id'));
    assert.deepEqual(touched, []);
  });
}

test('BYO Azure hosts are limited to resource origins in trusted Azure clouds', () => {
  for (const suffix of [
    'openai.azure.com', 'openai.azure.us', 'openai.azure.cn',
    'cognitiveservices.azure.com', 'cognitiveservices.azure.us', 'cognitiveservices.azure.cn',
    'services.ai.azure.com', 'services.ai.azure.us', 'services.ai.azure.cn',
  ]) {
    assert.equal(resolveByoRequestConfig({ ...request().byo, endpoint: `https://resource.${suffix}` }, true).endpoint, `https://resource.${suffix}/`);
  }
});

for (const endpoint of [
  undefined, null, 'https://localhost', 'https://127.0.0.1', 'https://[::1]',
  'https://arbitrary.example', 'https://api.openai.com.attacker.example',
  'https://resource.openai.azure.com.attacker.example', 'http://resource.openai.azure.com',
  'https://resource.openai.azure.com:443', 'https://resource.openai.azure.com:8443',
  'https://resource.openai.azure.com/openai/v1', 'https://resource.openai.azure.com/../',
  'https://resource.openai.azure.com?', 'https://resource.openai.azure.com#',
  'https://resource.openai.azure.com?key=secret', 'https://resource.openai.azure.com#key=secret',
  'https://user:secret@resource.openai.azure.com', 'https://@resource.openai.azure.com',
  ' https://resource.openai.azure.com', 'https://resource.openai.azure.com\\',
  'https://.openai.azure.com', 'https://resource..openai.azure.com', 'https://resource.openai.azure.com.',
]) {
  test(`untrusted or noncanonical BYO origin is rejected offline: ${String(endpoint)}`, () => {
    for (const provider of ['azure-openai', 'openai']) {
      assert.throws(() => resolveByoRequestConfig({ provider, endpoint, apiKey: userKey }, true), { code: 'invalid_byo_endpoint' });
    }
  });
}
test('providers cannot reinterpret each other’s endpoint', () => {
  assert.throws(() => resolveByoRequestConfig({ ...request().byo, endpoint: 'https://api.openai.com' }, true), { code: 'invalid_byo_endpoint' });
  assert.throws(() => resolveByoRequestConfig({ ...request().byo, provider: 'openai' }, true), { code: 'invalid_byo_endpoint' });
  for (const endpoint of ['https://api.openai.com/v1', 'https://api.openai.com:443', 'https://api.openai.com?', 'https://api.openai.com#']) {
    assert.throws(() => resolveByoRequestConfig({ provider: 'openai', endpoint, apiKey: userKey }, true), { code: 'invalid_byo_endpoint' });
  }
});

for (const options of [
  { endpoint: managedEndpoint }, { astraDeployment: managedAlias },
  { allowedDeployments: new Set(['gpt-5.6-sol']) }, { apiKey: 'orphan-server-key' },
  { endpoint: 'https://api.openai.com', astraDeployment: managedAlias, allowedDeployments: new Set([managedAlias]) },
]) {
  test('opting into BYO does not hide partial or invalid managed configuration', async t => {
    const post = await startServer(t, options);
    const response = await post();
    assert.equal(response.status, 503);
    assert.ok(['astra_not_configured', 'proxy_not_configured'].includes((await response.json()).error.code));
  });
}

for (const body of [
  { input: 'text', background: true }, { input: 'text', previous_response_id: 'stored' },
  { input: 'text', conversation: 'stored' }, { input: 'text', prompt: { id: 'stored' } },
  { input: 'text', tools: [{ type: 'web_search' }] }, { input: 'text', container: { id: 'stored' } },
  { input: [{ type: 'input_image', image_url: 'https://remote.example/image.png' }] },
  { input: [{ type: 'input_file', file_id: 'stored' }] },
  { messages: [{ role: 'user', content: [{ type: 'image_url', image_url: { url: 'https://remote.example/image.png' } }] }] },
  { messages: [], functions: [{ name: 'remote_tool' }] },
  { messages: [], modalities: ['text', 'audio'] }, { messages: [], audio: { format: 'mp3' } },
  { messages: [], n: 2 }, { messages: [{ content: [{ type: 'input_audio', input_audio: { data: 'AA==' } }] }] },
]) {
  test('BYO cannot bypass complete-response inline-input and tool/background metering guards', async t => {
    const touched = [];
    const post = await startServer(t, {
      consumeRateLimit() { touched.push('rate'); },
      budget: { reserve() { touched.push('budget'); } }, fetchImpl() { touched.push('fetch'); },
    });
    const response = await post(request({
      apiFormat: Object.hasOwn(body, 'messages') ? 'chat-completions' : 'responses',
      body: { model: customModel, ...body },
    }));
    assert.equal(response.status, 400);
    assert.equal((await response.json()).error.code, 'unsupported_request_mode');
    assert.deepEqual(touched, []);
  });
}

test('connection probes use the normal rate limit and budget, with no special bypass', async t => {
  for (const options of [
    { consumeRateLimit: async () => 60 },
    { budget: manager({ dailyTokens: 1 }) },
    { budget: { reserve() { throw new Error('private-storage-diagnostic'); } } },
  ]) {
    let dispatched = false;
    const post = await startServer(t, { ...options, fetchImpl() { dispatched = true; } });
    const response = await post(request({ connectionTest: true, skipBudget: true }));
    assert.ok([429, 503].includes(response.status));
    assert.ok(Number(response.headers.get('Retry-After')) > 0);
    assert.equal(dispatched, false);
    assert.doesNotMatch(await response.text(), /private-storage-diagnostic/);
  }
});

test('managed generation and a BYO connection probe share the same identity concurrency budget', async t => {
  const budget = manager({ concurrency: 1 });
  const lease = await budget.reserve('local-development', 100);
  const post = await startServer(t, {
    budget, endpoint: managedEndpoint, astraDeployment: managedAlias, allowedDeployments: new Set([managedAlias]),
  });
  const response = await post();
  assert.equal(response.status, 429);
  assert.equal((await response.json()).error.code, 'ai_concurrency_limit');
  assert.equal((await budget.status('local-development')).concurrentRequests, 1);
  await budget.settle('local-development', lease, 0);
});

test('public BYO requires authentication, access-list membership, exact origin and a shared budget', async t => {
  const budget = manager();
  let calls = 0;
  const post = await startServer(t, { mode: 'public', budget, fetchImpl: async () => { calls++; return jsonResponse(); } });
  const principal = {
    Origin: origin, 'x-ms-client-principal-name': 'allowed@example.com', 'x-ms-client-principal-id': 'allowed-id',
  };
  assert.equal((await post()).status, 401);
  assert.equal((await post(request(), { headers: { ...principal, 'x-ms-client-principal-name': 'outsider@example.com' } })).status, 403);
  assert.equal((await post(request(), { headers: { ...principal, Origin: 'https://untrusted.example' } })).status, 403);
  assert.equal((await post(request(), { headers: { ...principal, Origin: '' } })).status, 403);
  assert.equal(calls, 0);
  const response = await post(request(), { headers: principal });
  assert.equal(response.status, 200);
  await response.text();
  assert.equal((await budget.status('allowed-id')).usedTokens, 8);
  assert.equal((await budget.status('local-development')).usedTokens, 0);
  const unbudgeted = await startServer(t, { mode: 'public' });
  assert.equal((await unbudgeted(request(), { headers: principal })).status, 503);
});

for (const [status, code] of [
  [401, 'byo_authentication_failed'], [403, 'byo_authentication_failed'], [404, 'deployment_not_found'],
  [429, 'byo_rate_limited'], [500, 'byo_unavailable'], [502, 'byo_unavailable'], [503, 'byo_unavailable'],
  [504, 'byo_timeout'], [400, 'invalid_upstream_request'],
]) {
  test(`BYO HTTP ${status} exposes stable diagnostics, never user keys/endpoints/prompts`, async t => {
    const logs = [];
    const budget = manager();
    const logger = { info: value => logs.push(value), error: value => logs.push(value), warn: value => logs.push(value) };
    const secret = `${userKey} ${customEndpoint} private-probe-text`;
    const post = await startServer(t, {
      budget, logger, fetchImpl: async () => jsonResponse(status, { error: { code: secret, message: secret } }, {
        'x-request-id': userKey, 'apim-request-id': customEndpoint,
        'retry-after': secret, 'x-ms-retry-after-ms': '1200',
      }),
    });
    const response = await post(request({ body: { ...request().body, input: 'private-probe-text' } }));
    const payload = await response.json();
    assert.equal(payload.error.code, code);
    assert.equal(payload.error.requestId, response.headers.get('X-AzureDiagarm-Request-Id'));
    assert.equal(response.headers.get('X-Upstream-Request-Id'), null);
    assert.equal(response.headers.get('Retry-After'), '2');
    const diagnostics = JSON.stringify({ payload, logs, headers: [...response.headers] });
    for (const value of [userKey, customEndpoint, customModel, 'private-probe-text']) assert.equal(diagnostics.includes(value), false);
    assert.equal((await budget.status('local-development')).concurrentRequests, 0);
    assert.equal((await budget.status('local-development')).usedTokens > 0, status >= 500);
  });
}

test('a rate-limited BYO retry is a fresh metered request without model/provider/format fallback', async t => {
  const budget = manager();
  const captured = [];
  let rates = 0;
  const post = await startServer(t, {
    budget, consumeRateLimit() { rates++; return 0; },
    fetchImpl: async (url, init) => {
      captured.push({ url, body: init.body });
      return captured.length === 1 ? jsonResponse(429, { error: {} }, { 'retry-after': '1' }) : jsonResponse();
    },
  });
  const first = await post();
  assert.equal(first.status, 429);
  assert.equal((await first.json()).error.code, 'byo_rate_limited');
  const second = await post();
  assert.equal(second.status, 200);
  await second.text();
  assert.equal(rates, 2);
  assert.equal(captured.length, 2);
  assert.deepEqual(captured[0], captured[1]);
  assert.equal((await budget.status('local-development')).usedTokens, 8);
});

test('timeouts and cancellation abort BYO upstream without refunds for unknown usage', async t => {
  const budget = manager();
  let aborts = 0;
  let markStarted;
  let started = new Promise(resolve => { markStarted = resolve; });
  const post = await startServer(t, {
    budget, timeoutMs: 100,
    fetchImpl: async (_url, init) => new Promise((_resolve, reject) => {
      markStarted();
      init.signal.addEventListener('abort', () => { aborts++; reject(init.signal.reason); }, { once: true });
    }),
  });
  const timedOut = await post();
  assert.equal(timedOut.status, 504);
  assert.equal((await timedOut.json()).error.code, 'byo_timeout');
  const firstCharge = (await budget.status('local-development')).usedTokens;
  assert.ok(firstCharge > 0);
  started = new Promise(resolve => { markStarted = resolve; });
  const controller = new AbortController();
  const pending = post(request(), { signal: controller.signal }).catch(error => error);
  await started;
  controller.abort();
  await pending;
  for (let attempt = 0; attempt < 40 && (await budget.status('local-development')).concurrentRequests; attempt++) {
    await new Promise(resolve => setTimeout(resolve, 10));
  }
  assert.equal(aborts, 2);
  assert.equal((await budget.status('local-development')).concurrentRequests, 0);
  assert.equal((await budget.status('local-development')).usedTokens, firstCharge * 2);
});

test('BYO transport/redirect rejection and malformed usage keep reservations without leaking diagnostics', async t => {
  const budget = manager();
  const logs = [];
  const secretError = Object.assign(new Error(userKey), { name: customEndpoint, code: userKey });
  const outcomes = [
    () => { throw secretError; },
    () => ({ ok: true, status: 200, headers: new Headers(), text: async () => { throw secretError; } }),
    () => new Response('<html>private</html>', { headers: { 'Content-Type': 'text/html' } }),
    () => jsonResponse(200, { usage: { prompt_tokens: 1 } }),
  ];
  const post = await startServer(t, {
    budget, logger: { info: value => logs.push(value), error: value => logs.push(value), warn() {} },
    fetchImpl: async (_url, init) => { assert.equal(init.redirect, 'error'); return outcomes.shift()(); },
  });

  await t.test('cancelling during a BYO reservation prevents dispatch and releases the unused lease', async t => {
    let entered;
    const started = new Promise(resolve => { entered = resolve; });
    let release;
    const waiting = new Promise(resolve => { release = resolve; });
    let settled;
    const finished = new Promise(resolve => { settled = resolve; });
    let dispatched = false;
    const lease = { id: 'test-lease', tokens: 5000 };
    const post = await startServer(t, {
      budget: {
        async reserve() { entered(); await waiting; return lease; },
        async settle(identity, actualLease, usage) {
          assert.equal(identity, 'local-development');
          assert.equal(actualLease, lease);
          assert.equal(usage, 0);
          settled();
        },
      },
      fetchImpl() { dispatched = true; },
    });
    const controller = new AbortController();
    const pending = post(request(), { signal: controller.signal }).catch(error => error);
    await started;
    controller.abort();
    await pending;
    await new Promise(resolve => setTimeout(resolve, 20));
    release();
    await finished;
    assert.equal(dispatched, false);
  });
  let lastUsage = 0;
  for (let count = 0; count < 4; count++) {
    const response = await post(request({ apiFormat: 'chat-completions', body: { model: customModel, messages: [], max_tokens: 8 } }));
    assert.equal(response.status, count === 3 ? 200 : 502);
    await response.text();
    const status = await budget.status('local-development');
    assert.ok(status.usedTokens > lastUsage);
    assert.equal(status.concurrentRequests, 0);
    lastUsage = status.usedTokens;
  }
  for (const value of [userKey, customEndpoint, customModel]) assert.equal(logs.join('\n').includes(value), false);
});
