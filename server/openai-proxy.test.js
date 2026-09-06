// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

const assert = require('node:assert/strict');
const { test } = require('node:test');
const express = require('express');
const { MemoryBudgetStore, createBudgetManager } = require('./ai-budget');
const {
  createOpenAIProxyRouter,
  isJsonMediaType,
  logFoundryConfiguration,
  normalizeAzureOpenAIEndpoint,
} = require('./openai-proxy');

async function startServer(options) {
  const app = express();
  app.use(express.json({ limit: '12mb' }));
  app.use('/api/openai', createOpenAIProxyRouter(options));
  const server = await new Promise((resolve) => {
    const listening = app.listen(0, '127.0.0.1', () => resolve(listening));
  });
  const { port } = server.address();
  return {
    baseUrl: `http://127.0.0.1:${port}`,
    close: () => new Promise((resolve, reject) => {
      server.close((error) => (error ? reject(error) : resolve()));
    }),
  };
}

function requestBody(overrides = {}) {
  return {
    apiFormat: 'responses',
    deployment: 'gpt-5.6-sol',
    body: {
      model: 'client-supplied-model',
      input: [{ role: 'user', content: 'test prompt' }],
      max_output_tokens: 99_999,
      store: true,
    },
    ...overrides,
  };
}

function anthropicRequestBody(overrides = {}) {
  return {
    apiFormat: 'anthropic-messages',
    deployment: 'claude-opus-5',
    body: {
      model: 'client-supplied-model',
      max_tokens: 99_999,
      messages: [{ role: 'user', content: [{ type: 'text', text: 'test prompt' }] }],
      thinking: { type: 'disabled' },
      output_config: { effort: 'xhigh' },
      stream: true,
    },
    ...overrides,
  };
}

function byoRequestBody(overrides = {}) {
  return requestBody({
    deployment: 'my-gpt-deployment',
    byo: {
      provider: 'azure-openai',
      endpoint: 'https://customer-resource.openai.azure.com/',
      apiKey: 'customer-secret-key',
    },
    ...overrides,
  });
}

function jsonResponse(status, body, headers = {}) {
  return new Response(JSON.stringify(body), {
    status,
    headers: { 'Content-Type': 'application/json', ...headers },
  });
}

const silentLogger = { info() {}, error() {} };

test('Foundry logging treats an intentionally disabled optional provider as informational', () => {
  const messages = { info: [], warn: [] };
  const logger = {
    info(message) {
      messages.info.push(message);
    },
    warn(message) {
      messages.warn.push(message);
    },
  };

  logFoundryConfiguration(undefined, new Set(), logger);

  assert.equal(messages.info.length, 1);
  assert.match(messages.info[0], /optional Microsoft Foundry provider is disabled/i);
  assert.deepEqual(messages.warn, []);
});

test('Foundry logging warns only about partial configuration', () => {
  const messages = { info: [], warn: [] };
  const logger = {
    info(message) {
      messages.info.push(message);
    },
    warn(message) {
      messages.warn.push(message);
    },
  };

  logFoundryConfiguration('https://example.services.ai.azure.com', new Set(), logger);
  assert.deepEqual(messages.info, []);
  assert.equal(messages.warn.length, 1);
  assert.match(messages.warn[0], /AZURE_FOUNDRY_ALLOWED_DEPLOYMENTS is empty/);

  messages.warn.length = 0;
  logFoundryConfiguration(undefined, new Set(['claude-opus-5']), logger);
  assert.equal(messages.warn.length, 1);
  assert.match(messages.warn[0], /AZURE_FOUNDRY_ENDPOINT is not set/);
});

test('BYO endpoint validation accepts only trusted Azure OpenAI HTTPS hosts', () => {
  assert.equal(
    normalizeAzureOpenAIEndpoint('https://customer-resource.openai.azure.com/'),
    'https://customer-resource.openai.azure.com/',
  );
  assert.equal(
    normalizeAzureOpenAIEndpoint('https://customer-resource.services.ai.azure.com/'),
    'https://customer-resource.services.ai.azure.com/',
  );
  assert.throws(
    () => normalizeAzureOpenAIEndpoint('http://customer-resource.openai.azure.com/'),
    error => error.code === 'invalid_byo_endpoint',
  );
  assert.throws(
    () => normalizeAzureOpenAIEndpoint('https://customer-resource.openai.azure.com.attacker.example/'),
    error => error.code === 'invalid_byo_endpoint',
  );
  assert.throws(
    () => normalizeAzureOpenAIEndpoint('https://customer-resource.openai.azure.com/openai/v1'),
    error => error.code === 'invalid_byo_endpoint',
  );
});

test('JSON media type validation accepts JSON and vendor JSON only', () => {
  assert.equal(isJsonMediaType('application/json; charset=utf-8'), true);
  assert.equal(isJsonMediaType('application/problem+json'), true);
  assert.equal(isJsonMediaType('application/notjson'), false);
  assert.equal(isJsonMediaType('text/json'), false);
});

test('OpenAI proxy rejects BYO requests unless the server enables them', async (t) => {
  const server = await startServer({
    endpoint: 'https://example.openai.azure.com/',
    apiKey: 'test-key',
    allowedDeployments: new Set(['approved']),
    logger: silentLogger,
  });
  t.after(server.close);

  const response = await fetch(`${server.baseUrl}/api/openai`, {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify(byoRequestBody()),
  });

  assert.equal(response.status, 403);
  assert.equal((await response.json()).error.code, 'byo_not_enabled');
});

test('OpenAI proxy routes BYO Azure OpenAI requests without logging credentials', async (t) => {
  let captured;
  const events = [];
  const server = await startServer({
    allowByoAIEndpoints: true,
    fetchImpl: async (url, init) => {
      captured = { url, init };
      return jsonResponse(200, {
        model: 'customer-model',
        output_text: '{}',
        usage: {},
      });
    },
    logger: {
      info(message) { events.push(message); },
      error(message) { events.push(message); },
    },
  });
  t.after(server.close);

  const response = await fetch(`${server.baseUrl}/api/openai`, {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify(byoRequestBody()),
  });

  assert.equal(response.status, 200);
  assert.equal(
    captured.url,
    'https://customer-resource.openai.azure.com/openai/v1/responses',
  );
  assert.equal(captured.init.headers['api-key'], 'customer-secret-key');
  assert.equal(captured.init.headers.Authorization, undefined);
  assert.equal(captured.init.redirect, 'error');
  assert.doesNotMatch(
    events.join('\n'),
    /customer-secret-key|customer-resource|my-gpt-deployment/,
  );
});

test('OpenAI proxy routes BYO official OpenAI chat requests to the fixed API host', async (t) => {
  let captured;
  const server = await startServer({
    allowByoAIEndpoints: true,
    fetchImpl: async (url, init) => {
      captured = { url, init };
      return jsonResponse(200, {
        model: 'gpt-custom',
        choices: [{ message: { content: '{}' } }],
        usage: {},
      });
    },
    logger: silentLogger,
  });
  t.after(server.close);

  const response = await fetch(`${server.baseUrl}/api/openai`, {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify(byoRequestBody({
      apiFormat: 'chat-completions',
      deployment: 'gpt-custom',
      byo: {
        provider: 'openai',
        endpoint: 'https://api.openai.com',
        apiKey: 'sk-customer-secret',
      },
    })),
  });

  assert.equal(response.status, 200);
  assert.equal(captured.url, 'https://api.openai.com/v1/chat/completions');
  assert.equal(captured.init.headers.Authorization, 'Bearer sk-customer-secret');
  assert.equal(captured.init.headers['api-key'], undefined);
  assert.equal(JSON.parse(captured.init.body).model, 'gpt-custom');
});

test('OpenAI proxy routes BYO Azure reasoning chat through v1 with compatible parameters', async (t) => {
  let captured;
  const server = await startServer({
    allowByoAIEndpoints: true,
    fetchImpl: async (url, init) => {
      captured = { url, init };
      return jsonResponse(200, {
        model: 'gpt-5-customer',
        choices: [{ message: { content: '{}' } }],
        usage: {},
      });
    },
    logger: silentLogger,
  });
  t.after(server.close);

  const response = await fetch(`${server.baseUrl}/api/openai`, {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify(byoRequestBody({
      apiFormat: 'chat-completions',
      deployment: 'gpt-5-customer',
      body: {
        messages: [{ role: 'user', content: 'test prompt' }],
        max_tokens: 99_999,
        reasoning_effort: 'high',
        temperature: 0.2,
        top_p: 0.8,
      },
      byo: {
        provider: 'azure-openai',
        endpoint: 'https://customer-resource.services.ai.azure.com/',
        apiKey: 'customer-secret-key',
      },
    })),
  });

  assert.equal(response.status, 200);
  assert.equal(
    captured.url,
    'https://customer-resource.services.ai.azure.com/openai/v1/chat/completions',
  );
  const upstreamBody = JSON.parse(captured.init.body);
  assert.equal(upstreamBody.model, 'gpt-5-customer');
  assert.equal(upstreamBody.max_completion_tokens, 32768);
  assert.equal(upstreamBody.max_tokens, undefined);
  assert.equal(upstreamBody.reasoning_effort, 'high');
  assert.equal(upstreamBody.temperature, undefined);
  assert.equal(upstreamBody.top_p, undefined);
});

test('OpenAI proxy preserves an explicit non-reasoning chat configuration', async (t) => {
  let captured;
  const server = await startServer({
    allowByoAIEndpoints: true,
    fetchImpl: async (url, init) => {
      captured = { url, init };
      return jsonResponse(200, {
        model: 'gpt-5-legacy-alias',
        choices: [{ message: { content: '{}' } }],
        usage: {},
      });
    },
    logger: silentLogger,
  });
  t.after(server.close);

  const response = await fetch(`${server.baseUrl}/api/openai`, {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify(byoRequestBody({
      apiFormat: 'chat-completions',
      deployment: 'gpt-5-legacy-alias',
      body: {
        messages: [{ role: 'user', content: 'test prompt' }],
        max_tokens: 500,
        temperature: 0.2,
      },
    })),
  });

  assert.equal(response.status, 200);
  const upstreamBody = JSON.parse(captured.init.body);
  assert.equal(upstreamBody.model, 'gpt-5-legacy-alias');
  assert.equal(upstreamBody.max_tokens, 500);
  assert.equal(upstreamBody.max_completion_tokens, undefined);
  assert.equal(upstreamBody.reasoning_effort, undefined);
  assert.equal(upstreamBody.temperature, 0.2);
});

test('OpenAI proxy rejects punctuation-only BYO model names', async (t) => {
  const server = await startServer({
    allowByoAIEndpoints: true,
    fetchImpl: async () => jsonResponse(200, {}),
    logger: silentLogger,
  });
  t.after(server.close);

  const response = await fetch(`${server.baseUrl}/api/openai`, {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify(byoRequestBody({ deployment: '..' })),
  });

  assert.equal(response.status, 400);
  assert.equal((await response.json()).error.code, 'invalid_deployment_name');
});

test('OpenAI proxy returns a BYO-specific authentication error', async (t) => {
  const server = await startServer({
    allowByoAIEndpoints: true,
    fetchImpl: async () => jsonResponse(401, {
      error: {
        code: 'invalid_api_key',
        message: 'sensitive provider detail',
      },
    }),
    logger: silentLogger,
  });
  t.after(server.close);

  const response = await fetch(`${server.baseUrl}/api/openai`, {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify(byoRequestBody()),
  });

  assert.equal(response.status, 401);
  const payload = await response.json();
  assert.equal(payload.error.source, 'byo_azure_openai');
  assert.equal(payload.error.code, 'byo_authentication_failed');
  assert.doesNotMatch(JSON.stringify(payload), /sensitive provider detail|customer-secret-key/);
});

test('OpenAI proxy rejects deployments outside the server allowlist', async (t) => {
  const server = await startServer({
    endpoint: 'https://example.openai.azure.com/',
    apiKey: 'test-key',
    allowedDeployments: new Set(['approved']),
    logger: silentLogger,
  });
  t.after(server.close);

  const response = await fetch(`${server.baseUrl}/api/openai`, {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify(requestBody()),
  });

  assert.equal(response.status, 403);
  const payload = await response.json();
  assert.equal(payload.error.source, 'proxy');
  assert.equal(payload.error.code, 'deployment_not_allowed');
  assert.match(response.headers.get('x-azurediagarm-request-id'), /^[0-9a-f-]{36}$/);
});

test('OpenAI proxy reports managed credential acquisition failures separately', async (t) => {
  const server = await startServer({
    endpoint: 'https://example.openai.azure.com/',
    credential: { async getToken() { throw new Error('identity unavailable'); } },
    allowedDeployments: new Set(['gpt-5.6-sol']),
    logger: silentLogger,
  });
  t.after(server.close);

  const response = await fetch(`${server.baseUrl}/api/openai`, {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify(requestBody()),
  });

  assert.equal(response.status, 502);
  const payload = await response.json();
  assert.equal(payload.error.source, 'credential');
  assert.equal(payload.error.code, 'credential_acquisition_failed');
  assert.doesNotMatch(JSON.stringify(payload), /identity unavailable/);
});

test('OpenAI proxy preserves safe upstream auth diagnostics without relaying the body', async (t) => {
  const events = [];
  const server = await startServer({
    endpoint: 'https://example.openai.azure.com/',
    credential: { async getToken() { return { token: 'secret-token' }; } },
    allowedDeployments: new Set(['gpt-5.6-sol']),
    fetchImpl: async () => jsonResponse(403, {
      error: {
        code: 'PermissionDenied',
        message: 'sensitive upstream detail',
      },
    }, {
      'apim-request-id': 'upstream-request-123',
    }),
    logger: {
      info(message) { events.push(message); },
      error(message) { events.push(message); },
    },
  });
  t.after(server.close);

  const response = await fetch(`${server.baseUrl}/api/openai`, {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify(requestBody()),
  });

  assert.equal(response.status, 403);
  assert.equal(response.headers.get('x-upstream-request-id'), 'upstream-request-123');
  const payload = await response.json();
  assert.equal(payload.error.source, 'azure_openai');
  assert.equal(payload.error.code, 'azure_openai_authentication_failed');
  assert.equal(payload.error.upstreamCode, 'PermissionDenied');
  assert.equal(payload.error.upstreamRequestId, 'upstream-request-123');
  assert.doesNotMatch(JSON.stringify(payload), /sensitive upstream detail|secret-token|test prompt/);
  assert.doesNotMatch(events.join('\n'), /sensitive upstream detail|secret-token|test prompt/);
});

test('OpenAI proxy classifies throttling and service failures', async (t) => {
  const responses = [
    jsonResponse(429, { error: { code: 'RateLimitReached' } }, { 'Retry-After': '17' }),
    jsonResponse(503, { error: { code: 'ServiceUnavailable' } }),
    jsonResponse(504, { error: { code: 'Timeout' } }),
  ];
  const server = await startServer({
    endpoint: 'https://example.openai.azure.com/',
    apiKey: 'test-key',
    allowedDeployments: new Set(['gpt-5.6-sol']),
    fetchImpl: async () => responses.shift(),
    logger: silentLogger,
  });
  t.after(server.close);

  const expected = [
    [429, 'azure_openai_rate_limited'],
    [503, 'azure_openai_unavailable'],
    [504, 'azure_openai_timeout'],
  ];
  for (const [status, code] of expected) {
    const response = await fetch(`${server.baseUrl}/api/openai`, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify(requestBody()),
    });
    assert.equal(response.status, status);
    const payload = await response.json();
    assert.equal(payload.error.code, code);
    if (status === 429) assert.equal(response.headers.get('retry-after'), '17');
  }
});

test('OpenAI proxy preserves provider cooldowns supplied in milliseconds', async (t) => {
  const cases = [
    { headers: { 'retry-after-ms': '1500' }, expected: '2' },
    { headers: { 'x-ms-retry-after-ms': '250' }, expected: '1' },
    { headers: { 'Retry-After': '17', 'retry-after-ms': '1500' }, expected: '17' },
    { headers: { 'retry-after-ms': 'not-a-duration' }, expected: null },
    { headers: { 'retry-after-ms': '-1' }, expected: null },
  ];
  let index = 0;
  const server = await startServer({
    endpoint: 'https://example.openai.azure.com/', apiKey: 'test-key',
    allowedDeployments: new Set(['gpt-5.6-sol']), logger: silentLogger,
    fetchImpl: async () => jsonResponse(429, { error: { code: 'RateLimitReached' } }, cases[index++].headers),
  });
  t.after(server.close);
  for (const entry of cases) {
    const response = await fetch(`${server.baseUrl}/api/openai`, {
      method: 'POST', headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify(requestBody()),
    });
    assert.equal(response.status, 429);
    assert.equal(response.headers.get('retry-after'), entry.expected);
    assert.equal((await response.json()).error.code, 'azure_openai_rate_limited');
  }
});

test('OpenAI proxy rejects malformed bodies and misleading non-JSON success responses', async (t) => {
  const server = await startServer({
    endpoint: 'https://example.openai.azure.com/',
    apiKey: 'test-key',
    allowedDeployments: new Set(['gpt-5.6-sol']),
    fetchImpl: async () => new Response('not actually json', {
      status: 200,
      headers: { 'Content-Type': 'application/notjson' },
    }),
    logger: silentLogger,
  });
  t.after(server.close);

  let response = await fetch(`${server.baseUrl}/api/openai`, {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify(requestBody({ body: null })),
  });
  assert.equal(response.status, 400);
  assert.equal((await response.json()).error.code, 'missing_request_body');

  response = await fetch(`${server.baseUrl}/api/openai`, {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify(requestBody()),
  });
  assert.equal(response.status, 502);
  assert.equal((await response.json()).error.code, 'invalid_upstream_response');
});

test('OpenAI proxy enforces trusted model and token limits on successful requests', async (t) => {
  let captured;
  const server = await startServer({
    endpoint: 'https://example.openai.azure.com/',
    credential: { async getToken() { return { token: 'managed-token' }; } },
    allowedDeployments: new Set(['gpt-5.6-sol']),
    fetchImpl: async (url, init) => {
      captured = { url, init };
      return jsonResponse(200, {
        model: 'gpt-5.6-sol',
        output_text: '{"services":[]}',
        usage: {},
      });
    },
    logger: silentLogger,
  });
  t.after(server.close);

  const response = await fetch(`${server.baseUrl}/api/openai`, {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify(requestBody()),
  });

  assert.equal(response.status, 200);
  assert.equal(captured.url, 'https://example.openai.azure.com/openai/v1/responses');
  assert.equal(captured.init.headers.Authorization, 'Bearer managed-token');
  const upstreamBody = JSON.parse(captured.init.body);
  assert.equal(upstreamBody.model, 'gpt-5.6-sol');
  assert.equal(upstreamBody.store, false);
  assert.equal(upstreamBody.max_output_tokens, 32768);
});

test('budget denial and unavailable storage return actionable errors without dispatch', async (t) => {
  let dispatched = 0;
  const server = await startServer({
    endpoint: 'https://example.openai.azure.com/', apiKey: 'test',
    allowedDeployments: new Set(['gpt-5.6-sol']),
    budget: createBudgetManager({ store: new MemoryBudgetStore(), dailyTokens: 100 }),
    fetchImpl: async () => { dispatched++; return jsonResponse(200, {}); }, logger: silentLogger,
  });
  t.after(server.close);
  const response = await fetch(`${server.baseUrl}/api/openai`, {
    method: 'POST', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify(requestBody()),
  });
  assert.equal(response.status, 429);
  assert.ok(Number(response.headers.get('Retry-After')) > 0);
  assert.equal((await response.json()).error.code, 'ai_daily_budget_exceeded');
  assert.equal(dispatched, 0);
});

test('budget reservations reconcile success and release concurrency after all upstream outcomes', async (t) => {
  const manager = createBudgetManager({ store: new MemoryBudgetStore(), dailyTokens: 1_000_000 });
  const outcomes = [
    () => jsonResponse(200, { usage: { total_tokens: 12 } }),
    () => jsonResponse(429, { error: { code: 'RateLimitReached' } }),
    () => { throw new Error('Network down'); },
    () => ({ ok: true, status: 200, headers: new Headers({ 'content-type': 'application/json' }), text: async () => { throw new Error('Body reset'); } }),
  ];
  const server = await startServer({
    endpoint: 'https://example.openai.azure.com/', apiKey: 'test', budget: manager,
    allowedDeployments: new Set(['gpt-5.6-sol']),
    fetchImpl: async () => outcomes.shift()(), logger: silentLogger,
  });
  t.after(server.close);
  for (const expected of [200, 429, 502, 502]) {
    const response = await fetch(`${server.baseUrl}/api/openai`, {
      method: 'POST', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify(requestBody()),
    });
    assert.equal(response.status, expected);
    await response.text();
    assert.equal((await manager.status('local-development')).concurrentRequests, 0);
  }
  const status = await manager.status('local-development');
  assert.ok(status.usedTokens > 12, 'unknown transport outcomes retain reservations');
});

test('timeout and client cancellation abort upstream and release shared concurrency', async (t) => {
  const manager = createBudgetManager({ store: new MemoryBudgetStore(), dailyTokens: 1_000_000 });
  let observedAbort = 0;
  let markStarted;
  let started = new Promise(resolve => { markStarted = resolve; });
  const server = await startServer({
    endpoint: 'https://example.openai.azure.com/', apiKey: 'test', budget: manager, timeoutMs: 100,
    allowedDeployments: new Set(['gpt-5.6-sol']),
    fetchImpl: async (_url, init) => new Promise((_resolve, reject) => {
      markStarted();
      init.signal.addEventListener('abort', () => { observedAbort++; reject(init.signal.reason); }, { once: true });
    }), logger: silentLogger,
  });
  t.after(server.close);
  const init = { method: 'POST', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify(requestBody()) };
  const response = await fetch(`${server.baseUrl}/api/openai`, init);
  assert.equal(response.status, 504);
  assert.equal((await manager.status('local-development')).concurrentRequests, 0);
  started = new Promise(resolve => { markStarted = resolve; });
  const controller = new AbortController();
  const pending = fetch(`${server.baseUrl}/api/openai`, { ...init, signal: controller.signal }).catch(error => error);
  await started;
  controller.abort();
  await pending;
  for (let count = 0; count < 20 && (await manager.status('local-development')).concurrentRequests; count++) {
    await new Promise(resolve => setTimeout(resolve, 10));
  }
  assert.equal((await manager.status('local-development')).concurrentRequests, 0);
  assert.equal(observedAbort, 2);
});

test('public AI proxy requires authenticated identity even when invoked without nginx', async (t) => {
  const server = await startServer({
    endpoint: 'https://example.openai.azure.com/', apiKey: 'test', mode: 'public',
    allowedDeployments: new Set(['gpt-5.6-sol']),
    budget: createBudgetManager({ store: new MemoryBudgetStore() }), logger: silentLogger,
    fetchImpl: async () => { throw new Error('Must not dispatch'); },
  });
  t.after(server.close);
  const response = await fetch(`${server.baseUrl}/api/openai`, {
    method: 'POST', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify(requestBody()),
  });
  assert.equal(response.status, 401);
});

test('OpenAI proxy routes Anthropic Messages through Microsoft Foundry', async (t) => {
  let captured;
  let requestedScope;
  const server = await startServer({
    foundryEndpoint: 'https://example.services.ai.azure.com/',
    credential: {
      async getToken(scope) {
        requestedScope = scope;
        return { token: 'managed-token' };
      },
    },
    allowedFoundryDeployments: new Set(['claude-opus-5']),
    fetchImpl: async (url, init) => {
      captured = { url, init };
      return jsonResponse(200, {
        model: 'claude-opus-5',
        content: [{ type: 'text', text: '{"services":[]}' }],
        usage: { input_tokens: 10, output_tokens: 5 },
      });
    },
    logger: silentLogger,
  });
  t.after(server.close);

  const response = await fetch(`${server.baseUrl}/api/openai`, {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify(anthropicRequestBody()),
  });

  assert.equal(response.status, 200);
  assert.equal(captured.url, 'https://example.services.ai.azure.com/anthropic/v1/messages');
  assert.equal(requestedScope, 'https://ai.azure.com/.default');
  assert.equal(captured.init.headers['anthropic-version'], '2023-06-01');
  const upstreamBody = JSON.parse(captured.init.body);
  assert.equal(upstreamBody.model, 'claude-opus-5');
  assert.equal(upstreamBody.max_tokens, 32768);
  assert.deepEqual(upstreamBody.thinking, { type: 'adaptive' });
  assert.deepEqual(upstreamBody.output_config, { effort: 'low' });
  assert.equal(upstreamBody.stream, false);
});

test('OpenAI proxy keeps Foundry and Azure OpenAI allowlists separate', async (t) => {
  const server = await startServer({
    endpoint: 'https://example.openai.azure.com/',
    foundryEndpoint: 'https://example.services.ai.azure.com/',
    apiKey: 'openai-key',
    foundryApiKey: 'foundry-key',
    allowedDeployments: new Set(['claude-opus-5']),
    allowedFoundryDeployments: new Set(['approved-claude']),
    logger: silentLogger,
  });
  t.after(server.close);

  const response = await fetch(`${server.baseUrl}/api/openai`, {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify(anthropicRequestBody()),
  });

  assert.equal(response.status, 403);
  assert.equal((await response.json()).error.code, 'deployment_not_allowed');
});

test('OpenAI proxy fails closed when the Foundry allowlist is missing', async (t) => {
  const server = await startServer({
    foundryEndpoint: 'https://example.services.ai.azure.com/',
    foundryApiKey: 'foundry-key',
    logger: silentLogger,
  });
  t.after(server.close);

  const response = await fetch(`${server.baseUrl}/api/openai`, {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify(anthropicRequestBody()),
  });

  assert.equal(response.status, 503);
  assert.equal((await response.json()).error.code, 'deployment_allowlist_not_configured');
});

test('OpenAI proxy fails closed when the Azure OpenAI allowlist is missing', async (t) => {
  const server = await startServer({
    endpoint: 'https://example.openai.azure.com/',
    apiKey: 'test-key',
    // allowedDeployments omitted — defaults to empty Set
    logger: silentLogger,
  });
  t.after(server.close);

  const response = await fetch(`${server.baseUrl}/api/openai`, {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify(requestBody()),
  });

  assert.equal(response.status, 503);
  const payload = await response.json();
  assert.equal(payload.error.code, 'deployment_allowlist_not_configured');
});

test('OpenAI proxy enforces Azure OpenAI rate limit per client key', async (t) => {
  let callCount = 0;
  const server = await startServer({
    endpoint: 'https://example.openai.azure.com/',
    apiKey: 'test-key',
    allowedDeployments: new Set(['gpt-5.6-sol']),
    fetchImpl: async () => {
      callCount += 1;
      return jsonResponse(200, { model: 'gpt-5.6-sol', output_text: '{}', usage: {} });
    },
    consumeRateLimit: (() => {
      let calls = 0;
      return () => {
        calls += 1;
        return calls > 2 ? 60 : 0;
      };
    })(),
    logger: silentLogger,
  });
  t.after(server.close);

  const makeRequest = () => fetch(`${server.baseUrl}/api/openai`, {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify(requestBody()),
  });

  const r1 = await makeRequest();
  assert.equal(r1.status, 200);
  const r2 = await makeRequest();
  assert.equal(r2.status, 200);
  const r3 = await makeRequest();
  assert.equal(r3.status, 429);
  const payload = await r3.json();
  assert.equal(payload.error.code, 'proxy_rate_limit_exceeded');
  assert.equal(r3.headers.get('retry-after'), '60');
  // The rate-limited request must not reach the upstream
  assert.equal(callCount, 2);
});

test('shared rate limits reject before token reservation or upstream dispatch', async (t) => {
  let reservations = 0;
  let calls = 0;
  const events = [];
  const server = await startServer({
    endpoint: 'https://example.openai.azure.com/', apiKey: 'test',
    allowedDeployments: new Set(['gpt-5.6-sol']),
    logger: { warn(message) { events.push(message); } },
    consumeRateLimit: async () => 5,
    budget: { async reserve() { reservations++; } },
    fetchImpl: async () => { calls++; return jsonResponse(200, {}); },
  });
  t.after(server.close);
  const response = await fetch(`${server.baseUrl}/api/openai`, {
    method: 'POST', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify(requestBody()),
  });
  assert.equal(response.status, 429);
  assert.equal(response.headers.get('Retry-After'), '5');
  const payload = await response.json();
  assert.equal(payload.error.code, 'proxy_rate_limit_exceeded');
  const event = JSON.parse(events[0]?.replace('[openai-proxy] ', '') || '{}');
  assert.equal(event.event, 'proxy_rate_limit_exceeded');
  assert.equal(event.requestId, payload.error.requestId);
  assert.equal(event.status, 429);
  assert.equal(event.retryAfterSeconds, 5);
  assert.doesNotMatch(events.join('\n'), /test prompt|test-key|client-supplied-model/);
  assert.equal(reservations, 0);
  assert.equal(calls, 0);
});

test('provider failures followed by exhausted budget remain distinguishable and privately traceable', async (t) => {
  const events = [];
  let calls = 0;
  const budget = createBudgetManager({ store: new MemoryBudgetStore(), dailyTokens: 50_000 });
  const server = await startServer({
    endpoint: 'https://example.openai.azure.com/', apiKey: 'private-test-key',
    allowedDeployments: new Set(['gpt-5.6-sol']), budget,
    logger: {
      error(message) { events.push(message); },
      warn(message) { events.push(message); },
    },
    fetchImpl: async () => {
      calls++;
      return jsonResponse(500, { error: { code: 'server_error', message: 'Private provider detail' } });
    },
  });
  t.after(server.close);
  const request = () => fetch(`${server.baseUrl}/api/openai`, {
    method: 'POST', headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify(requestBody()),
  });
  const first = await request();
  assert.equal(first.status, 500);
  assert.equal((await first.json()).error.code, 'azure_openai_unavailable');
  const second = await request();
  assert.equal(second.status, 429);
  const payload = await second.json();
  assert.equal(payload.error.source, 'budget');
  assert.equal(payload.error.code, 'ai_daily_budget_exceeded');
  assert.equal(calls, 1, 'budget exhaustion must not dispatch another model request');
  const event = events.map(message => JSON.parse(message.replace('[openai-proxy] ', '')))
    .find(entry => entry.requestId === payload.error.requestId);
  assert.equal(event?.event, 'ai_daily_budget_exceeded');
  assert.equal(event?.status, 429);
  assert.equal(event?.retryAfterSeconds, Number(second.headers.get('retry-after')));
  assert.doesNotMatch(events.join('\n'), /test prompt|private-test-key|Private provider detail|local-development/);
  const balance = await budget.status('local-development');
  assert.ok(balance.usedTokens > 32_000, 'unknown upstream usage must not be silently refunded');
  assert.equal(balance.concurrentRequests, 0);
});

test('unavailable shared token storage fails closed instead of showing a full or empty budget', async (t) => {
  let calls = 0;
  const server = await startServer({
    endpoint: 'https://example.openai.azure.com/', apiKey: 'test',
    allowedDeployments: new Set(['gpt-5.6-sol']), logger: silentLogger,
    budget: createBudgetManager({ store: { async read() { throw new Error('Private storage detail'); } } }),
    fetchImpl: async () => { calls++; return jsonResponse(200, {}); },
  });
  t.after(server.close);
  const response = await fetch(`${server.baseUrl}/api/openai`, {
    method: 'POST', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify(requestBody()),
  });
  assert.equal(response.status, 503);
  const payload = await response.json();
  assert.equal(payload.error.code, 'ai_budget_unavailable');
  assert.doesNotMatch(JSON.stringify(payload), /Private storage detail/);
  assert.equal(calls, 0);
});

test('BYO reasoning and Foundry requests share the same budget without losing provider behavior', async (t) => {
  const budget = createBudgetManager({ store: new MemoryBudgetStore(), dailyTokens: 1_000_000 });
  const captured = [];
  const server = await startServer({
    foundryEndpoint: 'https://example.services.ai.azure.com/', foundryApiKey: 'test',
    allowedFoundryDeployments: new Set(['claude-opus-5']), allowByoAIEndpoints: true,
    logger: silentLogger, budget,
    fetchImpl: async (url, init) => {
      captured.push({ url, body: JSON.parse(init.body), redirect: init.redirect });
      return jsonResponse(200, { usage: { input_tokens: 10, output_tokens: 5, cache_read_input_tokens: 20 } });
    },
  });
  t.after(server.close);
  const requests = [
    byoRequestBody({
      apiFormat: 'chat-completions',
      body: { messages: [{ role: 'user', content: 'test' }], max_completion_tokens: 500.9, reasoning_effort: 'high', n: 100, stream: true },
    }),
    anthropicRequestBody(),
  ];
  for (const body of requests) {
    const response = await fetch(`${server.baseUrl}/api/openai`, {
      method: 'POST', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify(body),
    });
    assert.equal(response.status, 200);
    await response.text();
  }
  assert.equal(captured[0].body.max_completion_tokens, 500);
  assert.equal(captured[0].body.max_tokens, undefined);
  assert.equal(captured[0].body.n, 1);
  assert.equal(captured[0].body.store, false);
  assert.equal(captured[0].body.stream, false);
  assert.equal(captured[1].body.store, undefined);
  assert.equal(captured[1].body.stream, false);
  assert.equal(captured[1].body.max_tokens, 32768);
  assert.equal(captured.every(request => request.redirect === 'error'), true);
  assert.equal((await budget.status('local-development')).usedTokens, 70);
  assert.equal((await budget.status('local-development')).concurrentRequests, 0);
});

test('configured Astra dispatches the actual deployment while preserving explicit GPT-5.6 fallback', async (t) => {
  const captured = [];
  const budget = createBudgetManager({ store: new MemoryBudgetStore(), dailyTokens: 1_000_000 });
  const server = await startServer({
    endpoint: 'https://example.openai.azure.com/', apiKey: 'test',
    allowedDeployments: new Set(['gpt-6-astra', 'gpt-5.6-sol']), budget, logger: silentLogger,
    fetchImpl: async (url, init) => {
      captured.push({ url, body: JSON.parse(init.body) });
      return jsonResponse(200, { usage: { total_tokens: 10 } });
    },
  });
  t.after(server.close);
  for (const deployment of ['gpt-6-astra', 'gpt-5.6-sol']) {
    const response = await fetch(`${server.baseUrl}/api/openai`, {
      method: 'POST', headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify(requestBody({
        deployment,
        body: { model: 'client-supplied-model', input: 'test', max_output_tokens: 100 },
      })),
    });
    assert.equal(response.status, 200);
    await response.text();
  }
  assert.deepEqual(captured.map(request => request.body.model), ['gpt-6-astra', 'gpt-5.6-sol']);
  assert.equal(captured.every(request => request.url === 'https://example.openai.azure.com/openai/v1/responses'), true);
  assert.equal((await budget.status('local-development')).usedTokens, 20);
});

test('an unconfigured Astra deployment is rejected rather than aliased to GPT-5.6', async (t) => {
  let dispatched = false;
  const server = await startServer({
    endpoint: 'https://example.openai.azure.com/', apiKey: 'test',
    allowedDeployments: new Set(['gpt-5.6-sol']), logger: silentLogger,
    fetchImpl: async () => { dispatched = true; return jsonResponse(200, {}); },
  });
  t.after(server.close);
  const response = await fetch(`${server.baseUrl}/api/openai`, {
    method: 'POST', headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify(requestBody({ deployment: 'gpt-6-astra' })),
  });
  assert.equal(response.status, 403);
  assert.equal((await response.json()).error.code, 'deployment_not_allowed');
  assert.equal(dispatched, false);
});
