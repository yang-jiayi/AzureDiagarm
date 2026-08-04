// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

const assert = require('node:assert/strict');
const { test } = require('node:test');
const express = require('express');
const { createOpenAIProxyRouter, logFoundryConfiguration } = require('./openai-proxy');

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

test('OpenAI proxy rejects malformed bodies and non-JSON success responses', async (t) => {
  const server = await startServer({
    endpoint: 'https://example.openai.azure.com/',
    apiKey: 'test-key',
    allowedDeployments: new Set(['gpt-5.6-sol']),
    fetchImpl: async () => new Response('<html>unexpected</html>', {
      status: 200,
      headers: { 'Content-Type': 'text/html' },
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
