// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

const assert = require('node:assert/strict');
const { test } = require('node:test');
const express = require('express');
const { createOpenAIProxyRouter } = require('./openai-proxy');

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

function jsonResponse(status, body, headers = {}) {
  return new Response(JSON.stringify(body), {
    status,
    headers: { 'Content-Type': 'application/json', ...headers },
  });
}

const silentLogger = { info() {}, error() {} };

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
