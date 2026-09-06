import test, { afterEach } from 'node:test';
import assert from 'node:assert/strict';
import {
  buildApiUrl,
  buildRequestBody,
  callAzureOpenAIProxy,
  createOpenAIProxyError,
  parseApiResponse,
} from '../src/services/apiHelper.ts';

const originalFetch = globalThis.fetch;

afterEach(() => {
  globalThis.fetch = originalFetch;
});

test('buildRequestBody converts text and image input to Anthropic Messages format', () => {
  const body = buildRequestBody({
    deployment: 'claude-opus-5',
    messages: [
      { role: 'system', content: 'Return JSON only.' },
      {
        role: 'user',
        content: [
          { type: 'input_text', text: 'Analyze this diagram.' },
          { type: 'input_image', image_url: 'data:image/png;base64,QUJD' },
        ],
      },
    ],
    maxTokens: 12000,
    apiFormat: 'anthropic-messages',
    isReasoning: true,
    reasoningEffort: 'high',
  });

  assert.equal(
    buildApiUrl('https://example.services.ai.azure.com', 'claude-opus-5', 'anthropic-messages'),
    'https://example.services.ai.azure.com/anthropic/v1/messages',
  );
  assert.equal(body.model, 'claude-opus-5');
  assert.equal(body.max_tokens, 12000);
  assert.deepEqual(body.thinking, { type: 'adaptive' });
  assert.deepEqual(body.output_config, { effort: 'high' });
  assert.deepEqual(body.system, [{ type: 'text', text: 'Return JSON only.' }]);
  assert.deepEqual(body.messages[0], {
    role: 'user',
    content: [
      { type: 'text', text: 'Analyze this diagram.' },
      {
        type: 'image',
        source: { type: 'base64', media_type: 'image/png', data: 'QUJD' },
      },
    ],
  });
});

test('parseApiResponse extracts Anthropic text and token usage', () => {
  const parsed = parseApiResponse({
    content: [
      { type: 'thinking', thinking: 'internal' },
      { type: 'text', text: '{"services":[]}' },
    ],
    usage: { input_tokens: 120, output_tokens: 30 },
  }, 'anthropic-messages');

  assert.deepEqual(parsed, {
    content: '{"services":[]}',
    promptTokens: 120,
    completionTokens: 30,
    totalTokens: 150,
  });
});

test('reasoning Chat Completions uses modern token and reasoning parameters', () => {
  const body = buildRequestBody({
    deployment: 'gpt-5',
    messages: [{ role: 'user', content: 'Return JSON.' }],
    maxTokens: 2048,
    apiFormat: 'chat-completions',
    isReasoning: true,
    reasoningEffort: 'high',
  });

  assert.equal(body.max_completion_tokens, 2048);
  assert.equal(body.reasoning_effort, 'high');
  assert.equal(body.max_tokens, undefined);
  assert.equal(body.temperature, undefined);
  assert.deepEqual(body.response_format, { type: 'json_object' });
  assert.equal(
    buildApiUrl(
      'https://example.services.ai.azure.com',
      'gpt-5-deployment',
      'chat-completions',
    ),
    'https://example.services.ai.azure.com/openai/v1/chat/completions',
  );
});

test('callAzureOpenAIProxy sends BYO credentials only in the server request body', async () => {
  let requestUrl = '';
  // Collected, not assigned: TypeScript's flow analysis cannot see the write
  // inside the fetch stub, so a `let` initialised to null narrows to `never`
  // and every assertion below it silently checks nothing.
  const sent: Array<Record<string, unknown>> = [];
  globalThis.fetch = (async (input, init) => {
    requestUrl = String(input);
    sent.push(JSON.parse(String(init?.body)) as Record<string, unknown>);
    return new Response(JSON.stringify({ output_text: '{"services":[]}' }), {
      status: 200,
      headers: { 'Content-Type': 'application/json' },
    });
  }) as typeof fetch;

  const result = await callAzureOpenAIProxy({
    apiFormat: 'responses',
    deployment: 'gpt-5',
    body: { model: 'gpt-5', input: [{ role: 'user', content: 'Hello' }] },
    byo: {
      provider: 'openai',
      endpoint: 'https://api.openai.com',
      apiKey: 'sk-test-secret-value',
    },
  });

  assert.equal(result.ok, true);
  assert.equal(requestUrl, '/api/openai');
  assert.deepEqual(sent[0]?.byo, {
    provider: 'openai',
    endpoint: 'https://api.openai.com',
    apiKey: 'sk-test-secret-value',
  });
});

test('callAzureOpenAIProxy rejects misleading non-JSON response media types', async () => {
  globalThis.fetch = (async () => new Response('{"output_text":"{}"}', {
    status: 200,
    headers: { 'Content-Type': 'application/notjson' },
  })) as typeof fetch;

  const result = await callAzureOpenAIProxy({
    apiFormat: 'responses',
    deployment: 'gpt-5',
    body: { model: 'gpt-5', input: [{ role: 'user', content: 'Hello' }] },
  });

  assert.equal(result.ok, false);
  assert.equal(result.error?.code, 'invalid_upstream_response');
});

test('BYO authentication errors produce custom-endpoint guidance', () => {
  const error = createOpenAIProxyError({
    ok: false,
    status: 401,
    data: null,
    error: {
      source: 'byo_ai',
      code: 'byo_authentication_failed',
    },
  });

  assert.equal(
    error.message,
    'The custom AI endpoint rejected the API key. Check the key and try again.',
  );
});

for (const [retryAfter, date, expectedMs] of [
  ['60', undefined, 60_000],
  ['0', undefined, 0],
  ['Sun, 06 Sep 2026 04:01:00 GMT', 'Sun, 06 Sep 2026 04:00:00 GMT', 60_000],
  ['-1', undefined, undefined],
  ['not-a-date', undefined, undefined],
] as const) {
  test(`proxy preserves validated Retry-After ${retryAfter} and request provenance`, async () => {
    globalThis.fetch = async () => new Response(JSON.stringify({
      error: { source: 'azure_openai', code: 'azure_openai_rate_limited', upstreamCode: '429' },
    }), {
      status: 429,
      headers: {
        'Content-Type': 'application/json',
        'Retry-After': retryAfter,
        'x-azurediagarm-request-id': 'rate-limit-request',
        'x-upstream-request-id': 'azure-request',
        ...(date ? { Date: date } : {}),
      },
    });
    const result = await callAzureOpenAIProxy({
      apiFormat: 'responses', deployment: 'gpt-6-astra', body: { model: 'gpt-6-astra' },
    });
    const error = createOpenAIProxyError(result);
    assert.equal(error.retryAfterMs, expectedMs);
    assert.equal(error.status, 429);
    assert.equal(error.code, 'azure_openai_rate_limited');
    assert.equal(error.requestId, 'rate-limit-request');
    assert.equal(error.upstreamRequestId, 'azure-request');
  });
}

test('an unstructured 429 retains wait guidance and diagnostics rather than a generic request failure', async () => {
  globalThis.fetch = async () => new Response(JSON.stringify({
    error: { code: '429', message: 'Requests exceed the current token rate limit.' },
  }), {
    status: 429,
    headers: {
      'Content-Type': 'application/json',
      'Retry-After': '60',
      'x-azurediagarm-request-id': '00000000-0000-4000-8000-000000000001',
    },
  });
  const error = createOpenAIProxyError(await callAzureOpenAIProxy({
    apiFormat: 'responses', deployment: 'gpt-6-astra', body: { model: 'gpt-6-astra' },
  }));
  assert.equal(error.code, '429');
  assert.equal(error.source, 'unknown');
  assert.equal(error.upstreamCode, '429');
  assert.equal(error.retryAfterMs, 60_000);
  assert.match(error.message, /rate.limit/i);
  assert.doesNotMatch(error.message, /request failed \(429\)/);
  assert.match(error.message, /00000000-0000-4000-8000-000000000001/);
});

test('millisecond retry headers retain the longest valid provider cooldown', async () => {
  globalThis.fetch = async () => new Response('{}', {
    status: 429,
    headers: {
      'Content-Type': 'application/json', 'Retry-After': '1',
      'retry-after-ms': '1250', 'x-ms-retry-after-ms': '1500',
    },
  });
  const result = await callAzureOpenAIProxy({
    apiFormat: 'responses', deployment: 'gpt-6-astra', body: {},
  });
  assert.equal(result.error?.retryAfterMs, 1500);
  assert.equal(createOpenAIProxyError(result).retryAfterMs, 1500);
});

for (const includeSource of [true, false]) {
  test(`daily-budget exhaustion retains its identity and UTC-reset guidance (source present: ${includeSource})`, async () => {
    globalThis.fetch = async () => new Response(JSON.stringify({
      error: {
        ...(includeSource ? { source: 'budget' } : {}),
        code: 'ai_daily_budget_exceeded', requestId: 'daily-budget-request',
      },
    }), { status: 429, headers: { 'Content-Type': 'application/json', 'Retry-After': '5' } });
    const error = createOpenAIProxyError(await callAzureOpenAIProxy({
      apiFormat: 'responses', deployment: 'gpt-6-astra', body: {},
    }));
    assert.equal(error.source, 'budget');
    assert.equal(error.code, 'ai_daily_budget_exceeded');
    assert.equal(error.requestId, 'daily-budget-request');
    assert.equal(error.retryAfterMs, 5000);
    assert.match(error.message, /daily AI budget.*midnight UTC/);
    assert.doesNotMatch(error.message, /provider.*rate.limit|deployment capacity|Wait a moment/);
  });
}

test('provider 500 diagnostics retain both normalized and original upstream codes', () => {
  const error = createOpenAIProxyError({
    ok: false, status: 500, data: null,
    error: {
      source: 'azure_openai', code: 'azure_openai_unavailable',
      requestId: 'proxy-500', upstreamRequestId: 'provider-500',
      upstreamStatus: 500, upstreamCode: 'server_error',
    },
  });
  assert.equal(error.status, 500);
  assert.equal(error.source, 'azure_openai');
  assert.equal(error.code, 'azure_openai_unavailable');
  assert.equal(error.upstreamStatus, 500);
  assert.equal(error.upstreamCode, 'server_error');
  assert.equal(error.requestId, 'proxy-500');
  assert.equal(error.upstreamRequestId, 'provider-500');
});
