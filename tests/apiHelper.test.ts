import test, { afterEach } from 'node:test';
import assert from 'node:assert/strict';
import { build } from 'esbuild';

const bundled = await build({
  stdin: {
    contents: `export * from './src/services/apiHelper';`,
    resolveDir: process.cwd(), loader: 'ts',
  },
  bundle: true, write: false, platform: 'node', format: 'cjs', logLevel: 'silent',
  define: { 'import.meta.env': JSON.stringify({
    VITE_AZURE_OPENAI_ENDPOINT: 'https://offline.openai.azure.com/',
    VITE_AZURE_OPENAI_DEPLOYMENT_GPT6ASTRA: 'gpt-6-astra',
  }) },
});
const module = { exports: {} };
new Function('module', 'exports', bundled.outputFiles[0].text)(module, module.exports);
const {
  buildRequestBody,
  callAzureOpenAIProxy,
  createOpenAIProxyError,
  parseApiResponse,
} = module.exports as typeof import('../src/services/apiHelper');

const originalFetch = globalThis.fetch;

afterEach(() => {
  globalThis.fetch = originalFetch;
});

test('parseApiResponse extracts Responses output text and token usage', () => {
  const parsed = parseApiResponse({
    output: [
      { type: 'reasoning', summary: [] },
      { type: 'message', content: [{ type: 'output_text', text: '{"services":[]}' }] },
    ],
    usage: { input_tokens: 120, output_tokens: 30, total_tokens: 150 },
  }, 'responses');

  assert.deepEqual(parsed, {
    content: '{"services":[]}',
    promptTokens: 120,
    completionTokens: 30,
    totalTokens: 150,
  });
});

test('format builders and parsers reject unsupported non-OpenAI formats', () => {
  for (const apiFormat of ['anthropic-messages', 'legacy-responses', '']) {
    assert.throws(() => Reflect.apply(buildRequestBody, undefined, [{
      deployment: 'gpt-6-astra', messages: [{ role: 'user', content: 'Offline request' }],
      maxTokens: 32000, apiFormat, isReasoning: true, reasoningEffort: 'max',
    }]), { name: 'AIModelConfigurationError', code: 'unsupported_api_format' });
    assert.throws(() => Reflect.apply(parseApiResponse, undefined, [{}, apiFormat]),
      { name: 'AIModelConfigurationError', code: 'unsupported_api_format' });
  }
});

test('Responses-only vision preserves image data, detail, MAX reasoning and 32K without mutation', () => {
  const imageUrl = 'data:image/png;base64,QUJD';
  const messages = [
    { role: 'system', content: 'Describe every label.' },
    { role: 'user', content: [
      { type: 'input_text', text: 'Read this image exactly.' },
      { type: 'input_image', image_url: imageUrl, detail: 'high' },
    ] },
  ];
  const original = structuredClone(messages);
  const params = {
    deployment: 'gpt-6-astra', messages, maxTokens: 32000,
    isReasoning: true, reasoningEffort: 'max' as const, jsonOutput: false,
  };
  const responses = buildRequestBody({ ...params, apiFormat: 'responses' });
  assert.deepEqual(responses.input, original);
  assert.equal(responses.model, 'gpt-6-astra');
  assert.equal(responses.max_output_tokens, 32000);
  assert.deepEqual(responses.reasoning, { effort: 'max' });
  assert.equal(responses.store, false);
  assert.equal(responses.text, undefined);
  assert.deepEqual(messages, original);
});

test('callAzureOpenAIProxy sends only the managed Astra request envelope', async () => {
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
    deployment: 'gpt-6-astra',
    body: { model: 'gpt-6-astra', input: [{ role: 'user', content: 'Hello' }] },
  });

  assert.equal(result.ok, true);
  assert.equal(requestUrl, '/api/openai');
  assert.deepEqual(sent[0], {
    apiFormat: 'responses', deployment: 'gpt-6-astra',
    body: { model: 'gpt-6-astra', input: [{ role: 'user', content: 'Hello' }] },
  });
});

test('proxy boundary rejects alternate providers, formats and deployments without HTTP', async () => {
  let calls = 0;
  globalThis.fetch = async () => { calls++; throw new Error('Unexpected HTTP'); };
  const valid = { apiFormat: 'responses' as const, deployment: 'gpt-6-astra', body: { model: 'gpt-6-astra' } };
  for (const invalid of [
    { ...valid, apiFormat: 'chat-completions' },
    { ...valid, apiFormat: 'anthropic-messages' },
    { ...valid, deployment: 'gpt-5.6-terra' },
    { ...valid, body: { model: 'gpt-5.6-sol' } },
    { ...valid, body: {} },
    { ...valid, byo: { provider: 'openai', endpoint: 'https://api.openai.com', apiKey: 'synthetic-key' } },
    { ...valid, byo: null },
    { ...valid, byo: undefined },
  ]) {
    await assert.rejects(Reflect.apply(callAzureOpenAIProxy, undefined, [invalid]), { name: 'AIModelConfigurationError' });
  }
  assert.equal(calls, 0);
});

test('callAzureOpenAIProxy rejects misleading non-JSON response media types', async () => {
  globalThis.fetch = (async () => new Response('{"output_text":"{}"}', {
    status: 200,
    headers: { 'Content-Type': 'application/notjson' },
  })) as typeof fetch;

  const result = await callAzureOpenAIProxy({
    apiFormat: 'responses',
    deployment: 'gpt-6-astra',
    body: { model: 'gpt-6-astra', input: [{ role: 'user', content: 'Hello' }] },
  });

  assert.equal(result.ok, false);
  assert.equal(result.error?.code, 'invalid_upstream_response');
});

test('proxy policy codes produce stable Astra guidance without echoing arbitrary raw messages', async () => {
  for (const [status, code, message] of [
    [400, 'invalid_api_format', 'GPT-6 Astra requests must use the Responses API.'],
    [403, 'byo_not_enabled', 'Bring-your-own AI is disabled by the application administrator.'],
    [503, 'astra_not_configured', 'GPT-6 Astra is not configured. Contact the application administrator to configure the managed Astra deployment.'],
    [403, 'deployment_not_allowed', 'Only the configured GPT-6 Astra deployment can run.'],
    [503, 'proxy_not_configured', 'The managed Azure OpenAI endpoint is not configured correctly.'],
  ] as const) {
    globalThis.fetch = async () => new Response(JSON.stringify({
      error: { source: 'proxy', code, message: 'Arbitrary raw diagnostic text', requestId: 'policy-request' },
    }), { status, headers: { 'Content-Type': 'application/json' } });
    const result = await callAzureOpenAIProxy({
      apiFormat: 'responses', deployment: 'gpt-6-astra', body: { model: 'gpt-6-astra' },
    });
    const error = createOpenAIProxyError(result);
    assert.equal(error.message, `${message} Request ID: policy-request`, code);
    assert.equal(error.source, 'proxy');
    assert.equal(error.code, code);
    assert.equal(error.status, status);
    assert.equal(error.requestId, 'policy-request');
    assert.equal(result.error?.message, message);
    assert.doesNotMatch(JSON.stringify(result), /Arbitrary raw diagnostic text/);
  }
});

test('unconfigured server Astra returns administrator guidance with configuration provenance', async () => {
  globalThis.fetch = async () => new Response(JSON.stringify({
    error: { source: 'configuration', code: 'astra_not_configured', requestId: 'astra-setup' },
  }), { status: 503, headers: { 'Content-Type': 'application/json' } });
  const result = await callAzureOpenAIProxy({
    apiFormat: 'responses', deployment: 'gpt-6-astra', body: { model: 'gpt-6-astra' },
  });
  const error = createOpenAIProxyError(result);
  assert.equal(error.message, 'GPT-6 Astra is not configured. Contact the application administrator to configure the managed Astra deployment. Request ID: astra-setup');
  assert.equal(error.code, 'astra_not_configured');
  assert.equal(error.source, 'configuration');
  assert.equal(error.status, 503);
});

test('BYO provider errors retain diagnostics and actionable key-retest guidance', () => {
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
    'The AI provider rejected this profile’s API key. Re-enter the key and test the connection again.',
  );
  assert.equal(error.code, 'byo_authentication_failed');
  assert.equal(error.source, 'byo_ai');
});

test('frozen backend BYO validation codes retain safe guidance and exact error provenance', async () => {
  for (const [code, status, expected] of [
    ['invalid_byo_configuration', 400, 'The AI connection settings are invalid. Check the provider, endpoint, and model, then test again.'],
    ['invalid_byo_provider', 400, 'The AI connection settings are invalid. Check the provider, endpoint, and model, then test again.'],
    ['invalid_byo_endpoint', 400, 'The AI connection settings are invalid. Check the provider, endpoint, and model, then test again.'],
    ['invalid_byo_api_key', 400, 'The AI provider rejected this profile’s API key. Re-enter the key and test the connection again.'],
    ['invalid_deployment_name', 400, 'Model or deployment not found. Check the configured name.'],
    ['byo_request_failed', 422, 'AI provider request failed (422). Please try again.'],
  ] as const) {
    globalThis.fetch = async () => new Response(JSON.stringify({
      error: { source: 'byo_ai', code, requestId: 'byo-policy-request', message: 'raw-private-provider-detail' },
    }), { status, headers: { 'Content-Type': 'application/json' } });
    const result = await callAzureOpenAIProxy({
      apiFormat: 'responses', deployment: 'gpt-6-astra', body: { model: 'gpt-6-astra' },
    });
    const error = createOpenAIProxyError(result);
    assert.equal(error.code, code);
    assert.equal(error.source, 'byo_ai');
    assert.equal(error.status, status);
    assert.equal(error.requestId, 'byo-policy-request');
    assert.equal(error.message, `${expected} Request ID: byo-policy-request`);
    assert.doesNotMatch(JSON.stringify(result), /raw-private-provider-detail/);
  }
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
    apiFormat: 'responses', deployment: 'gpt-6-astra', body: { model: 'gpt-6-astra' },
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
      apiFormat: 'responses', deployment: 'gpt-6-astra', body: { model: 'gpt-6-astra' },
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
