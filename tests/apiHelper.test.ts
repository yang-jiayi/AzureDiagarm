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
