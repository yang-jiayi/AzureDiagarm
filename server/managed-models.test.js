// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.
const test = require('node:test');
const assert = require('node:assert/strict');
const { createOpenAIExecutor } = require('./openai-proxy');
const { managedModelsConfiguration, runtimeManagedModelsConfiguration } = require('./astra-policy');
const { createBudgetManager, MemoryBudgetStore } = require('./ai-budget');

const env = {
  AZURE_OPENAI_ENDPOINT: 'https://sol.openai.azure.com',
  AZURE_OPENAI_DEPLOYMENT_GPT61SOL: 'sol-alias',
  AZURE_OPENAI_ALLOWED_DEPLOYMENTS: 'sol-alias',
  AZURE_FOUNDRY_ENDPOINT: 'https://opus.services.ai.azure.com',
  AZURE_FOUNDRY_DEPLOYMENT_OPUS55: 'opus-alias',
  AZURE_FOUNDRY_ALLOWED_DEPLOYMENTS: 'opus-alias',
};
const envelope = () => ({
  apiFormat: 'anthropic-messages', deployment: 'opus-alias',
  body: {
    model: 'opus-alias', system: [{ type: 'text', text: 'Design an architecture' }],
    messages: [{ role: 'user', content: [{ type: 'text', text: 'JSON please' }, {
      type: 'image', source: { type: 'base64', media_type: 'image/png', data: 'QUJD' },
    }] }],
    max_tokens: 32000, thinking: { type: 'adaptive' }, output_config: { effort: 'max' },
  },
});
function executor(overrides = {}) {
  return createOpenAIExecutor({
    endpoint: env.AZURE_OPENAI_ENDPOINT, solDeployment: 'sol-alias', allowedDeployments: new Set(['sol-alias']),
    foundryEndpoint: env.AZURE_FOUNDRY_ENDPOINT, opusDeployment: 'opus-alias', allowedFoundryDeployments: new Set(['opus-alias']),
    logger: { info() {}, warn() {}, error() {} },
    ...overrides,
  });
}
test('model policy requires independent exact routes, never adopts Astra or legacy Claude', () => {
  assert.equal(managedModelsConfiguration(env).configured, true);
  const runtime = runtimeManagedModelsConfiguration(managedModelsConfiguration(env));
  assert.equal(runtime.ai.models.length, 2);
  assert.equal(runtime.ai.defaultModel, 'gpt-6.1-sol');
  for (const change of [
    { AZURE_OPENAI_DEPLOYMENT_GPT6ASTRA: 'old-astra' },
    { AZURE_FOUNDRY_API_KEY: 'do-not-use' },
    { AZURE_FOUNDRY_ENDPOINT: 'https://other.example.com' },
    { AZURE_FOUNDRY_ALLOWED_DEPLOYMENTS: 'opus-alias,old-opus' },
    { AZURE_FOUNDRY_DEPLOYMENT_OPUS55: 'sol-alias', AZURE_FOUNDRY_ALLOWED_DEPLOYMENTS: 'sol-alias' },
  ]) assert.throws(() => managedModelsConfiguration({ ...env, ...change }));
});
test('Claude dispatch is keyless, isolated, bounded and charges cached usage to the common budget', async () => {
  const budget = createBudgetManager({ store: new MemoryBudgetStore(), dailyTokens: 1_000_000 });
  let captured;
  const scopes = [];
  const run = executor({
    budget, apiKey: 'sol-only-secret',
    credential: { async getToken(scope) { scopes.push(scope); return { token: 'opus-token' }; } },
    fetchImpl: async (url, init) => {
      captured = { url, init, body: JSON.parse(init.body) };
      return new Response(JSON.stringify({
        type: 'message', stop_reason: 'end_turn', content: [{ type: 'thinking', thinking: 'not application text' }, { type: 'text', text: '{}' }],
        usage: { input_tokens: 10, output_tokens: 20, cache_creation_input_tokens: 30, cache_read_input_tokens: 40 },
      }), { headers: { 'Content-Type': 'application/json' } });
    },
  });
  const request = envelope();
  request.body.max_tokens = 99999;
  const result = await run({ body: request });
  assert.equal(result.status, 200);
  assert.equal(captured.url, 'https://opus.services.ai.azure.com/anthropic/v1/messages');
  assert.deepEqual(scopes, ['https://ai.azure.com/.default']);
  assert.equal(captured.init.headers.Authorization, 'Bearer opus-token');
  assert.equal(captured.init.headers['anthropic-version'], '2026-10-21');
  assert.equal(captured.init.headers['api-key'], undefined);
  assert.equal(captured.body.store, undefined);
  assert.equal(captured.body.stream, false);
  assert.equal(captured.body.max_tokens, 32768);
  assert.deepEqual(captured.body.messages, request.body.messages);
  assert.equal((await budget.status('local-development')).usedTokens, 100);
});
for (const [label, change] of [
  ['Sol format with Claude alias', value => { value.apiFormat = 'responses'; }],
  ['Claude format with Sol alias', value => { value.deployment = value.body.model = 'sol-alias'; }],
  ['retired Astra alias', value => { value.deployment = value.body.model = 'gpt-6-astra'; }],
  ['remote image', value => { value.body.messages[0].content[1].source = { type: 'url', url: 'https://example.com/image.png' }; }],
  ['disabled thinking', value => { value.body.thinking = { type: 'disabled' }; }],
  ['manual thinking', value => { value.body.thinking = { type: 'enabled', budget_tokens: 100 }; }],
  ['invalid effort', value => { value.body.output_config.effort = 'none'; }],
  ['sampling controls', value => { value.body.temperature = 1; }],
  ['assistant prefill', value => { value.body.messages[0].role = 'assistant'; }],
  ['tools', value => { value.body.tools = [{ name: 'remote' }]; }],
  ['endpoint injection', value => { value.body.endpoint = 'https://attacker.example'; }],
]) {
  test(`Claude rejects ${label} before credentials, budgets or dispatch`, async () => {
    let touched = false;
    const request = envelope();
    change(request);
    const run = executor({
      credential: { getToken() { touched = true; } },
      budget: { reserve() { touched = true; } },
      fetchImpl() { touched = true; },
    });
    const result = await run({ body: request });
    assert.ok([400, 403].includes(result.status));
    assert.equal(touched, false);
  });
}
test('Claude cancellation never dispatches again or refunds unknown usage', async () => {
  let calls = 0;
  const budget = createBudgetManager({ store: new MemoryBudgetStore(), dailyTokens: 1_000_000 });
  const controller = new AbortController();
  const run = executor({
    budget, credential: { async getToken() { return { token: 'token' }; } },
    fetchImpl: async (_url, init) => {
      calls++;
      controller.abort(Object.assign(new Error('Cancelled'), { code: 'ai_job_cancelled' }));
      throw init.signal.reason;
    },
  });
  const result = await run({ body: envelope() }, { signal: controller.signal });
  assert.equal(result.status, 409);
  assert.equal(calls, 1);
  const status = await budget.status('local-development');
  assert.ok(status.usedTokens > 32000);
  assert.equal(status.concurrentRequests, 0);
});
