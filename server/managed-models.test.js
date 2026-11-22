// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.
const test = require('node:test');
const assert = require('node:assert/strict');
const { createOpenAIExecutor } = require('./openai-proxy');
const { managedModelsConfiguration, runtimeManagedModelsConfiguration } = require('./astra-policy');
const { createBudgetManager, MemoryBudgetStore } = require('./ai-budget');

const env = {
  AZURE_OPENAI_ENDPOINT: 'https://models.openai.azure.com',
  AZURE_OPENAI_DEPLOYMENT_GPT62ASTRA: 'astra-alias',
  AZURE_OPENAI_DEPLOYMENT_GPT62SOL: 'sol-alias',
  AZURE_OPENAI_ALLOWED_DEPLOYMENTS: 'astra-alias,sol-alias',
};
const baseRequest = (deployment, effort = 'low') => ({
  apiFormat: 'responses',
  deployment,
  body: {
    model: deployment, input: 'Harmless synthetic request',
    reasoning: { effort }, max_output_tokens: 32, store: false,
  },
});
function executor(overrides = {}) {
  return createOpenAIExecutor({
    endpoint: env.AZURE_OPENAI_ENDPOINT,
    astraDeployment: 'astra-alias',
    solDeployment: 'sol-alias',
    allowedDeployments: new Set(['astra-alias', 'sol-alias']),
    logger: { info() {}, warn() {}, error() {} },
    ...overrides,
  });
}

test('model policy exposes exactly GPT-6.2 Astra and Sol with Sol as the default', () => {
  const config = managedModelsConfiguration(env);
  assert.equal(config.configured, true);
  assert.deepEqual([config.astra.model, config.sol.model], ['gpt-6.2-astra', 'gpt-6.2-sol']);
  const runtime = runtimeManagedModelsConfiguration(config);
  assert.equal(runtime.ai.models.length, 2);
  assert.equal(runtime.ai.defaultModel, 'gpt-6.2-sol');
  for (const change of [
    { AZURE_OPENAI_DEPLOYMENT_GPT6ASTRA: 'legacy-astra' },
    { AZURE_OPENAI_DEPLOYMENT_GPT61SOL: 'legacy-sol' },
    { AZURE_FOUNDRY_DEPLOYMENT_OPUS55: 'legacy-opus' },
    { AZURE_OPENAI_ALLOWED_DEPLOYMENTS: 'astra-alias' },
    { AZURE_OPENAI_DEPLOYMENT_GPT62SOL: 'astra-alias' },
  ]) assert.throws(() => managedModelsConfiguration({ ...env, ...change }));
});

for (const deployment of ['astra-alias', 'sol-alias']) {
  test(`${deployment} uses the isolated managed Responses route and Entra scope`, async () => {
    const budget = createBudgetManager({ store: new MemoryBudgetStore(), dailyTokens: 1_000_000 });
    let captured;
    const scopes = [];
    const run = executor({
      budget,
      credential: { async getToken(scope) { scopes.push(scope); return { token: 'managed-token' }; } },
      fetchImpl: async (url, init) => {
        captured = { url, init, body: JSON.parse(init.body) };
        return new Response(JSON.stringify({
          status: 'completed', output_text: 'OK', usage: { input_tokens: 10, output_tokens: 2 },
        }), { headers: { 'Content-Type': 'application/json' } });
      },
    });
    const result = await run({ body: baseRequest(deployment) });
    assert.equal(result.status, 200);
    assert.equal(captured.url, `https://models.openai.azure.com/openai/v1/responses`);
    assert.deepEqual(scopes, ['https://cognitiveservices.azure.com/.default']);
    assert.equal(captured.init.headers.Authorization, 'Bearer managed-token');
    assert.equal(captured.init.headers['api-key'], undefined);
    assert.equal(captured.body.model, deployment);
    assert.equal(captured.body.store, false);
    assert.equal(captured.body.reasoning.effort, 'low');
    assert.equal(captured.body.max_output_tokens, 32);
    assert.equal((await budget.status('local-development')).usedTokens, 12);
  });
}

test('unsupported managed reasoning is rejected before credentials, budgets or dispatch; BYO semantics stay independent', async () => {
  let touched = false;
  const run = executor({
    allowByoAIEndpoints: true,
    credential: { getToken() { touched = true; } },
    budget: { reserve() { touched = true; } },
    fetchImpl() { touched = true; },
  });
  for (const deployment of ['astra-alias', 'sol-alias']) {
    for (const effort of ['none', 'minimal', 'unexpected']) {
      const result = await run({ body: baseRequest(deployment, effort) });
      assert.equal(result.status, 400);
      assert.equal(result.body.error.code, 'invalid_api_format');
      assert.equal(touched, false);
    }
  }
  let sent;
  const byo = executor({
    allowByoAIEndpoints: true,
    fetchImpl: async (_url, init) => {
      sent = JSON.parse(init.body);
      return new Response(JSON.stringify({ status: 'completed', output_text: 'OK' }), {
        headers: { 'Content-Type': 'application/json' },
      });
    },
  });
  const result = await byo({ body: {
    apiFormat: 'responses', deployment: 'customer-model',
    byo: { provider: 'openai', endpoint: 'https://api.openai.com', apiKey: 'offline-only-key' },
    body: { model: 'customer-model', input: 'Offline request', reasoning: { effort: 'none' }, max_output_tokens: 32 },
  } });
  assert.equal(result.status, 200);
  assert.equal(sent.reasoning.effort, 'none');
});
