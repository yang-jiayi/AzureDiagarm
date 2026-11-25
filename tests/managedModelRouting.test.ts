import test from 'node:test';
import assert from 'node:assert/strict';
import { jsonResponse, loadBYOClient } from './fixtures/byoClientHarness';

const env = {
  VITE_AZURE_OPENAI_ENDPOINT: 'https://offline.openai.azure.com',
  VITE_AZURE_OPENAI_DEPLOYMENT_GPT62ASTRA: 'astra-production-alias',
  VITE_AZURE_OPENAI_DEPLOYMENT_GPT62SOL: 'sol-production-alias',
};
test('both managed models reject unsupported direct efforts before dispatch rather than silently lowering them', async t => {
  let calls = 0;
  t.mock.method(globalThis, 'fetch', async () => { calls++; return jsonResponse({}); });
  const { client } = await loadBYOClient(undefined, env);
  for (const effort of ['none', 'minimal'] as const) {
    const body = client.buildRequestBody({
      apiFormat: 'responses', deployment: env.VITE_AZURE_OPENAI_DEPLOYMENT_GPT62SOL,
      messages: [{ role: 'user', content: 'Offline request' }], maxTokens: 32000,
      isReasoning: true, reasoningEffort: effort,
    });
    await assert.rejects(client.callAzureOpenAIProxy({
      apiFormat: 'responses', deployment: body.model, body,
    }), { code: 'unsupported_reasoning_effort' });
    assert.equal(body.reasoning.effort, effort);
  }
  assert.equal(calls, 0);
});

for (const model of ['gpt-6.2-astra', 'gpt-6.2-sol'] as const) {
  test(`${model} preserves JSON, vision, metadata and selected limits through every AI feature`, async t => {
    t.mock.method(console, 'log', () => {});
    t.mock.method(console, 'warn', () => {});
    const sent: any[] = [];
    let content = '';
    t.mock.method(globalThis, 'fetch', async (url: unknown, options?: RequestInit) => {
      if (url === '/api/docs-search') return jsonResponse({ results: [] });
      if (url === '/api/ai/budget') return jsonResponse({ available: true, concurrentLimit: 4, concurrentRequests: 0 });
      assert.equal(url, '/api/openai');
      const request = JSON.parse(String(options?.body));
      sent.push(request);
      return jsonResponse({
        status: 'completed', output_text: content, usage: { input_tokens: 10, output_tokens: 20, total_tokens: 30 },
      });
    });
    const { client } = await loadBYOClient(undefined, env);
    client.updateModelSettings({ model, reasoningEffort: 'max' });
    const top = client.captureRuntimeModelOverride('architectureGeneration');
    const blueprint = client.captureRuntimeModelOverride('blueprint');
    const validation = client.captureRuntimeModelOverride('validation');
    const guide = client.captureRuntimeModelOverride('deploymentGuide');
    content = JSON.stringify({ groups: [], services: [{ id: 'app', name: 'App Service', type: 'App Service', category: 'app services' }], connections: [] });
    const topology = await client.generateArchitectureWithAI('Offline architecture', top);
    await client.generateArchitectureFromIaC({ format: 'arm', content: { resources: [] }, filenames: ['offline.json'] }, 'en', top);
    content = JSON.stringify({ title: 'Offline', stages: [{ id: 'stage', label: 'Application', services: [{ id: 'app', name: 'App Service', category: 'app services', role: 'Application' }] }], connections: [] });
    await client.generateReferenceArchitectureWithAI('Offline reference', top);
    content = JSON.stringify({
      title: 'Offline', zones: [{ id: 'zone', label: 'Application', x: 0, y: 0, width: 640, height: 480 }],
      nodes: [{ id: 'app', name: 'App Service', kind: 'service', zone: 'zone', x: 1, y: 1 }], edges: [],
    });
    await client.generateBlueprintArchitectureWithAI('Offline blueprint', blueprint);
    content = JSON.stringify({ title: 'Offline', zones: [{ id: 'zone', label: 'Application' }], components: [{ id: 'app', name: 'App Service', role: 'Application', zone: 'zone' }], connections: [] });
    await client.generateComponentManifest('Offline manifest', top);
    content = JSON.stringify({ overallScore: 80, summary: 'Offline review', pillars: [], quickWins: [] });
    const review = await client.validateArchitecture([{ name: 'App Service', type: 'App Service', category: 'app services' }], [], undefined, undefined, validation);
    content = JSON.stringify({ title: 'Offline guide', deploymentSteps: [{ step: 1, title: 'Deploy', description: 'Offline step' }] });
    const deploymentGuide = await client.generateDeploymentGuide([], [], undefined, undefined, undefined, 'en', { modelOverride: guide });
    content = 'An exact image description.';
    const vision = await client.analyzeArchitectureDiagramImage('QUJD', 'image/png', 'en', { modelOverride: top });
    const imageRequest = sent.at(-1);
    assert.equal(imageRequest.body.input[1].content[1].type, 'input_image');
    content = '{"suggestions":["Enable private endpoints"]}';
    await client.generateFollowUpSuggestions({ services: ['App Service'], lastChange: 'Added app', recentRequests: [], modelOverride: top });
    assert.equal(sent.length, 9);
    for (const request of sent) {
      assert.equal(request.deployment, model === 'gpt-6.2-astra' ? env.VITE_AZURE_OPENAI_DEPLOYMENT_GPT62ASTRA : env.VITE_AZURE_OPENAI_DEPLOYMENT_GPT62SOL);
      assert.equal(request.body.model, request.deployment);
      assert.equal(request.apiFormat, 'responses');
      assert.equal(request.byo, undefined);
      assert.equal(request.body.max_output_tokens, 64000);
      assert.equal(request.body.reasoning.effort, 'max');
    }
    for (const metrics of [topology.metrics, review.metrics, deploymentGuide.metrics, vision.metrics]) {
      assert.equal(metrics?.source, 'managed');
      assert.equal(metrics?.deployment, sent[0].deployment);
      assert.equal(metrics?.reasoningEffort, 'max');
      assert.equal(metrics?.totalTokens, 30);
    }
  });
}

test('managed max budget follows the captured effort, including explicit per-request overrides', async () => {
  const { client } = await loadBYOClient(undefined, env);
  client.updateModelSettings({ model: 'gpt-6.2-sol', reasoningEffort: 'low' });
  const max = client.captureRuntimeModelOverride('architectureGeneration', {
    model: 'gpt-6.2-astra', reasoningEffort: 'max',
  });
  assert.equal(max.connection?.maxCompletionTokens, 64000);
  assert.equal(max.connection?.reasoningEffort, 'max');
  assert.equal(client.resolveAIModelRuntime('architectureGeneration', max).connection, max.connection);
  const low = client.captureRuntimeModelOverride('architectureGeneration', {
    model: 'gpt-6.2-astra', reasoningEffort: 'low',
  });
  assert.equal(low.connection?.maxCompletionTokens, 32000);
  assert.equal(client.getEffectiveAIModelInfo('architectureGeneration').maxCompletionTokens, 32000);
  client.updateModelSettings({ reasoningEffort: 'max' });
  assert.equal(client.getEffectiveAIModelInfo('architectureGeneration').maxCompletionTokens, 64000);
});

test('managed client rejects unsupported provider-specific formats', async () => {
  const { client } = await loadBYOClient(undefined, env);
  assert.throws(() => client.parseApiResponse({}, 'anthropic-messages' as never), { code: 'unsupported_api_format' });
  assert.throws(() => client.buildRequestBody({
    apiFormat: 'anthropic-messages' as never, deployment: 'opus', messages: [{ role: 'user', content: 'hello' }],
    reasoningEffort: 'high', isReasoning: true, maxTokens: 32000,
  }), { code: 'unsupported_api_format' });
});
