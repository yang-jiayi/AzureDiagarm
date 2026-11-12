import test from 'node:test';
import assert from 'node:assert/strict';
import { jsonResponse, loadBYOClient } from './fixtures/byoClientHarness';

const env = {
  VITE_AZURE_OPENAI_ENDPOINT: 'https://offline.openai.azure.com',
  VITE_AZURE_OPENAI_DEPLOYMENT_GPT61SOL: 'sol-production-alias',
  VITE_AZURE_FOUNDRY_ENDPOINT: 'https://offline.services.ai.azure.com',
  VITE_AZURE_FOUNDRY_DEPLOYMENT_OPUS55: 'opus-production-alias',
};
for (const model of ['gpt-6.1-sol', 'claude-opus-5-5'] as const) {
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
      if (model === 'gpt-6.1-sol') return jsonResponse({
        status: 'completed', output_text: content, usage: { input_tokens: 10, output_tokens: 20, total_tokens: 30 },
      });
      return jsonResponse({
        type: 'message', stop_reason: 'end_turn', content: [
          { type: 'thinking', thinking: 'not output' },
          { type: 'text', text: request.body.output_config.format ? JSON.stringify({ azureDiagramJson: content }) : content },
        ], usage: { input_tokens: 10, output_tokens: 20, cache_read_input_tokens: 5 },
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
    assert.equal(imageRequest.body[model === 'gpt-6.1-sol' ? 'input' : 'messages'][model === 'gpt-6.1-sol' ? 1 : 0].content[1].type,
      model === 'gpt-6.1-sol' ? 'input_image' : 'image');
    content = '{"suggestions":["Enable private endpoints"]}';
    await client.generateFollowUpSuggestions({ services: ['App Service'], lastChange: 'Added app', recentRequests: [], modelOverride: top });
    assert.equal(sent.length, 9);
    for (const request of sent) {
      assert.equal(request.deployment, model === 'gpt-6.1-sol' ? env.VITE_AZURE_OPENAI_DEPLOYMENT_GPT61SOL : env.VITE_AZURE_FOUNDRY_DEPLOYMENT_OPUS55);
      assert.equal(request.body.model, request.deployment);
      assert.equal(request.apiFormat, model === 'gpt-6.1-sol' ? 'responses' : 'anthropic-messages');
      assert.equal(request.byo, undefined);
      assert.equal(model === 'gpt-6.1-sol' ? request.body.max_output_tokens : request.body.max_tokens, 32000);
      assert.equal(model === 'gpt-6.1-sol' ? request.body.reasoning.effort : request.body.output_config.effort, 'max');
    }
    for (const metrics of [topology.metrics, review.metrics, deploymentGuide.metrics, vision.metrics]) {
      assert.equal(metrics?.source, 'managed');
      assert.equal(metrics?.deployment, sent[0].deployment);
      assert.equal(metrics?.reasoningEffort, 'max');
      assert.equal(metrics?.totalTokens, model === 'gpt-6.1-sol' ? 30 : 35);
    }
  });
}

test('Claude parsing excludes thinking and rejects truncation, refusal and malformed JSON envelopes', async () => {
  const { client } = await loadBYOClient(undefined, env);
  for (const stop_reason of ['max_tokens', 'refusal', 'tool_use', 'pause_turn']) {
    assert.throws(() => client.parseApiResponse({ stop_reason, content: [{ type: 'text', text: '{}' }] }, 'anthropic-messages'), { code: 'incomplete_ai_response' });
  }
  assert.throws(() => client.parseApiResponse({
    stop_reason: 'end_turn', content: [{ type: 'text', text: '{"azureDiagramJson":"broken"}' }],
  }, 'anthropic-messages'), { code: 'invalid_upstream_response' });
  assert.throws(() => client.buildRequestBody({
    apiFormat: 'anthropic-messages', deployment: 'opus', messages: [{ role: 'user', content: [
      { type: 'input_image', image_url: 'https://remote.example/image.png' },
    ] }], reasoningEffort: 'high', isReasoning: true, maxTokens: 32000,
  }), { code: 'invalid_upstream_request' });
});
