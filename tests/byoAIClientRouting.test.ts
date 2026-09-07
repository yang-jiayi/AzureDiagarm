import test from 'node:test';
import assert from 'node:assert/strict';
import type { FeatureType } from '../src/stores/modelSettingsStore';
import type { BYOAIProfile } from '../src/stores/byoAISettingsStore';
import {
  capabilityResponse, deferred, jsonResponse, loadBYOClient, settle, testResponse, verifiedProfile,
} from './fixtures/byoClientHarness';

const features: FeatureType[] = ['architectureGeneration', 'validation', 'deploymentGuide', 'blueprint'];
const key = 'sk-offline-routing-only-key';
const profile = (apiFormat: BYOAIProfile['apiFormat'] = 'responses', provider: BYOAIProfile['provider'] = 'azure-openai'): BYOAIProfile => ({
  id: 'custom-production', name: 'Customer production', provider,
  endpoint: provider === 'openai' ? 'https://api.openai.com' : 'https://customer.openai.azure.com',
  model: provider === 'openai' ? 'ft:customer-model:org:2026' : 'actual-customer-deployment-alias',
  apiFormat, reasoningEffort: 'max', isReasoning: true, supportsVision: true, maxCompletionTokens: 32768,
});

const bodyText = (envelope: any) => JSON.stringify(envelope.body.input ?? envelope.body.messages);
function providerResponse(content: string, apiFormat: string) {
  return jsonResponse(apiFormat === 'responses'
    ? { status: 'completed', output_text: content, usage: { input_tokens: 10, output_tokens: 20, total_tokens: 30 } }
    : { choices: [{ finish_reason: 'stop', message: { content } }], usage: { prompt_tokens: 10, completion_tokens: 20, total_tokens: 30 } });
}

for (const provider of ['azure-openai', 'openai'] as const) {
  for (const apiFormat of ['responses', 'chat-completions'] as const) {
    test(`all feature snapshots route ${provider}/${apiFormat} to the actual identifier without managed configuration`, async t => {
      const envelopes: any[] = [];
      t.mock.method(globalThis, 'fetch', async (url: unknown, options?: RequestInit) => {
        if (url === '/api/runtime-config') return capabilityResponse();
        assert.equal(url, '/api/openai');
        const envelope = JSON.parse(String(options?.body));
        envelopes.push(envelope);
        return testResponse(apiFormat);
      });
      const { client } = await loadBYOClient(undefined, {});
      assert.equal(client.isManagedAIModelConfigured(), false);
      const saved = await verifiedProfile(client, profile(apiFormat, provider), key);
      client.selectBYOAIProfile(saved.id);
      assert.equal(client.isAnyAIModelConfigured(), true);
      assert.deepEqual(Object.keys(client.MODEL_CONFIG), ['gpt-6-astra']);
      for (const feature of features) {
        const capture = client.captureRuntimeModelOverride(feature);
        assert.equal(Object.isFrozen(capture.connection), true);
        assert.doesNotMatch(JSON.stringify(capture), /apiKey|sk-offline/);
        const runtime = client.resolveAIModelRuntime(feature, { ...capture });
        assert.equal(runtime.source, 'bring-your-own');
        assert.equal(runtime.model, saved.model);
        assert.equal(runtime.deployment, saved.model);
        assert.equal(runtime.reasoningEffort, 'max');
        assert.equal(runtime.maxCompletionTokens, 32768);
        assert.doesNotMatch(JSON.stringify(runtime), /apiKey|sk-offline/);
        const body = client.buildRequestBody({
          ...runtime, messages: [{ role: 'user', content: 'Generation, not a test' }], maxTokens: runtime.maxCompletionTokens,
        });
        assert.equal(body.model, saved.model);
        assert.equal(apiFormat === 'responses' ? body.max_output_tokens : body.max_completion_tokens, 32768);
        assert.equal(apiFormat === 'responses' ? body.reasoning.effort : body.reasoning_effort, 'max');
        const result = await client.callAzureOpenAIProxy({
          apiFormat, deployment: runtime.deployment, body, connection: runtime.connection,
        });
        assert.equal(result.ok, true);
        assert.deepEqual(envelopes.at(-1), {
          apiFormat, deployment: saved.model, body,
          byo: { provider, endpoint: saved.endpoint, apiKey: key },
        });
      }
      assert.equal(envelopes.length, 5);
      client.selectBYOAIProfile(null);
      assert.equal(client.isAnyAIModelConfigured(), false);
      assert.throws(() => client.resolveAIModelRuntime('validation'), { code: 'astra_not_configured' });
    });
  }
}

test('connection testing always re-reads server opt-in and cannot use a stale enabled UI flag', async t => {
  let enabled = true;
  let posts = 0;
  t.mock.method(globalThis, 'fetch', async (url: unknown) => {
    if (url === '/api/runtime-config') return capabilityResponse(enabled);
    posts++;
    return testResponse();
  });
  const { client } = await loadBYOClient(undefined, {});
  await client.loadRuntimeConfig();
  const saved = client.upsertBYOAIProfile(profile());
  client.setBYOAIApiKey(saved.id, key);
  enabled = false;
  await assert.rejects(client.testBYOAIConnection(saved.id), { code: 'byo_not_enabled' });
  assert.equal(posts, 0);
  assert.equal(client.getBYOAIConnectionState(saved.id).verified, false);
});

test('profile/key/selection changes cannot redirect queued captured generation, even when switching back', async t => {
  const envelopes: any[] = [];
  t.mock.method(globalThis, 'fetch', async (url: unknown, options?: RequestInit) => {
    if (url === '/api/runtime-config') return capabilityResponse();
    envelopes.push(JSON.parse(String(options?.body)));
    return testResponse();
  });
  const { client } = await loadBYOClient();
  const a = await verifiedProfile(client, profile(), key);
  const b = await verifiedProfile(client, { ...profile(), id: 'other', model: 'other-actual-deployment' }, 'sk-other-offline-key');
  client.selectBYOAIProfile(a.id);
  const captures = features.map(feature => client.captureRuntimeModelOverride(feature));
  client.selectBYOAIProfile(b.id);
  client.selectBYOAIProfile(a.id);
  for (let i = 0; i < features.length; i++) {
    assert.throws(() => client.resolveAIModelRuntime(features[i], captures[i]), { code: 'stale_ai_configuration' });
  }
  const current = client.captureRuntimeModelOverride('architectureGeneration');
  client.setBYOAIApiKey(a.id, 'sk-edited-offline-key');
  await assert.rejects(client.callAzureOpenAI([{ role: 'user', content: 'Queued work' }], current), { code: 'stale_ai_configuration' });
  assert.equal(envelopes.length, 2, 'only explicit connection tests were sent');
});

test('forged executable fields and legacy managed models are rejected even while BYO is active', async t => {
  let calls = 0;
  t.mock.method(globalThis, 'fetch', async (url: unknown) => {
    if (url === '/api/runtime-config') return capabilityResponse();
    calls++;
    return testResponse();
  });
  const { client } = await loadBYOClient();
  const saved = await verifiedProfile(client, profile(), key);
  client.selectBYOAIProfile(saved.id);
  for (const override of [
    { model: 'gpt-5.6-sol', reasoningEffort: 'max' },
    { model: 'gpt-6-astra', reasoningEffort: 'max', forceManaged: true },
    { model: 'gpt-6-astra', reasoningEffort: 'max', byo: { apiKey: key } },
    { model: 'gpt-6-astra', reasoningEffort: 'max', deployment: saved.model },
    { model: 'gpt-6-astra', reasoningEffort: 'max', connection: { source: 'bring-your-own', profileId: saved.id } },
  ]) assert.throws(() => Reflect.apply(client.resolveAIModelRuntime, undefined, ['validation', override]));
  const capture = client.captureRuntimeModelOverride('validation');
  assert.throws(() => client.resolveAIModelRuntime('validation', {
    ...capture, connection: { ...capture.connection! },
  }), { code: 'stale_ai_configuration' }, 'copying/reconstructing connection fields must not forge a capture');
  const runtime = client.resolveAIModelRuntime('validation', capture);
  for (const request of [
    { apiFormat: 'responses', deployment: saved.model, body: { model: 'different' }, connection: runtime.connection },
    { apiFormat: 'responses', deployment: saved.model, body: { model: saved.model, apiKey: key }, connection: runtime.connection },
    { apiFormat: 'responses', deployment: saved.model, body: { model: saved.model }, connection: runtime.connection, byo: { apiKey: key } },
    { apiFormat: 'responses', deployment: saved.model, body: { model: saved.model }, connection: runtime.connection, endpoint: 'https://evil.example' },
  ]) await assert.rejects(Reflect.apply(client.callAzureOpenAIProxy, undefined, [request]));
  assert.equal(calls, 1);
});

test('Responses-to-Chat vision conversion preserves detail and input snapshots without dropping the actual model', async () => {
  const { client } = await loadBYOClient();
  const messages = [
    { role: 'system', content: 'Read exact labels' },
    { role: 'user', content: [
      { type: 'input_text', text: 'Read this image' },
      { type: 'input_image', image_url: 'data:image/png;base64,QUJD', detail: 'high' },
    ] },
  ];
  const original = structuredClone(messages);
  const body = client.buildRequestBody({
    deployment: 'actual-chat-alias', messages, maxTokens: 32768,
    apiFormat: 'chat-completions', isReasoning: true, reasoningEffort: 'max',
  });
  assert.equal(body.model, 'actual-chat-alias');
  assert.equal(body.max_completion_tokens, 32768);
  assert.equal(body.reasoning_effort, 'max');
  assert.deepEqual(body.messages[1].content, [
    { type: 'text', text: 'Read this image' },
    { type: 'image_url', image_url: { url: 'data:image/png;base64,QUJD', detail: 'high' } },
  ]);
  body.messages[1].content[1].image_url.url = 'mutated';
  assert.deepEqual(messages, original);
  const responseBody = client.buildRequestBody({
    deployment: 'alias', messages, maxTokens: 32000, apiFormat: 'responses', isReasoning: true, reasoningEffort: 'max',
  });
  responseBody.input[1].content[0].text = 'mutated';
  assert.deepEqual(messages, original);
});

test('nonreasoning Chat omits reasoning rather than substituting a managed model or a guessed effort', async () => {
  const { client } = await loadBYOClient();
  const body = client.buildRequestBody({
    deployment: 'customer-chat-model', messages: [{ role: 'user', content: 'hello' }],
    maxTokens: 32768, apiFormat: 'chat-completions', isReasoning: false, reasoningEffort: 'max', jsonOutput: false,
  });
  assert.equal(body.model, 'customer-chat-model');
  assert.equal(body.max_tokens, 32768);
  assert.equal('reasoning_effort' in body, false);
  assert.equal('response_format' in body, false);
  assert.equal('temperature' in body, false);
});

for (const format of ['responses', 'chat-completions'] as const) {
  test(`actual ${format} providers preserve captured BYO routing, quality, provenance and image format`, async t => {
    t.mock.method(console, 'log', () => {});
    const sent: any[] = [];
    let content = '';
    t.mock.method(globalThis, 'fetch', async (url: unknown, options?: RequestInit) => {
      if (url === '/api/runtime-config') return capabilityResponse();
      if (url === '/api/docs-search') return jsonResponse({ results: [] });
      assert.equal(url, '/api/openai');
      const envelope = JSON.parse(String(options?.body));
      sent.push(envelope);
      return bodyText(envelope).includes('Connection test only.') ? testResponse(format) : providerResponse(content, format);
    });
    const { client } = await loadBYOClient(undefined, {});
    const saved = await verifiedProfile(client, profile(format), key);
    client.selectBYOAIProfile(saved.id);
    const top = client.captureRuntimeModelOverride('architectureGeneration');
    const blueprint = client.captureRuntimeModelOverride('blueprint');
    const validation = client.captureRuntimeModelOverride('validation');
    const guide = client.captureRuntimeModelOverride('deploymentGuide');
    const architecture = { groups: [], services: [{ id: 'app', name: 'App Service', type: 'App Service', category: 'app services' }], connections: [] };
    content = JSON.stringify(architecture);
    const topology = await client.generateArchitectureWithAI('Offline architecture', top);
    await client.generateArchitectureFromIaC({ format: 'arm', content: { resources: [] }, filenames: ['offline.json'] }, 'en', top);
    content = JSON.stringify({
      title: 'Offline', stages: [{ id: 'stage', label: 'Application', services: [{ id: 'app', name: 'App Service', category: 'app services', role: 'Application' }] }], connections: [],
    });
    await client.generateReferenceArchitectureWithAI('Offline reference', top);
    content = JSON.stringify({
      title: 'Offline', zones: [{ id: 'zone', label: 'Application', x: 0, y: 0, width: 640, height: 480 }],
      nodes: [{ id: 'app', name: 'App Service', kind: 'service', zone: 'zone', x: 1, y: 1 }], edges: [],
    });
    await client.generateBlueprintArchitectureWithAI('Offline blueprint', blueprint);
    content = JSON.stringify({
      title: 'Offline', zones: [{ id: 'zone', label: 'Application' }],
      components: [{ id: 'app', name: 'App Service', role: 'Application', zone: 'zone' }], connections: [],
    });
    await client.generateComponentManifest('Offline manifest', top);
    content = JSON.stringify({ overallScore: 80, summary: 'Offline review', pillars: [], quickWins: [] });
    const review = await client.validateArchitecture([{ name: 'App Service', type: 'App Service', category: 'app services' }], [], undefined, undefined, validation);
    content = JSON.stringify({ title: 'Offline guide', deploymentSteps: [{ step: 1, title: 'Deploy', description: 'Offline step' }] });
    const deploymentGuide = await client.generateDeploymentGuide([], [], undefined, undefined, undefined, 'en', { modelOverride: guide });
    content = 'An exact offline image description.';
    const vision = await client.analyzeArchitectureDiagramImage('QUJD', 'image/png', 'en', { modelOverride: top });
    const imageRequest = sent.at(-1);
    const image = (format === 'responses' ? imageRequest.body.input : imageRequest.body.messages)[1].content[1];
    assert.equal(image.type, format === 'responses' ? 'input_image' : 'image_url');
    assert.equal(format === 'responses' ? image.image_url : image.image_url.url, 'data:image/png;base64,QUJD');
    content = '{"suggestions":["Enable private endpoints"]}';
    await client.generateFollowUpSuggestions({ services: ['App Service'], lastChange: 'Added app', recentRequests: [], modelOverride: top });
    assert.equal(sent.length, 10, 'one test plus nine actual feature calls');
    for (const envelope of sent.slice(1)) {
      assert.equal(envelope.deployment, saved.model);
      assert.equal(envelope.body.model, saved.model);
      assert.equal(envelope.apiFormat, format);
      assert.equal(envelope.byo.apiKey, key);
      assert.equal(format === 'responses' ? envelope.body.max_output_tokens : envelope.body.max_completion_tokens, 32768);
      assert.equal(format === 'responses' ? envelope.body.reasoning.effort : envelope.body.reasoning_effort, 'max');
    }
    for (const metrics of [topology.metrics, review.metrics, deploymentGuide.metrics, vision.metrics]) {
      assert.equal(metrics?.source, 'bring-your-own');
      assert.equal(metrics?.deployment, saved.model);
      assert.equal(metrics?.profileId, saved.id);
      assert.equal(metrics?.reasoningEffort, 'max');
      assert.match(metrics?.model || '', /actual-customer-deployment-alias/);
      assert.doesNotMatch(JSON.stringify(metrics), /apiKey|sk-offline/);
    }
  });
}

test('preparation delay cannot redirect deployment guidance after a connection edit', async t => {
  t.mock.method(console, 'log', () => {});
  t.mock.method(console, 'error', () => {});
  const grounding = deferred<Response>();
  let generations = 0;
  t.mock.method(globalThis, 'fetch', async (url: unknown, options?: RequestInit) => {
    if (url === '/api/runtime-config') return capabilityResponse();
    if (url === '/api/docs-search') return grounding.promise;
    if (!bodyText(JSON.parse(String(options?.body))).includes('Connection test only.')) generations++;
    return testResponse();
  });
  const { client } = await loadBYOClient();
  const saved = await verifiedProfile(client, profile(), key);
  client.selectBYOAIProfile(saved.id);
  const request = client.generateDeploymentGuide([], []);
  const rejected = assert.rejects(request, { code: 'stale_ai_configuration' });
  client.upsertBYOAIProfile({ ...saved, model: 'edited-after-submission' });
  grounding.resolve(jsonResponse({ results: [] }));
  await rejected;
  assert.equal(generations, 0);
});

test('a dispatched response retains captured provenance even if the profile is renamed, edited and switched later', async t => {
  t.mock.method(console, 'log', () => {});
  const completion = deferred<Response>();
  let sent: any;
  t.mock.method(globalThis, 'fetch', async (url: unknown, options?: RequestInit) => {
    if (url === '/api/runtime-config') return capabilityResponse();
    const envelope = JSON.parse(String(options?.body));
    if (bodyText(envelope).includes('Connection test only.')) return testResponse();
    sent = envelope;
    return completion.promise;
  });
  const { client } = await loadBYOClient();
  const saved = await verifiedProfile(client, profile(), key);
  client.selectBYOAIProfile(saved.id);
  const captured = client.captureRuntimeModelOverride('architectureGeneration');
  const request = client.callAzureOpenAI([{ role: 'user', content: 'Queued architecture' }], captured);
  await settle();
  assert.equal(sent.deployment, saved.model);
  client.upsertBYOAIProfile({ ...saved, name: 'Renamed', model: 'new-deployment' });
  client.setBYOAIApiKey(saved.id, 'sk-edited-after-dispatch');
  client.selectBYOAIProfile(null);
  completion.resolve(providerResponse('{"result":"ok"}', 'responses'));
  const result = await request;
  assert.equal(result.metrics.model, captured.connection!.displayName);
  assert.equal(result.metrics.deployment, saved.model);
  assert.equal(result.metrics.source, 'bring-your-own');
  assert.equal(result.metrics.profileId, saved.id);
  assert.equal(sent.byo.apiKey, key);
});

test('test-connection is distinctly labeled, bounded, abortable, and does not log generation telemetry', async t => {
  t.mock.timers.enable({ apis: ['setTimeout'] });
  const pending = deferred<Response>();
  let init: RequestInit | undefined;
  let envelope: any;
  t.mock.method(globalThis, 'fetch', async (url: unknown, options?: RequestInit) => {
    if (url === '/api/runtime-config') return capabilityResponse();
    init = options;
    envelope = JSON.parse(String(options?.body));
    return pending.promise;
  });
  const { client } = await loadBYOClient(undefined, {});
  const saved = client.upsertBYOAIProfile(profile());
  client.setBYOAIApiKey(saved.id, key);
  const controller = new AbortController();
  const request = client.testBYOAIConnection(saved.id, { signal: controller.signal });
  const rejection = assert.rejects(request, { name: 'AbortError', userCancelled: true });
  await settle();
  assert.equal((init!.headers as Record<string, string>)['X-AzureDiagarm-Operation'], 'byo-connection-test');
  assert.match(bodyText(envelope), /Connection test only/);
  assert.equal(envelope.body.max_output_tokens, client.BYO_AI_TEST_MAX_TOKENS);
  assert.equal(envelope.body.reasoning.effort, 'max', 'testing never rewrites saved reasoning');
  assert.equal(client.getBYOAISettings().profiles[0].maxCompletionTokens, 32768);
  controller.abort();
  await rejection;
  assert.equal(init?.signal?.aborted, true);
  pending.resolve(testResponse());
  await settle();
  t.mock.timers.tick(100_000);
  assert.equal(client.getBYOAIConnectionState(saved.id).status, 'unverified');
  assert.equal(client.getTestModelUsage().length, 0);
  assert.equal(client.getBYOAISettings().activeProfileId, null);
});

test('connection test timeout is bounded even when the provider ignores cancellation', async t => {
  t.mock.timers.enable({ apis: ['setTimeout'] });
  let signal: AbortSignal | undefined;
  t.mock.method(globalThis, 'fetch', async (url: unknown, options?: RequestInit) => {
    if (url === '/api/runtime-config') return capabilityResponse();
    signal = options?.signal ?? undefined;
    return new Promise<Response>(() => {});
  });
  const { client } = await loadBYOClient();
  const saved = client.upsertBYOAIProfile(profile());
  client.setBYOAIApiKey(saved.id, key);
  const request = client.testBYOAIConnection(saved.id);
  const rejection = assert.rejects(request, { code: 'byo_test_timeout' });
  await settle();
  t.mock.timers.tick(client.BYO_AI_TEST_TIMEOUT_MS);
  await rejection;
  assert.equal(signal?.aborted, true);
  assert.equal(client.getBYOAIConnectionState(saved.id).status, 'failed');
});

test('cancelling before testing prevents capability and inference HTTP entirely', async t => {
  const fetch = t.mock.method(globalThis, 'fetch', async () => { throw new Error('Unexpected HTTP'); });
  const { client } = await loadBYOClient();
  const saved = client.upsertBYOAIProfile(profile());
  client.setBYOAIApiKey(saved.id, key);
  const controller = new AbortController();
  controller.abort();
  await assert.rejects(client.testBYOAIConnection(saved.id, { signal: controller.signal }), { name: 'AbortError' });
  assert.equal(fetch.mock.callCount(), 0);
  assert.equal(client.getBYOAIConnectionState(saved.id).status, 'unverified');
});

test('failed test preserves HTTP/source/code/request IDs and never exposes raw errors or keys', async t => {
  t.mock.method(globalThis, 'fetch', async (url: unknown) => url === '/api/runtime-config' ? capabilityResponse() : jsonResponse({
    error: {
      source: 'byo_ai', code: 'byo_authentication_failed', message: `Raw upstream credential=${key}`,
      requestId: 'app-request-id', upstreamStatus: 401, upstreamCode: 'invalid_api_key',
      upstreamRequestId: 'provider-request-id', credentials: { apiKey: key },
    },
  }, 401, { 'x-azurediagarm-request-id': 'app-request-id' }));
  const { client } = await loadBYOClient();
  const saved = client.upsertBYOAIProfile(profile());
  client.setBYOAIApiKey(saved.id, key);
  await assert.rejects(client.testBYOAIConnection(saved.id), (error: any) => {
    assert.equal(error.status, 401);
    assert.equal(error.source, 'byo_ai');
    assert.equal(error.code, 'byo_authentication_failed');
    assert.equal(error.requestId, 'app-request-id');
    assert.equal(error.upstreamRequestId, 'provider-request-id');
    assert.equal(error.upstreamStatus, 401);
    assert.equal(error.upstreamCode, 'invalid_api_key');
    assert.match(error.message, /API key.*test the connection/i);
    assert.doesNotMatch(JSON.stringify(error), /sk-offline|Raw upstream|credentials/);
    return true;
  });
  const state = client.getBYOAIConnectionState(saved.id);
  assert.equal(state.error?.code, 'byo_authentication_failed');
  assert.equal(state.error?.requestId, 'app-request-id');
  assert.doesNotMatch(JSON.stringify(state), /sk-offline|Raw upstream|credentials/);
});

test('secret-shaped diagnostic identifiers and raw network failures are never returned as public messages', async t => {
  let scenario = 'diagnostics';
  t.mock.method(globalThis, 'fetch', async (url: unknown) => {
    if (url === '/api/runtime-config') return capabilityResponse();
    if (scenario === 'network') throw new Error(`Network failure with ${key}`);
    return jsonResponse({ error: {
      source: key, code: key, requestId: key, upstreamRequestId: key, message: key,
    } }, 403, { 'x-azurediagarm-request-id': key, 'x-upstream-request-id': key });
  });
  const { client } = await loadBYOClient();
  const saved = client.upsertBYOAIProfile(profile());
  client.setBYOAIApiKey(saved.id, key);
  for (const value of ['diagnostics', 'network']) {
    scenario = value;
    await assert.rejects(client.testBYOAIConnection(saved.id), (error: any) => {
      assert.doesNotMatch(error.message, /sk-offline|Network failure with/);
      assert.doesNotMatch(JSON.stringify(error), /sk-offline|Network failure with/);
      return true;
    });
    assert.doesNotMatch(JSON.stringify(client.getBYOAIConnectionState(saved.id)), /sk-offline|Network failure with/);
  }
});

for (const payload of [
  { choices: [] },
  { choices: [{ message: { content: '{"status":"ok"}' } }] },
  { choices: [{ finish_reason: 'length', message: { content: '{"status":"ok"}' } }] },
  { choices: [{ finish_reason: 'stop', message: { content: '{"status":"ok"}', refusal: 'Not allowed' } }] },
  { choices: [{ finish_reason: 'stop', message: { content: '{"status":"ok"}', tool_calls: [{}] } }] },
]) {
  test(`Chat verification rejects incomplete/tool/refusal output: ${JSON.stringify(payload)}`, async t => {
    t.mock.method(globalThis, 'fetch', async (url: unknown) => url === '/api/runtime-config' ? capabilityResponse() : jsonResponse(payload));
    const { client } = await loadBYOClient();
    const saved = client.upsertBYOAIProfile(profile('chat-completions'));
    client.setBYOAIApiKey(saved.id, key);
    await assert.rejects(client.testBYOAIConnection(saved.id), { code: 'byo_test_incomplete' });
    assert.equal(client.getBYOAIConnectionState(saved.id).verified, false);
  });
}

test('a selected profile without vision fails before HTTP instead of borrowing managed vision', async t => {
  let posts = 0;
  t.mock.method(globalThis, 'fetch', async (url: unknown) => {
    if (url === '/api/runtime-config') return capabilityResponse();
    posts++;
    return testResponse();
  });
  const { client } = await loadBYOClient();
  const saved = await verifiedProfile(client, { ...profile(), supportsVision: false }, key);
  client.selectBYOAIProfile(saved.id);
  const captured = client.captureRuntimeModelOverride('architectureGeneration');
  await assert.rejects(client.analyzeArchitectureDiagramImage('QUJD', 'image/png', 'en', { modelOverride: captured }),
    { code: 'byo_vision_not_supported' });
  await assert.rejects(client.callAzureOpenAI([
    { role: 'user', content: [{ type: 'input_image', image_url: 'data:image/png;base64,QUJD' }] },
  ], captured), { code: 'byo_vision_not_supported' });
  assert.equal(posts, 1, 'only the explicit connection test was sent');
});

test('BYO rate-limit retry preserves the captured model, key, MAX effort and complete output cap', async t => {
  t.mock.timers.enable({ apis: ['setTimeout'] });
  t.mock.method(console, 'log', () => {});
  t.mock.method(console, 'error', () => {});
  const sent: any[] = [];
  t.mock.method(globalThis, 'fetch', async (url: unknown, options?: RequestInit) => {
    if (url === '/api/runtime-config') return capabilityResponse();
    const envelope = JSON.parse(String(options?.body));
    if (bodyText(envelope).includes('Connection test only.')) return testResponse();
    sent.push(envelope);
    return sent.length === 1
      ? jsonResponse({ error: { source: 'byo_ai', code: 'byo_rate_limited' } }, 429, { 'Retry-After': '2' })
      : providerResponse('{"result":"ok"}', 'responses');
  });
  const { client } = await loadBYOClient();
  const saved = await verifiedProfile(client, profile(), key);
  client.selectBYOAIProfile(saved.id);
  const request = client.callAzureOpenAI([{ role: 'user', content: 'Preserve full fidelity' }], client.captureRuntimeModelOverride('architectureGeneration'));
  await settle();
  assert.equal(sent.length, 1);
  t.mock.timers.tick(2000);
  await request;
  assert.equal(sent.length, 2);
  assert.deepEqual(sent[1], sent[0]);
  assert.equal(sent[1].body.reasoning.effort, 'max');
  assert.equal(sent[1].body.max_output_tokens, 32768);
});

test('key changes during a provider cooldown reject stale retries rather than replacing the key or routing managed', async t => {
  t.mock.timers.enable({ apis: ['setTimeout'] });
  t.mock.method(console, 'log', () => {});
  t.mock.method(console, 'error', () => {});
  let generations = 0;
  t.mock.method(globalThis, 'fetch', async (url: unknown, options?: RequestInit) => {
    if (url === '/api/runtime-config') return capabilityResponse();
    if (bodyText(JSON.parse(String(options?.body))).includes('Connection test only.')) return testResponse();
    generations++;
    return jsonResponse({ error: { source: 'byo_ai', code: 'byo_rate_limited' } }, 429, { 'Retry-After': '2' });
  });
  const { client } = await loadBYOClient();
  const saved = await verifiedProfile(client, profile(), key);
  client.selectBYOAIProfile(saved.id);
  const request = client.callAzureOpenAI([{ role: 'user', content: 'Preserve full fidelity' }], client.captureRuntimeModelOverride('architectureGeneration'));
  const rejected = assert.rejects(request, { code: 'stale_ai_configuration' });
  await settle();
  client.setBYOAIApiKey(saved.id, 'sk-updated-during-retry');
  t.mock.timers.tick(2000);
  await rejected;
  assert.equal(generations, 1);
  assert.equal(client.getTestModelUsage().length, 0);
});

for (const state of ['key-required', 'unverified', 'missing-profile', 'disabled', 'failed'] as const) {
  test(`selected BYO ${state} never dispatches an automatic managed fallback`, async t => {
    let posts = 0;
    let enabled = true;
    let successful = true;
    t.mock.method(globalThis, 'fetch', async (url: unknown) => {
      if (url === '/api/runtime-config') return capabilityResponse(enabled);
      posts++;
      return successful ? testResponse() : jsonResponse({ status: 'incomplete' });
    });
    const { client } = await loadBYOClient();
    const saved = await verifiedProfile(client, profile(), key);
    client.selectBYOAIProfile(saved.id);
    if (state === 'key-required') client.setBYOAIApiKey(saved.id, '');
    if (state === 'unverified') client.upsertBYOAIProfile({ ...saved, maxCompletionTokens: 32000 });
    if (state === 'missing-profile') client.removeBYOAIProfile(saved.id);
    if (state === 'disabled') {
      enabled = false;
      await client.loadRuntimeConfig(true);
    }
    if (state === 'failed') {
      successful = false;
      await assert.rejects(client.testBYOAIConnection(saved.id));
    }
    const before = posts;
    for (const feature of features) {
      const info = client.getEffectiveAIModelInfo(feature);
      assert.equal(info.source, 'bring-your-own');
      assert.equal(info.ready, false);
      assert.throws(() => client.captureRuntimeModelOverride(feature));
    }
    await assert.rejects(client.callAzureOpenAI([{ role: 'user', content: 'Must not substitute managed' }]));
    assert.equal(posts, before);
    assert.equal(client.getBYOAISettings().activeProfileId, saved.id);
  });
}

test('captured transport rejects body-level reasoning or output downgrades before dispatch', async t => {
  let posts = 0;
  t.mock.method(globalThis, 'fetch', async (url: unknown) => {
    if (url === '/api/runtime-config') return capabilityResponse();
    posts++;
    return testResponse();
  });
  const { client } = await loadBYOClient();
  const saved = await verifiedProfile(client, profile(), key);
  client.selectBYOAIProfile(saved.id);
  const runtime = client.resolveAIModelRuntime('architectureGeneration');
  const body = client.buildRequestBody({
    ...runtime, messages: [{ role: 'user', content: 'Full fidelity' }], maxTokens: runtime.maxCompletionTokens,
  });
  for (const edit of [{ max_output_tokens: 4096 }, { reasoning: { effort: 'none' } }]) {
    await assert.rejects(client.callAzureOpenAIProxy({
      apiFormat: runtime.apiFormat, deployment: runtime.deployment, body: { ...body, ...edit }, connection: runtime.connection,
    }), { code: 'stale_ai_configuration' });
  }
  assert.equal(posts, 1);
});

test('completed BYO guide and Bicep ZIP filenames retain artifact provenance after switching to managed', async t => {
  t.mock.method(console, 'log', () => {});
  t.mock.method(Date, 'now', () => 777);
  t.mock.method(globalThis, 'fetch', async (url: unknown, options?: RequestInit) => {
    if (url === '/api/runtime-config') return capabilityResponse();
    if (url === '/api/docs-search') return jsonResponse({ results: [] });
    const envelope = JSON.parse(String(options?.body));
    if (bodyText(envelope).includes('Connection test only.')) return testResponse();
    return providerResponse(JSON.stringify({
      title: 'Offline deployment', overview: 'Artifact provenance regression',
      deploymentSteps: [{ step: 1, title: 'Deploy', description: 'Offline command' }],
      bicepTemplates: [{ name: 'Main', description: 'Offline template', filename: 'main.bicep', content: 'param location string' }],
    }), 'responses');
  });
  const { client } = await loadBYOClient();
  const saved = await verifiedProfile(client, profile(), key);
  client.selectBYOAIProfile(saved.id);
  const capture = client.captureRuntimeModelOverride('deploymentGuide');
  const capturedConnection = capture.connection;
  const guide = await client.generateDeploymentGuide([], [], undefined, undefined, undefined, 'en', { modelOverride: capture });
  client.selectBYOAIProfile(null);
  client.updateModelSettings({ reasoningEffort: 'low' });
  const downloads: string[] = [];
  const previousDocument = Object.getOwnPropertyDescriptor(globalThis, 'document');
  Object.defineProperty(globalThis, 'document', { configurable: true, value: {
    createElement: () => ({
      href: '', download: '',
      click() { downloads.push(this.download); },
    }),
    body: { appendChild() {}, removeChild() {} },
  } });
  t.after(() => {
    if (previousDocument) Object.defineProperty(globalThis, 'document', previousDocument);
    else Reflect.deleteProperty(globalThis, 'document');
  });
  t.mock.method(URL, 'createObjectURL', () => 'blob:offline-artifact');
  t.mock.method(URL, 'revokeObjectURL', () => {});
  client.downloadDeploymentGuide(guide);
  await client.downloadAllBicepTemplates(guide);
  assert.deepEqual(downloads, [
    'deployment-guide-777-byo-actual-customer-deployment-alias-max.md',
    'bicep-templates-777-byo-actual-customer-deployment-alias-max.zip',
  ]);
  assert.equal(capture.connection, capturedConnection);
  assert.equal(client.generateModelFilename('manual-diagram', 'json', 777), 'manual-diagram-777.json');
});

for (const [label, provenance] of [
  ['BYO MAX', { source: 'bring-your-own', model: 'BYO original profile', deployment: 'actual-review-alias', reasoningEffort: 'max' }],
  ['managed Astra MAX', { source: 'managed', model: 'GPT-6 Astra', reasoningEffort: 'max' }],
  ['historical model', { model: 'gpt-5.6-terra', reasoningEffort: 'high' }],
  ['missing provenance', undefined],
] as const) {
  test(`validation report diagram link matches the explicit artifact filename: ${label}`, async () => {
    const { client } = await loadBYOClient();
    client.updateModelSettings({ reasoningEffort: 'low' });
    const validation = {
      overallScore: 80, summary: 'Offline report', timestamp: '2026-09-07T00:00:00.000Z',
      pillars: [], quickWins: [], diagramImageDataUrl: 'data:image/png;base64,QUJD',
      ...(provenance ? { metrics: Object.freeze({
        ...provenance, promptTokens: 10, completionTokens: 20, totalTokens: 30, elapsedTimeMs: 1,
      }) } : {}),
    };
    const expectedFilename = client.generateModelFilename(
      'architecture-validation-diagram', 'png', new Date(validation.timestamp).getTime(), validation.metrics,
    );
    const report = client.formatValidationReport(validation);
    assert.ok(report.includes(`![Architecture Diagram](./${expectedFilename})`), `Expected paired image ${expectedFilename}`);
    assert.doesNotMatch(report, /data:image\/png/);
  });
}
