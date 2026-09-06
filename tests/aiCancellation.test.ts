import test, { afterEach, beforeEach, mock } from 'node:test';
import assert from 'node:assert/strict';
import { build } from 'esbuild';
import { exactJapanese } from '../src/i18n/LanguageContext';

// Bundle the actual providers/proxy with only configuration and telemetry
// replaced. No browser, deployment, credentials, or network is required.
const bundled = await build({
  stdin: {
    contents: `
      export * from './src/services/azureOpenAI';
      export { generateReferenceArchitectureWithAI } from './src/services/referenceArchitectureAI';
      export { generateBlueprintArchitectureWithAI } from './src/services/blueprintArchitectureAI';
      export { generateComponentManifest } from './src/services/componentManifestAI';
      export { validateArchitecture } from './src/services/architectureValidator';
      export { runAIBudgetQueue } from './src/services/aiBudgetQueue';
      export { getAIBudget } from './src/services/aiBudgetService';
      export { getTestModelUsage, resetTestModelUsage } from './src/services/telemetryService';
      export { setTestModelSettings } from './src/stores/modelSettingsStore';
    `,
    resolveDir: process.cwd(), loader: 'ts',
  },
  bundle: true, write: false, platform: 'node', format: 'esm', logLevel: 'silent',
  define: { 'import.meta.env': JSON.stringify({ VITE_AZURE_OPENAI_ENDPOINT: 'configured' }) },
  plugins: [{
    name: 'isolated-provider-config',
    setup(build) {
      build.onResolve({ filter: /modelSettingsStore$/ }, () => ({ path: 'settings', namespace: 'test-config' }));
      build.onResolve({ filter: /telemetryService$/ }, () => ({ path: 'telemetry', namespace: 'test-config' }));
      build.onLoad({ filter: /.*/, namespace: 'test-config' }, args => ({
        contents: args.path === 'telemetry' ? `
          const events = [];
          export const trackAIModelUsage = event => events.push(event);
          export const getTestModelUsage = () => events;
          export const resetTestModelUsage = () => { events.length = 0; };
        ` : `
          let settings = { model: 'test-model', reasoningEffort: 'none' };
          export const setTestModelSettings = value => { settings = value; };
          export const getModelSettings = () => settings;
          export const getModelSettingsForFeature = () => settings;
          export const getAvailableModels = () => ['test-model'];
          export const getDeploymentName = model => model === 'gpt-6-astra' ? 'gpt-6-astra' : 'test-deployment';
          export const MODEL_CONFIG = {
            'test-model': {displayName:'Test', apiFormat:'responses', isReasoning:false, maxCompletionTokens:1000},
            'gpt-6-astra': {displayName:'GPT-6 Astra', apiFormat:'responses', isReasoning:true, maxCompletionTokens:32000},
            'grok-4.1-fast': {displayName:'Fast', apiFormat:'chat-completions', isReasoning:false, maxCompletionTokens:1000}
          };
        `, loader: 'js',
      }));
    },
  }],
});
const provider = await import(`data:text/javascript;base64,${Buffer.from(`${bundled.outputFiles[0].text}\n//# sourceURL=ai-provider-test.mjs`).toString('base64')}`);
beforeEach(() => {
  provider.setTestModelSettings({ model: 'test-model', reasoningEffort: 'none' });
  provider.resetTestModelUsage();
  mock.method(console, 'log', () => {});
  mock.method(console, 'error', () => {});
  mock.method(console, 'warn', () => {});
});
afterEach(() => mock.reset());
const override = { model: 'test-model', reasoningEffort: 'none' };
const abortError = (error: unknown) => error instanceof Error && error.name === 'AbortError';
const response = (content = '{"services":[{"id":"app","name":"Web App","type":"App Service"}],"connections":[],"groups":[]}') =>
  new Response(JSON.stringify({ output_text: content, usage: {} }), { status: 200, headers: { 'content-type': 'application/json' } });
const validationServices = [{ name: 'Web App', type: 'App Service', category: 'app services' }];
const validationContent = '{"overallScore":80,"summary":"A test review.","pillars":[],"quickWins":[]}';
const validate = (signal?: AbortSignal) => provider.validateArchitecture(
  validationServices, [], undefined, undefined, { ...override, forceManaged: true, signal }, 'en',
);

test('Astra MAX honors a real 429 Retry-After without changing its prompt, model, or 32K output cap', async t => {
  t.mock.timers.enable({ apis: ['setTimeout'] });
  const bodies: any[] = [];
  const waits: unknown[] = [];
  const controller = new AbortController();
  t.after(() => controller.abort());
  t.mock.method(globalThis, 'fetch', async (_url: unknown, options: RequestInit) => {
    bodies.push(JSON.parse(options.body as string));
    if (bodies.length === 1) return new Response(JSON.stringify({
      error: { source: 'azure_openai', code: 'azure_openai_rate_limited' },
    }), { status: 429, headers: { 'Content-Type': 'application/json', 'Retry-After': '60' } });
    return response();
  });
  const request = provider.generateArchitectureWithAI(
    'もっともセキュアの構成で、FabricのE2EのArchitecture図を作成してください。',
    { model: 'gpt-6-astra', reasoningEffort: 'max', signal: controller.signal },
    undefined, 'ja', { onRetryWait: (wait: unknown) => waits.push(wait) },
  );
  void request.catch(() => {});
  await new Promise<void>(resolve => setImmediate(resolve));
  assert.equal(bodies.length, 1, '429 must wait instead of immediately replaying at lower quality');
  assert.equal(waits.length, 1, 'the fifth-argument progress callback must reach the transport');
  t.mock.timers.tick(59_999);
  await new Promise<void>(resolve => setImmediate(resolve));
  assert.equal(bodies.length, 1);
  t.mock.timers.tick(1);
  assert.equal((await request).services.length, 1);
  assert.equal(bodies.length, 2);
  assert.deepEqual(bodies[1], bodies[0]);
  assert.equal(bodies[1].deployment, 'gpt-6-astra');
  assert.equal(bodies[1].body.reasoning.effort, 'max');
  assert.equal(bodies[1].body.max_output_tokens, 32_000);
  assert.equal(provider.getTestModelUsage().length, 1);
  assert.equal(waits.at(-1), null);
});

test('cancelling an actual blueprint transport during Retry-After prevents replay and success telemetry', async t => {
  t.mock.timers.enable({ apis: ['setTimeout'] });
  const controller = new AbortController();
  const waits: unknown[] = [];
  const fetch = t.mock.method(globalThis, 'fetch', async () => new Response(JSON.stringify({
    error: { source: 'azure_openai', code: 'azure_openai_rate_limited' },
  }), { status: 429, headers: { 'Content-Type': 'application/json', 'Retry-After': '60' } }));
  const request = provider.generateBlueprintArchitectureWithAI('Secure Fabric E2E', {
    model: 'gpt-6-astra', reasoningEffort: 'max', signal: controller.signal,
    onRetryWait: (wait: unknown) => waits.push(wait),
  });
  const rejection = assert.rejects(request, { name: 'AbortError', userCancelled: true });
  await new Promise<void>(resolve => setImmediate(resolve));
  assert.equal(fetch.mock.callCount(), 1);
  assert.equal(waits.length, 1);
  controller.abort();
  await rejection;
  t.mock.timers.tick(1_000_000);
  await new Promise<void>(resolve => setImmediate(resolve));
  assert.equal(fetch.mock.callCount(), 1);
  assert.equal(provider.getTestModelUsage().length, 0);
  assert.equal(waits.length, 1, 'cancellation must not publish a resuming event');
});

for (const [stage, generate] of [
  ['manifest', provider.generateComponentManifest],
  ['topology', provider.generateArchitectureWithAI],
  ['blueprint', provider.generateBlueprintArchitectureWithAI],
] as const) {
  test(`${stage} preserves a structured daily-budget rejection without any retry or success`, async t => {
    const waits: unknown[] = [];
    const fetch = t.mock.method(globalThis, 'fetch', async () => new Response(JSON.stringify({
      error: { source: 'budget', code: 'ai_daily_budget_exceeded', requestId: 'budget-rejection-id' },
    }), { status: 429, headers: { 'Content-Type': 'application/json', 'Retry-After': '5' } }));
    await assert.rejects(generate('Secure Fabric E2E', {
      model: 'gpt-6-astra', reasoningEffort: 'max', onRetryWait: (wait: unknown) => waits.push(wait),
    }), (error: any) => {
      assert.equal(error.source, 'budget');
      assert.equal(error.code, 'ai_daily_budget_exceeded');
      assert.equal(error.status, 429);
      assert.equal(error.requestId, 'budget-rejection-id');
      assert.equal(error.retryAfterMs, 5000);
      assert.match(error.message, /daily AI budget.*midnight UTC/);
      return true;
    });
    assert.equal(fetch.mock.callCount(), 1);
    assert.equal(waits.length, 0);
    assert.equal(provider.getTestModelUsage().length, 0);
  });

  test(`${stage} preserves provider 500 provenance without automatic lower-quality replay`, async t => {
    const bodies: any[] = [];
    t.mock.method(globalThis, 'fetch', async (_url: unknown, options: RequestInit) => {
      bodies.push(JSON.parse(options.body as string));
      return new Response(JSON.stringify({
        error: {
          source: 'azure_openai', code: 'azure_openai_unavailable', requestId: 'proxy-500',
          upstreamStatus: 500, upstreamCode: 'server_error', upstreamRequestId: 'provider-500',
        },
      }), { status: 500, headers: { 'Content-Type': 'application/json' } });
    });
    await assert.rejects(generate('Secure Fabric E2E', {
      model: 'gpt-6-astra', reasoningEffort: 'max',
    }), (error: any) => {
      assert.equal(error.source, 'azure_openai');
      assert.equal(error.code, 'azure_openai_unavailable');
      assert.equal(error.status, 500);
      assert.equal(error.upstreamStatus, 500);
      assert.equal(error.upstreamCode, 'server_error');
      assert.equal(error.requestId, 'proxy-500');
      assert.equal(error.upstreamRequestId, 'provider-500');
      return true;
    });
    assert.equal(bodies.length, 1);
    assert.equal(bodies[0].body.reasoning.effort, 'max');
    assert.equal(bodies[0].body.max_output_tokens, 32000);
    assert.equal(provider.getTestModelUsage().length, 0);
  });
}

test('already-aborted generation never starts a request', async t => {
  const fetch = t.mock.method(globalThis, 'fetch', async () => response());
  const controller = new AbortController();
  controller.abort(new Error('Custom cancellation reason'));
  await assert.rejects(provider.generateArchitectureWithAI('test', override, undefined, 'en', { signal: controller.signal }), abortError);
  assert.equal(fetch.mock.callCount(), 0);
});

test('AbortSignal reaches actual fetch and cancellation never retries or becomes a timeout', async t => {
  let signal: AbortSignal | undefined;
  const fetch = t.mock.method(globalThis, 'fetch', async (_url: unknown, options: RequestInit) => {
    signal = options.signal as AbortSignal;
    return await new Promise<Response>((_resolve, reject) => signal!.addEventListener('abort',
      () => reject(new DOMException('Aborted', 'AbortError')), { once: true }));
  });
  const controller = new AbortController();
  const request = provider.generateArchitectureWithAI('test', override, undefined, 'en', { signal: controller.signal });
  assert.ok(signal);
  assert.equal(signal.aborted, false);
  controller.abort();
  await assert.rejects(request, abortError);
  assert.equal(signal.aborted, true);
  assert.equal(fetch.mock.callCount(), 1);
});

test('late successful provider responses cannot escape a cancelled request', async t => {
  let complete!: (response: Response) => void;
  t.mock.method(globalThis, 'fetch', () => new Promise<Response>(resolve => { complete = resolve; }));
  const controller = new AbortController();
  const request = provider.generateArchitectureWithAI('test', override, undefined, 'en', { signal: controller.signal });
  controller.abort();
  complete(response());
  await assert.rejects(request, abortError);
});

test('abort during response-body reading is not converted into a malformed-response failure', async t => {
  let finishBody!: (body: string) => void;
  t.mock.method(globalThis, 'fetch', async () => ({
    ok: true, status: 200, headers: new Headers({ 'content-type': 'application/json' }), url: '', redirected: false,
    text: () => new Promise<string>(resolve => { finishBody = resolve; }),
  } as Response));
  const controller = new AbortController();
  const request = provider.generateArchitectureWithAI('test', override, undefined, 'en', { signal: controller.signal });
  await Promise.resolve();
  controller.abort();
  finishBody(JSON.stringify({ output_text: '{"services":[]}' }));
  await assert.rejects(request, abortError);
});

test('transport timeouts remain distinguishable from user cancellation and abort the request once', async t => {
  t.mock.timers.enable({ apis: ['setTimeout'] });
  const fetch = t.mock.method(globalThis, 'fetch', async (_url: unknown, options: RequestInit) =>
    await new Promise<Response>((_resolve, reject) => options.signal!.addEventListener('abort',
      () => reject(new DOMException('Aborted', 'AbortError')), { once: true })));
  const request = provider.callAzureOpenAI([], override);
  t.mock.timers.tick(225001);
  await assert.rejects(request, /timed out after 225 seconds/);
  assert.equal(fetch.mock.callCount(), 1);
});

test('a cancelled request can be retried with a fresh signal; legacy calls still work', async t => {
  let invocation = 0;
  t.mock.method(globalThis, 'fetch', async (_url: unknown, options: RequestInit) => {
    if (++invocation > 1) return response();
    return await new Promise<Response>((_resolve, reject) => options.signal!.addEventListener('abort',
      () => reject(new DOMException('Aborted', 'AbortError')), { once: true }));
  });
  const cancelled = new AbortController();
  const first = provider.generateArchitectureWithAI('test', override, undefined, 'en', { signal: cancelled.signal });
  cancelled.abort();
  await assert.rejects(first, abortError);
  assert.equal((await provider.generateArchitectureWithAI('test', override, undefined, 'en',
    { signal: new AbortController().signal })).services.length, 1);
  assert.equal((await provider.generateArchitectureWithAI('test', override)).services.length, 1);
  assert.equal(invocation, 3);
});

test('completion removes the external abort listener', async t => {
  t.mock.method(globalThis, 'fetch', async () => response());
  const controller = new AbortController();
  const add = t.mock.method(controller.signal, 'addEventListener');
  const remove = t.mock.method(controller.signal, 'removeEventListener');
  await provider.generateArchitectureWithAI('test', override, undefined, 'en', { signal: controller.signal });
  assert.equal(add.mock.callCount(), 1);
  assert.equal(remove.mock.callCount(), 1);
  assert.equal(add.mock.calls[0].arguments[1], remove.mock.calls[0].arguments[1]);
});

test('already-aborted validation never dispatches or records model usage', async t => {
  const fetch = t.mock.method(globalThis, 'fetch', async () => response(validationContent));
  const controller = new AbortController();
  controller.abort();
  await assert.rejects(validate(controller.signal), (error: any) => abortError(error) && error.userCancelled === true);
  assert.equal(fetch.mock.callCount(), 0);
  assert.deepEqual(provider.getTestModelUsage(), []);
});

test('validation cancellation reaches fetch and remains distinct from a provider timeout', async t => {
  let signal: AbortSignal | undefined;
  const fetch = t.mock.method(globalThis, 'fetch', async (_url: unknown, options: RequestInit) => {
    signal = options.signal as AbortSignal;
    return await new Promise<Response>((_resolve, reject) => signal!.addEventListener('abort',
      () => reject(new DOMException('Aborted', 'AbortError')), { once: true }));
  });
  const controller = new AbortController();
  const request = validate(controller.signal);
  assert.equal(signal?.aborted, false);
  controller.abort();
  await assert.rejects(request, (error: any) => abortError(error) && error.userCancelled === true);
  assert.equal(signal?.aborted, true);
  assert.equal(fetch.mock.callCount(), 1);
  assert.deepEqual(provider.getTestModelUsage(), []);
});

test('late validation responses cannot become successful reviews or usage telemetry after cancellation', async t => {
  let complete!: (response: Response) => void;
  t.mock.method(globalThis, 'fetch', () => new Promise<Response>(resolve => { complete = resolve; }));
  const controller = new AbortController();
  const request = validate(controller.signal);
  controller.abort();
  complete(response(validationContent));
  await assert.rejects(request, abortError);
  assert.deepEqual(provider.getTestModelUsage(), []);
});

test('validation cancellation during body reading is not reported as malformed JSON', async t => {
  let finishBody!: (body: string) => void;
  t.mock.method(globalThis, 'fetch', async () => ({
    ok: true, status: 200, headers: new Headers({ 'content-type': 'application/json' }), url: '', redirected: false,
    text: () => new Promise<string>(resolve => { finishBody = resolve; }),
  } as Response));
  const controller = new AbortController();
  const request = validate(controller.signal);
  await Promise.resolve();
  controller.abort();
  finishBody(JSON.stringify({ output_text: validationContent, usage: {} }));
  await assert.rejects(request, abortError);
  assert.deepEqual(provider.getTestModelUsage(), []);
});

test('validation keeps its 225-second timeout and does not count a timed-out call as success', async t => {
  t.mock.timers.enable({ apis: ['setTimeout'] });
  let signal: AbortSignal | undefined;
  const fetch = t.mock.method(globalThis, 'fetch', async (_url: unknown, options: RequestInit) => {
    signal = options.signal as AbortSignal;
    return await new Promise<Response>((_resolve, reject) => signal!.addEventListener('abort',
      () => reject(new DOMException('Aborted', 'AbortError')), { once: true }));
  });
  const request = validate();
  t.mock.timers.tick(225001);
  await assert.rejects(request, (error: any) =>
    !abortError(error) && error.userCancelled !== true && /taking too long to respond/.test(error.message));
  assert.equal(signal?.aborted, true);
  assert.equal(fetch.mock.callCount(), 1);
  assert.deepEqual(provider.getTestModelUsage(), []);
});

test('validation releases caller listeners and preserves the legacy six-argument API', async t => {
  t.mock.method(globalThis, 'fetch', async () => response(validationContent));
  const controller = new AbortController();
  const add = t.mock.method(controller.signal, 'addEventListener');
  const remove = t.mock.method(controller.signal, 'removeEventListener');
  assert.equal((await validate(controller.signal)).overallScore, 80);
  assert.equal(add.mock.callCount(), 1);
  assert.equal(remove.mock.callCount(), 1);
  assert.equal(add.mock.calls[0].arguments[1], remove.mock.calls[0].arguments[1]);
  const legacy = await provider.validateArchitecture(validationServices, [], undefined, undefined, override, 'en');
  assert.equal(legacy.overallScore, 80);
  assert.equal(legacy.modelUsed, 'Test');
  assert.equal(provider.getTestModelUsage().length, 2);
});

for (const feature of ['generation', 'validation']) {
  test(`${feature} preserves a concurrency rejection without an immediate compact retry or success telemetry`, async t => {
    const fetch = t.mock.method(globalThis, 'fetch', async () => new Response(JSON.stringify({
      error: { code: 'ai_concurrency_limit', source: 'budget', requestId: 'comparison-request' },
    }), { status: 429, headers: { 'content-type': 'application/json' } }));
    const request = feature === 'generation'
      ? provider.generateArchitectureWithAI('A web application', override)
      : validate();
    await assert.rejects(request, { code: 'ai_concurrency_limit', status: 429, requestId: 'comparison-request' });
    assert.equal(fetch.mock.callCount(), 1);
    assert.deepEqual(provider.getTestModelUsage(), []);
  });

  for (const limit of [1, 2]) {
    test(`${feature} queue checks the actual budget client and dispatches three providers within cap ${limit}`, async t => {
      let inFlight = 0;
      let peak = 0;
      let dispatches = 0;
      let reads = 0;
      t.mock.method(globalThis, 'fetch', async (url: unknown, options: RequestInit) => {
        if (url === '/api/ai/budget') {
          reads += 1;
          assert.equal(options.cache, 'no-store');
          assert.ok(options.signal);
          return new Response(JSON.stringify({
            available: true, concurrentLimit: limit, concurrentRequests: inFlight,
            limitTokens: 100_000, remainingTokens: 100_000, resetAt: '2026-09-07T00:00:00Z',
          }), { status: 200, headers: { 'content-type': 'application/json' } });
        }
        assert.equal(options.method, 'POST');
        assert.ok(options.signal);
        const envelope = JSON.parse(options.body as string);
        assert.equal(envelope.byo, undefined);
        inFlight += 1;
        dispatches += 1;
        peak = Math.max(peak, inFlight);
        await new Promise<void>(resolve => setImmediate(resolve));
        inFlight -= 1;
        const content = feature === 'validation' ? validationContent
          : '{"services":[{"id":"app","name":"Web App","type":"App Service"}],"connections":[],"groups":[]}';
        return new Response(JSON.stringify({
          output_text: content, choices: [{ message: { content } }], usage: {},
        }), { status: 200, headers: { 'content-type': 'application/json' } });
      });
      const results = await provider.runAIBudgetQueue(
        ['test-model', 'gpt-6-astra', 'grok-4.1-fast'].map(model => async (signal: AbortSignal) => {
          const managed = { model, reasoningEffort: 'none', forceManaged: true, signal };
          return feature === 'validation'
            ? provider.validateArchitecture(validationServices, [], undefined, undefined, managed, 'en')
            : provider.generateArchitectureWithAI('A web application', managed, undefined, 'en');
        }),
        { getBudget: provider.getAIBudget },
      );
      assert.equal(results.filter((result: PromiseSettledResult<unknown>) => result.status === 'fulfilled').length, 3);
      assert.equal(dispatches, 3);
      assert.equal(peak, limit);
      assert.ok(reads >= 2);
      assert.equal(provider.getTestModelUsage().length, 3);
    });
  }
}

test('legacy direct AbortSignal arguments preserve cancellation and never retry', async t => {
  const fetch = t.mock.method(globalThis, 'fetch', async (_url: unknown, options: RequestInit) =>
    await new Promise<Response>((_resolve, reject) => options.signal!.addEventListener('abort',
      () => reject(new DOMException('Aborted', 'AbortError')), { once: true })));
  const controller = new AbortController();
  const request = provider.generateArchitectureWithAI('test', override, undefined, 'en', controller.signal);
  controller.abort();
  await assert.rejects(request, (error: any) => abortError(error) && error.userCancelled === true);
  assert.equal(fetch.mock.callCount(), 1);
});

test('internal timeouts never replay at lower quality and still release caller listeners', async t => {
  t.mock.timers.enable({ apis: ['setTimeout'] });
  let calls = 0;
  const fetch = t.mock.method(globalThis, 'fetch', async (_url: unknown, options: RequestInit) => {
    if (++calls > 1) return response();
    return await new Promise<Response>((_resolve, reject) => options.signal!.addEventListener('abort',
      () => reject(new DOMException('Aborted', 'AbortError')), { once: true }));
  });
  const controller = new AbortController();
  const add = t.mock.method(controller.signal, 'addEventListener');
  const remove = t.mock.method(controller.signal, 'removeEventListener');
  const request = provider.generateArchitectureWithAI('test', override, undefined, 'en', { signal: controller.signal });
  t.mock.timers.tick(225001);
  await assert.rejects(request, (error: any) => {
    assert.equal(error.code, 'ai_client_timeout');
    assert.equal(error.source, 'client');
    assert.match(error.message, /timed out after 225 seconds/);
    return true;
  });
  assert.equal(fetch.mock.callCount(), 1);
  assert.equal(add.mock.callCount(), 1);
  assert.equal(remove.mock.callCount(), 1);
  assert.equal(provider.getTestModelUsage().length, 0);
});

test('empty architectures fail without clearing the canvas or replaying a compact request', async t => {
  const fetch = t.mock.method(globalThis, 'fetch', async () => response('{"services":[]}'));
  await assert.rejects(provider.generateArchitectureWithAI('test', override), /empty architecture/);
  assert.equal(fetch.mock.callCount(), 1);
});

for (const content of ['{"status":"ok"}', '{"services":{}}']) {
  test(`incompatible architecture response ${content} fails without a second billable request`, async t => {
    const fetch = t.mock.method(globalThis, 'fetch', async () => response(content));
    const message = 'Failed to generate architecture. Please try again.';
    await assert.rejects(provider.generateArchitectureWithAI('test', override), { message });
    assert.equal(fetch.mock.callCount(), 1);
    assert.match(exactJapanese[message], /[\u3040-\u30ff\u4e00-\u9faf]/);
  });
}

test('Astra follow-ups use the genuine deployment and configured feature reasoning, not the legacy fast model', async t => {
  provider.setTestModelSettings({ model: 'gpt-6-astra', reasoningEffort: 'low' });
  let requestBody: any;
  const fetch = t.mock.method(globalThis, 'fetch', async (_url: unknown, options: RequestInit) => {
    requestBody = JSON.parse(options.body as string);
    return response('{"suggestions":["Add monitoring"]}');
  });
  const suggestions = await provider.generateFollowUpSuggestions({
    services: ['Web App'], lastChange: 'Added Web App', recentRequests: [],
  });
  assert.deepEqual(suggestions, ['Add monitoring']);
  assert.equal(fetch.mock.callCount(), 1);
  assert.equal(requestBody.deployment, 'gpt-6-astra');
  assert.equal(requestBody.apiFormat, 'responses');
  assert.equal(requestBody.body.model, 'gpt-6-astra');
  assert.equal(requestBody.body.reasoning.effort, 'low');
});

test('Astra follow-up cancellation remains terminal and does not fall back to another model', async t => {
  provider.setTestModelSettings({ model: 'gpt-6-astra', reasoningEffort: 'low' });
  const fetch = t.mock.method(globalThis, 'fetch', async (_url: unknown, options: RequestInit) =>
    await new Promise<Response>((_resolve, reject) => options.signal!.addEventListener('abort',
      () => reject(new DOMException('Aborted', 'AbortError')), { once: true })));
  const controller = new AbortController();
  const request = provider.generateFollowUpSuggestions({
    services: [], lastChange: '', recentRequests: [], signal: controller.signal,
  });
  controller.abort();
  await assert.rejects(request, abortError);
  assert.equal(fetch.mock.callCount(), 1);
});

for (const code of ['image_not_supported', 'invalid_upstream_request']) {
  test(`${code} vision guidance does not recommend an unconfigured model or lose diagnostics`, async t => {
    const fetch = t.mock.method(globalThis, 'fetch', async () => new Response(JSON.stringify({
      error: { code, source: 'azure_openai', requestId: 'vision-request', upstreamRequestId: 'upstream-vision' },
    }), { status: 400, headers: { 'content-type': 'application/json' } }));
    const message = 'The selected model may not support image analysis. Choose a vision-capable model in AI settings.';
    await assert.rejects(provider.analyzeArchitectureDiagramImage('image', 'image/png', 'en'), (error: any) => {
      assert.equal(error.name, 'OpenAIProxyError');
      assert.equal(error.code, code);
      assert.equal(error.status, 400);
      assert.equal(error.source, 'azure_openai');
      assert.equal(error.requestId, 'vision-request');
      assert.equal(error.upstreamRequestId, 'upstream-vision');
      assert.equal(error.message, `${message} Request ID: vision-request`);
      return true;
    });
    assert.equal(fetch.mock.callCount(), 1);
    assert.match(exactJapanese[message], /画像分析/);
    assert.doesNotMatch(exactJapanese[message], /GPT-/);
  });
}

for (const [name, call] of [
  ['reference', (signal: AbortSignal) => provider.generateReferenceArchitectureWithAI('test', { ...override, signal }, 'en')],
  ['blueprint', (signal: AbortSignal) => provider.generateBlueprintArchitectureWithAI('test', { ...override, signal }, undefined, 'en')],
  ['manifest', (signal: AbortSignal) => provider.generateComponentManifest('test', { ...override, signal }, 'en')],
  ['vision', (signal: AbortSignal) => provider.analyzeArchitectureDiagramImage('image', 'image/png', 'en', { signal })],
  ['follow-ups', (signal: AbortSignal) => provider.generateFollowUpSuggestions({ services: [], lastChange: '', recentRequests: [], signal })],
] as const) {
  test(`${name} provider propagates cancellation instead of fallback output`, async t => {
    const fetch = t.mock.method(globalThis, 'fetch', async (_url: unknown, options: RequestInit) =>
      await new Promise<Response>((_resolve, reject) => options.signal!.addEventListener('abort',
        () => reject(new DOMException('Aborted', 'AbortError')), { once: true })));
    const controller = new AbortController();
    const request = call(controller.signal);
    controller.abort();
    await assert.rejects(request, abortError);
    assert.equal(fetch.mock.callCount(), 1);
  });
}
