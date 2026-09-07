import test, { afterEach, beforeEach, mock, type TestContext } from 'node:test';
import assert from 'node:assert/strict';
import { build } from 'esbuild';
import { exactJapanese } from '../src/i18n/LanguageContext';

// Bundle the actual providers/proxy with only configuration and telemetry
// replaced. No browser, deployment, credentials, or network is required.
const bundled = await build({
  stdin: {
    contents: `
      export * from './src/services/azureOpenAI';
      export { generateReferenceArchitectureWithAI, referenceToTopology } from './src/services/referenceArchitectureAI';
      export { generateBlueprintArchitectureWithAI } from './src/services/blueprintArchitectureAI';
      export { generateComponentManifest } from './src/services/componentManifestAI';
      export { validateArchitecture } from './src/services/architectureValidator';
      export { runAIBudgetQueue } from './src/services/aiBudgetQueue';
      export { getAIBudget } from './src/services/aiBudgetService';
      export { getTestModelUsage, resetTestModelUsage } from './src/services/telemetryService';
      export { updateModelSettings as setTestModelSettings } from './src/stores/modelSettingsStore';
    `,
    resolveDir: process.cwd(), loader: 'ts',
  },
  bundle: true, write: false, platform: 'node', format: 'esm', logLevel: 'silent',
  define: { 'import.meta.env': JSON.stringify({
    VITE_AZURE_OPENAI_ENDPOINT: 'configured',
    VITE_AZURE_OPENAI_DEPLOYMENT_GPT6ASTRA: 'gpt-6-astra',
  }) },
  plugins: [{
    name: 'isolated-provider-config',
    setup(build) {
      build.onResolve({ filter: /telemetryService$/ }, () => ({ path: 'telemetry', namespace: 'test-config' }));
      build.onLoad({ filter: /.*/, namespace: 'test-config' }, () => ({
        contents: `
          const events = [];
          export const trackAIModelUsage = event => events.push(event);
          export const getTestModelUsage = () => events;
          export const resetTestModelUsage = () => { events.length = 0; };
        `, loader: 'js',
      }));
    },
  }],
});
const providerUrl = `data:text/javascript;base64,${Buffer.from(`${bundled.outputFiles[0].text}\n//# sourceURL=ai-provider-test.mjs`).toString('base64')}`;
const provider = await import(providerUrl);
let warningMessages: string[] = [];
beforeEach(() => {
  provider.setTestModelSettings({ model: 'gpt-6-astra', reasoningEffort: 'none' });
  provider.resetTestModelUsage();
  mock.method(console, 'log', () => {});
  mock.method(console, 'error', () => {});
  warningMessages = [];
  mock.method(console, 'warn', (...args: unknown[]) => { warningMessages.push(args.map(String).join(' ')); });
});
afterEach(() => mock.reset());
const override = { model: 'gpt-6-astra', reasoningEffort: 'none' };
const abortError = (error: unknown) => error instanceof Error && error.name === 'AbortError';
const response = (content = '{"services":[{"id":"app","name":"Web App","type":"App Service"}],"connections":[],"groups":[]}') =>
  new Response(JSON.stringify({ output_text: content, usage: {} }), { status: 200, headers: { 'content-type': 'application/json' } });
const validationServices = [{ name: 'Web App', type: 'App Service', category: 'app services' }];
const validationContent = '{"overallScore":80,"summary":"A test review.","pillars":[],"quickWins":[]}';
const validate = (signal?: AbortSignal) => provider.validateArchitecture(
  validationServices, [], undefined, undefined, { ...override, signal }, 'en',
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
  await new Promise<void>(resolve => setImmediate(resolve));
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
  await new Promise<void>(resolve => setImmediate(resolve));
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
  assert.equal(legacy.modelUsed, 'GPT-6 Astra (none)');
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
    test(`${feature} queue checks the actual budget client and dispatches three Astra requests within cap ${limit}`, async t => {
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
        ['first', 'second', 'third'].map(brief => async (signal: AbortSignal) => {
          const managed = { model: 'gpt-6-astra', reasoningEffort: 'none', signal };
          return feature === 'validation'
            ? provider.validateArchitecture(validationServices, [], undefined, undefined, managed, 'en')
            : provider.generateArchitectureWithAI(`A web application: ${brief}`, managed, undefined, 'en');
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
  test(`${code} vision guidance addresses the Astra deployment without model switching or lost diagnostics`, async t => {
    const fetch = t.mock.method(globalThis, 'fetch', async () => new Response(JSON.stringify({
      error: { code, source: 'azure_openai', requestId: 'vision-request', upstreamRequestId: 'upstream-vision' },
    }), { status: 400, headers: { 'content-type': 'application/json' } }));
    const message = 'The configured GPT-6 Astra deployment rejected the image analysis request. Check the image and contact the application administrator if the problem persists.';
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
    assert.match(exactJapanese[message], /GPT-6 Astra/);
    assert.doesNotMatch(exactJapanese[message], /別のモデル|モデルを選択/);
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

const astraOverride = { model: 'gpt-6-astra', reasoningEffort: 'max' };
function mockModelResponse(t: TestContext, payload: unknown) {
  provider.setTestModelSettings(astraOverride);
  return t.mock.method(globalThis, 'fetch', async (_url: unknown, options: RequestInit) => {
    const request = JSON.parse(String(options.body));
    assert.equal(request.deployment, 'gpt-6-astra');
    assert.equal(request.body.model, 'gpt-6-astra');
    assert.equal(request.body.reasoning.effort, 'max');
    assert.equal(request.body.max_output_tokens, 32000);
    return response(JSON.stringify(payload));
  });
}

const localResponseFailure = (error: unknown) => {
  assert.ok(error instanceof Error);
  assert.equal(error.name, 'AIResponseValidationError');
  assert.equal(Reflect.get(error, 'code'), 'invalid_model_response');
  assert.equal(Reflect.get(error, 'source'), 'client');
  assert.equal(Reflect.get(error, 'retryable'), false);
  assert.equal(typeof Reflect.get(error, 'detail'), 'string');
  assert.match(error.message, /AI model returned.*invalid/i);
  return true;
};

for (const [name, generate] of [
  ['architecture', () => provider.generateArchitectureWithAI('offline fixture', astraOverride)],
  ['IaC import', () => provider.generateArchitectureFromIaC({
    format: 'arm', content: { resources: [] }, filenames: ['offline.json'],
  })],
] as const) {
  test(`response contracts: ${name} allocates against occupied IDs without rebinding unknown parents`, async t => {
    const payload = {
      groups: [
        { id: 'api', label: 'App' },
        { id: 'group-api-2', label: 'Data' },
        { id: 'group-api-4', label: 'Other' },
      ],
      services: [
        { id: 'api', name: 'Web App', type: 'App Service', groupId: 'api' },
        { id: 'group-api', name: 'Storage', type: 'Storage Account', groupId: 'api' },
        { id: 'group-api-3', name: 'Worker', groupId: 'group-api-2' },
        { id: 'leaf', name: 'Other', groupId: 'group-api-4' },
        { id: 'orphan', name: 'Orphan', groupId: 'group-api-5' },
      ],
      connections: [{ from: 'api', to: 'group-api', label: 'Store' }],
      workflow: [],
    };
    const fetch = mockModelResponse(t, payload);
    const result = await generate();
    const ids = [...result.groups, ...result.services].map((item: { id: string }) => item.id);
    assert.equal(new Set(ids).size, ids.length);
    assert.equal(result.groups[0].id, 'group-api-5');
    assert.deepEqual(result.services.map((item: { groupId: string | null }) => item.groupId),
      ['group-api-5', 'group-api-5', 'group-api-2', 'group-api-4', null]);
    assert.deepEqual(result.services.map((item: { id: string }) => item.id), payload.services.map(item => item.id));
    assert.deepEqual(result.connections, payload.connections);
    assert.equal(fetch.mock.callCount(), 1);
    const warnings = warningMessages.join('\n');
    assert.match(warnings, /collides with a service ID/);
    assert.match(warnings, /references unknown group.*clearing/);
  });

  test(`response contracts: ${name} remaps simultaneous collisions from original IDs only`, async t => {
    const fetch = mockModelResponse(t, {
      groups: ['api', 'group-api'],
      services: [
        { id: 'api', name: 'One', groupId: 'api' },
        { id: 'group-api', name: 'Two', groupId: 'group-api' },
      ],
      connections: [],
    });
    const result = await generate();
    assert.deepEqual(result.groups.map((group: { id: string }) => group.id), ['group-api-2', 'group-group-api']);
    assert.deepEqual(result.services.map((service: { groupId: string }) => service.groupId), ['group-api-2', 'group-group-api']);
    assert.equal(fetch.mock.callCount(), 1);
  });

  for (const [label, groups, services] of [
    ['duplicate groups', ['app', 'app'], [{ id: 'n', name: 'Node', groupId: 'app' }]],
    ['duplicate services', ['app'], [{ id: 'n', name: 'One' }, { id: 'n', name: 'Two' }]],
    ['missing service ID', [], [{ name: 'Node' }]],
    ['null service', [], [null]],
    ['null group', [null], [{ id: 'n', name: 'Node' }]],
  ] as const) {
    test(`response contracts: ${name} rejects ${label} locally without replay`, async t => {
      const fetch = mockModelResponse(t, { groups, services, connections: [] });
      await assert.rejects(generate(), localResponseFailure);
      assert.equal(fetch.mock.callCount(), 1);
    });
  }
}

const blueprintFixture = () => ({
  title: 'Offline blueprint',
  canvas: { width: 1600, height: 1000 },
  nodes: [{ id: 'app', name: 'Web App', category: 'app services', x: 100, y: 100, zone: 'a' }],
  zones: [
    { id: 'a', label: 'A', x: 0, y: 0, width: 500, height: 500 },
    { id: 'b', label: 'B', x: 600, y: 0, width: 500, height: 500 },
  ],
  edges: [],
});

for (const [label, zones] of [
  ['self-cycle', [{ ...blueprintFixture().zones[0], parent: 'a' }]],
  ['parent cycle', blueprintFixture().zones.map((zone, index) => ({ ...zone, parent: index ? 'a' : 'b' }))],
  ['missing parent', [{ ...blueprintFixture().zones[0], parent: 'missing' }]],
  ['duplicate zones', [blueprintFixture().zones[0], blueprintFixture().zones[0]]],
  ['null zone', [null]],
] as const) {
  test(`response contracts: blueprint rejects ${label} before hierarchy traversal`, async t => {
    const fetch = mockModelResponse(t, { ...blueprintFixture(), zones });
    await assert.rejects(provider.generateBlueprintArchitectureWithAI('offline fixture', astraOverride), localResponseFailure);
    assert.equal(fetch.mock.callCount(), 1);
  });
}

test('response contracts: valid nested blueprint preserves workflow warning and partial output', async t => {
  const fixture = blueprintFixture();
  const fetch = mockModelResponse(t, {
    ...fixture, zones: [fixture.zones[0], { ...fixture.zones[1], parent: 'a' }],
    workflow: [{ step: 1, description: 'A later connection is not available yet.' }],
  });
  const result = await provider.generateBlueprintArchitectureWithAI('offline fixture', astraOverride);
  assert.equal(result.nodes.length, 1);
  assert.equal(result.edges.length, 0);
  assert.equal(result.zones[1].parent, 'a');
  assert.equal(result.metrics.model, 'GPT-6 Astra');
  assert.equal(fetch.mock.callCount(), 1);
  assert.ok(Number.isFinite(result.canvas.width) && Number.isFinite(result.canvas.height));
  assert.ok(warningMessages.some(message => /edges are missing step numbers/.test(message)));
});

const referenceStage = () => ({
  id: 'app', label: 'Application',
  services: [{ id: 'web', name: 'Web App', category: 'app services' }],
});
for (const [label, stages] of [
  ['no stages', []],
  ['empty services', [{ ...referenceStage(), services: [] }]],
  ['missing services', [{ id: 'app', label: 'Application' }]],
  ['non-array services', [{ ...referenceStage(), services: {} }]],
  ['null stage', [null]],
  ['null service', [{ ...referenceStage(), services: [null] }]],
  ['missing service name', [{ ...referenceStage(), services: [{ id: 'web', category: 'app services' }] }]],
  ['missing service category', [{ ...referenceStage(), services: [{ id: 'web', name: 'Web App' }] }]],
  ['empty stage label', [{ ...referenceStage(), label: ' ' }]],
  ['duplicate stages', [referenceStage(), { ...referenceStage(), services: [{ id: 'other', name: 'Other', category: 'general' }] }]],
  ['duplicate services across stages', [referenceStage(), { ...referenceStage(), id: 'other' }]],
  ['empty stage beside valid stage', [referenceStage(), { id: 'empty', label: 'Empty', services: [] }]],
] as const) {
  test(`response contracts: reference rejects ${label} rather than publishing unusable success`, async t => {
    const fixture = { title: 'Offline', stages, connections: [] };
    const fetch = mockModelResponse(t, fixture);
    await assert.rejects(provider.generateReferenceArchitectureWithAI('offline fixture', astraOverride), localResponseFailure);
    assert.throws(() => provider.referenceToTopology(fixture), localResponseFailure);
    assert.equal(fetch.mock.callCount(), 1);
  });
}

for (const [name, generate] of [
  ['blueprint', () => provider.generateBlueprintArchitectureWithAI('offline fixture', astraOverride)],
  ['reference', () => provider.generateReferenceArchitectureWithAI('offline fixture', astraOverride)],
] as const) {
  test(`response contracts: ${name} rejects a null root with a typed local error`, async t => {
    const fetch = mockModelResponse(t, null);
    await assert.rejects(generate(), localResponseFailure);
    assert.equal(fetch.mock.callCount(), 1);
  });
}

test('response contracts: a valid one-stage reference without optional output remains successful', async t => {
  const fetch = mockModelResponse(t, { title: 'Offline', stages: [referenceStage()] });
  const result = await provider.generateReferenceArchitectureWithAI('offline fixture', astraOverride);
  const topology = provider.referenceToTopology(result);
  assert.equal(topology.services.length, 1);
  assert.deepEqual(topology.connections, []);
  assert.equal(result.metrics.model, 'GPT-6 Astra');
  assert.equal(fetch.mock.callCount(), 1);
});

test('response contracts: migrated selected BYO cannot silently fall back to managed Astra vision', async t => {
  const requests: Array<{ apiFormat: string; deployment: string; byo?: unknown; body: { input: Array<{ content: unknown }> } }> = [];
  t.mock.method(globalThis, 'fetch', async (url: unknown, options: RequestInit) => {
    assert.equal(url, '/api/openai');
    requests.push(JSON.parse(String(options.body)));
    return new Response(JSON.stringify({ output_text: 'Offline description' }),
      { headers: { 'content-type': 'application/json' } });
  });
  const legacyRecord = JSON.stringify({
    version: 2, enabled: true, provider: 'openai', endpoint: 'https://retired.example.com',
    model: 'gpt-4o', apiFormat: 'chat-completions', reasoningEffort: 'max',
  });
  const entries = new Map([['azure-diagrams-byo-ai-settings', legacyRecord]]);
  const storageDescriptor = Object.getOwnPropertyDescriptor(globalThis, 'localStorage');
  Object.defineProperty(globalThis, 'localStorage', { configurable: true, value: {
    getItem: (key: string) => entries.get(key) ?? null,
    setItem: (key: string, value: string) => { entries.set(key, value); },
  } });
  t.after(() => {
    if (storageDescriptor) Object.defineProperty(globalThis, 'localStorage', storageDescriptor);
    else Reflect.deleteProperty(globalThis, 'localStorage');
  });
  // Initialize real services with stale storage already present, not after hydration.
  const staleProvider = await import(`${providerUrl}#stale-byo`);
  await assert.rejects(staleProvider.analyzeArchitectureDiagramImage('QUJD', 'image/png'),
    { name: 'AIModelConfigurationError', code: 'byo_availability_unknown' });
  assert.equal(requests.length, 0);
  const migrated = JSON.parse(entries.get('azure-diagrams-byo-ai-settings')!);
  assert.equal(migrated.version, 3);
  assert.equal(migrated.activeProfileId, migrated.profiles[0].id);
  assert.equal(migrated.profiles[0].model, 'gpt-4o');
});

test('Astra-only public providers reject legacy/forged overrides before any HTTP or success', async t => {
  const fetch = t.mock.method(globalThis, 'fetch', async () => { throw new Error('Unexpected HTTP'); });
  for (const model of ['gpt-5.6-sol', 'gpt-5.6-terra', 'gpt-5.6-luna', 'claude-opus-5', 'constructor', undefined]) {
    const unsupported = { model, reasoningEffort: 'max', forceManaged: true };
    for (const generate of [
      () => provider.callAzureOpenAI([], unsupported),
      () => provider.generateArchitectureWithAI('Offline', unsupported),
      () => provider.generateBlueprintArchitectureWithAI('Offline', unsupported),
      () => provider.generateReferenceArchitectureWithAI('Offline', unsupported),
      () => provider.generateComponentManifest('Offline', unsupported),
      () => provider.validateArchitecture(validationServices, [], undefined, undefined, unsupported),
    ]) {
      await assert.rejects(generate(), { name: 'AIModelConfigurationError', code: 'unsupported_ai_model' });
    }
  }
  assert.equal(fetch.mock.callCount(), 0);
  assert.equal(provider.getTestModelUsage().length, 0);
});
