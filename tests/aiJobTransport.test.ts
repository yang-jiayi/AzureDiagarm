import test, { type TestContext } from 'node:test';
import assert from 'node:assert/strict';
import { build } from 'esbuild';

const bundled = await build({
  stdin: { contents: "export * from './src/services/aiJobTransport';", resolveDir: process.cwd(), loader: 'ts' },
  bundle: true, write: false, platform: 'node', format: 'cjs', logLevel: 'silent',
  define: { 'import.meta.env': '{}' },
});
const module = { exports: {} };
new Function('module', 'exports', bundled.outputFiles[0].text)(module, module.exports);
const { fetchAIJob } = module.exports as typeof import('../src/services/aiJobTransport');
const json = (value: unknown, status = 200) => new Response(JSON.stringify(value), {
  status, headers: { 'Content-Type': 'application/json' },
});
const init = () => ({
  method: 'POST', headers: { 'Content-Type': 'application/json' },
  body: JSON.stringify({ body: { reasoning: { effort: 'max' }, max_output_tokens: 32000 }, byo: { apiKey: 'test-only-byo-key' } }),
});
const state = (id: string, status: string) => ({
  job: { id, status, elapsedMs: 250000, deadlineAt: Date.now() + 900000, pollAfterMs: 1000 },
});
async function pump<T>(t: TestContext, work: Promise<T>, ticks = 25): Promise<T> {
  let settled = false;
  void work.finally(() => { settled = true; }).catch(() => {});
  for (let i = 0; !settled && i < ticks; i++) {
    for (let flush = 0; flush < 4; flush++) await new Promise<void>(resolve => setImmediate(resolve));
    if (!settled) t.mock.timers.tick(1000);
  }
  assert.equal(settled, true, 'job transport should finish within the bounded test window');
  return work;
}

test('generation uses short owner-scoped polls and a separate result without resending BYO credentials', async t => {
  t.mock.timers.enable({ apis: ['setTimeout', 'Date'] });
  const requests: Array<{ url: string; init: RequestInit }> = [];
  const progress: string[] = [];
  let id = '';
  let polls = 0;
  t.mock.method(globalThis, 'fetch', async (url: unknown, options: RequestInit = {}) => {
    requests.push({ url: String(url), init: options });
    if (options.method === 'POST') {
      const headers = new Headers(options.headers);
      id = headers.get('Idempotency-Key')!;
      assert.equal(headers.get('Prefer'), 'respond-async');
      return json(state(id, 'queued'), 202);
    }
    if (String(url).endsWith('/result')) return json({ output_text: 'complete', usage: { total_tokens: 12 } });
    return json(state(id, ++polls === 1 ? 'running' : 'succeeded'));
  });
  const response = await pump(t, fetchAIJob(init(), { onProgress: value => progress.push(value.status) }));
  assert.equal((await response.json()).output_text, 'complete');
  assert.deepEqual(progress, ['queued', 'running', 'succeeded']);
  assert.equal(requests.filter(value => value.init.method === 'POST').length, 1);
  for (const request of requests.slice(1)) {
    assert.ok(request.url === `/api/openai/jobs/${id}` || request.url === `/api/openai/jobs/${id}/result`);
    assert.equal(request.init.body, undefined);
    assert.doesNotMatch(JSON.stringify([...new Headers(request.init.headers)]), /test-only-byo-key/);
    assert.equal(request.init.credentials, 'same-origin');
    assert.equal(request.init.redirect, 'error');
    assert.equal(request.init.cache, 'no-store');
  }
  assert.equal(requests.some(value => value.init.method === 'DELETE'), false);
});

test('lost submission acknowledgments recover by ID without retransmitting a MAX request or BYO key', async t => {
  t.mock.timers.enable({ apis: ['setTimeout', 'Date'] });
  const submitted: Array<{ id: string; body: RequestInit['body'] }> = [];
  let admitted = 0;
  t.mock.method(globalThis, 'fetch', async (url: unknown, options: RequestInit = {}) => {
    if (options.method === 'POST') {
      submitted.push({ id: new Headers(options.headers).get('Idempotency-Key')!, body: options.body });
      admitted++;
      throw new TypeError('Lost acknowledgment');
    }
    return String(url).endsWith('/result') ? json({ output_text: 'complete' })
      : json(state(submitted[0].id, 'succeeded'));
  });
  const result = await pump(t, fetchAIJob(init()));
  assert.equal(result.status, 200);
  assert.equal(submitted.length, 1);
  assert.deepEqual(JSON.parse(String(submitted[0].body)).body.reasoning, { effort: 'max' });
  assert.equal(admitted, 1);
});

test('temporary poll failures do not submit fresh inference and terminal provider errors preserve diagnostics', async t => {
  t.mock.timers.enable({ apis: ['setTimeout', 'Date'] });
  let id = '';
  let posts = 0;
  let polls = 0;
  t.mock.method(globalThis, 'fetch', async (url: unknown, options: RequestInit = {}) => {
    if (options.method === 'POST') {
      posts++;
      id = new Headers(options.headers).get('Idempotency-Key')!;
      return json(state(id, 'running'), 202);
    }
    if (String(url).endsWith('/result')) return json({
      error: { source: 'byo', code: 'byo_unavailable', requestId: id },
    }, 503);
    if (++polls === 1) throw new TypeError('Offline');
    if (polls === 2) return json({ error: { code: 'ai_jobs_unavailable' } }, 503);
    return json(state(id, 'failed'));
  });
  const result = await pump(t, fetchAIJob(init()));
  assert.equal(result.status, 503);
  assert.equal((await result.json()).error.code, 'byo_unavailable');
  assert.equal(posts, 1);
  assert.equal(polls, 3);
});

test('abort with a lost POST response cancels the precomputed job ID using a credential-free request', async t => {
  const requests: Array<{ url: string; init: RequestInit }> = [];
  t.mock.method(globalThis, 'fetch', async (url: unknown, options: RequestInit = {}) => {
    requests.push({ url: String(url), init: options });
    return options.method === 'POST' ? new Promise<Response>(() => {}) : json({});
  });
  const controller = new AbortController();
  const work = fetchAIJob(init(), { signal: controller.signal });
  controller.abort();
  await assert.rejects(work, { name: 'AbortError' });
  assert.equal(requests.length, 2);
  const id = new Headers(requests[0].init.headers).get('Idempotency-Key');
  assert.equal(requests[0].init.signal?.aborted, true);
  assert.equal(requests[1].url, `/api/openai/jobs/${id}`);
  assert.equal(requests[1].init.method, 'DELETE');
  assert.equal(requests[1].init.body, undefined);
  assert.equal(requests[1].init.signal?.aborted, false);
});

test('wrong job IDs and untrusted status URLs fail closed instead of following server-supplied locations', async t => {
  const requests: string[] = [];
  t.mock.method(globalThis, 'fetch', async (url: unknown, options: RequestInit = {}) => {
    requests.push(String(url));
    return options.method === 'DELETE' ? json({})
      : json({ ...state('wrong-owner-job-id', 'running'), statusUrl: 'https://untrusted.example/steal' }, 202);
  });
  const response = await fetchAIJob(init());
  assert.equal((await response.json()).error.code, 'ai_job_invalid_response');
  assert.ok(requests.every(value => value.startsWith('/api/openai')));
  assert.equal(requests.length, 2);
});

test('connection currentness is checked before any submission or cancellation is sent', async t => {
  t.mock.timers.enable({ apis: ['setTimeout', 'Date'] });
  let posts = 0;
  let cancels = 0;
  t.mock.method(globalThis, 'fetch', async (_url: unknown, options: RequestInit = {}) => {
    if (options.method === 'DELETE') { cancels++; return json({}); }
    posts++;
    throw new TypeError('Lost response');
  });
  const work = fetchAIJob(init(), { beforeSubmit: () => {
    throw new Error('The captured connection is stale');
  } });
  await assert.rejects(pump(t, work), /captured connection is stale/);
  assert.equal(posts, 0);
  assert.equal(cancels, 0);
});

test('rollback to a server without jobs never replays an ambiguous inference', async t => {
  let posts = 0;
  let cancels = 0;
  t.mock.method(globalThis, 'fetch', async (_url: unknown, options: RequestInit = {}) => {
    if (options.method === 'POST') { posts++; throw new TypeError('Lost legacy response'); }
    if (options.method === 'DELETE') { cancels++; return json({}); }
    return new Response('Cannot GET this route', { status: 404 });
  });
  const result = await fetchAIJob(init());
  assert.equal(result.status, 404);
  assert.equal(posts, 1);
  assert.equal(cancels, 1);
});

test('explicit synchronous rejection and rolling-release success are not retried as jobs', async t => {
  for (const response of [json({ output_text: 'compatible' }), json({ error: { code: 'byo_unavailable' } }, 503)]) {
    let calls = 0;
    t.mock.method(globalThis, 'fetch', async () => { calls++; return response; });
    const result = await fetchAIJob(init());
    assert.equal(result.status, response.status);
    assert.equal(calls, 1);
  }
});
