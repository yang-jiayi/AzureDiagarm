import test from 'node:test';
import assert from 'node:assert/strict';
import { isAIRateLimitError, runWithRateLimitRetry, type AIRetryWait } from '../src/services/aiRetry';

const flush = () => new Promise<void>(resolve => setImmediate(resolve));
const limited = (retryAfterMs?: number) => Object.assign(new Error('Provider rate limit.'), {
  status: 429, code: 'azure_openai_rate_limited', retryAfterMs,
});

test('rate-limit recovery stops after three dispatches and 120 seconds without rewriting the error', async t => {
  t.mock.timers.enable({ apis: ['setTimeout'] });
  const error = Object.freeze(limited(60_000));
  let calls = 0;
  const waits: Array<AIRetryWait | null> = [];
  const request = runWithRateLimitRetry(async () => { calls++; throw error; }, {
    onRetryWait: wait => waits.push(wait),
  });
  const rejection = assert.rejects(request, actual => actual === error);
  await flush();
  assert.equal(calls, 1);
  for (let attempt = 2; attempt <= 3; attempt++) {
    t.mock.timers.tick(59_999);
    await flush();
    assert.equal(calls, attempt - 1);
    t.mock.timers.tick(1);
    await flush();
    assert.equal(calls, attempt);
  }
  await rejection;
  t.mock.timers.tick(1_000_000);
  await flush();
  assert.equal(calls, 3);
  assert.deepEqual(waits.map(wait => wait && [wait.attempt, wait.maxAttempts, wait.delayMs]), [
    [2, 3, 60_000], null, [3, 3, 60_000], null,
  ]);
});

test('a provider cooldown longer than the remaining allowance is surfaced, never shortened', async t => {
  t.mock.timers.enable({ apis: ['setTimeout'] });
  for (const delay of [120_001, 3_600_000, Number.MAX_SAFE_INTEGER]) {
    const error = limited(delay);
    let calls = 0;
    let notifications = 0;
    await assert.rejects(runWithRateLimitRetry(async () => { calls++; throw error; }, {
      onRetryWait: () => { notifications++; },
    }), actual => actual === error);
    t.mock.timers.tick(1_000_000);
    assert.equal(calls, 1);
    assert.equal(notifications, 0);
  }
  let calls = 0;
  const error = limited(70_000);
  const request = runWithRateLimitRetry(async () => { calls++; throw error; });
  const rejection = assert.rejects(request, actual => actual === error);
  await flush();
  t.mock.timers.tick(70_000);
  await rejection;
  assert.equal(calls, 2, 'a second 70-second cooldown would exceed the 120-second limit');
});

for (const delay of [undefined, -1, NaN, Infinity]) {
  test(`missing/invalid Retry-After ${delay} uses bounded 30s/60s backoff`, async t => {
    t.mock.timers.enable({ apis: ['setTimeout'] });
    let calls = 0;
    const request = runWithRateLimitRetry(async () => {
      calls++;
      if (calls < 3) throw limited(delay);
      return 'completed';
    });
    await flush();
    t.mock.timers.tick(29_999);
    await flush();
    assert.equal(calls, 1);
    t.mock.timers.tick(1);
    await flush();
    assert.equal(calls, 2);
    t.mock.timers.tick(59_999);
    await flush();
    assert.equal(calls, 2);
    t.mock.timers.tick(1);
    assert.equal(await request, 'completed');
    assert.equal(calls, 3);
  });
}

test('a zero Retry-After still backs off rather than immediately replaying', async t => {
  t.mock.timers.enable({ apis: ['setTimeout'] });
  let calls = 0;
  const request = runWithRateLimitRetry(async () => {
    if (++calls === 1) throw limited(0);
    return 'completed';
  });
  await flush();
  t.mock.timers.tick(999);
  await flush();
  assert.equal(calls, 1);
  t.mock.timers.tick(1);
  assert.equal(await request, 'completed');
});

test('cancelling during provider waiting clears the timer and never reports resumption or success', async t => {
  t.mock.timers.enable({ apis: ['setTimeout'] });
  const controller = new AbortController();
  const waits: Array<AIRetryWait | null> = [];
  const add = t.mock.method(controller.signal, 'addEventListener');
  const remove = t.mock.method(controller.signal, 'removeEventListener');
  let calls = 0;
  const request = runWithRateLimitRetry(async () => { calls++; throw limited(60_000); }, {
    signal: controller.signal, onRetryWait: wait => waits.push(wait),
  });
  const rejection = assert.rejects(request, { name: 'AbortError', userCancelled: true });
  await flush();
  controller.abort();
  await rejection;
  t.mock.timers.tick(1_000_000);
  await flush();
  assert.equal(calls, 1);
  assert.equal(waits.length, 1);
  assert.equal(add.mock.callCount(), remove.mock.callCount());
});

test('already-cancelled requests dispatch nothing, even with a custom abort reason', async () => {
  const controller = new AbortController();
  controller.abort(new Error('User cancelled.'));
  let calls = 0;
  await assert.rejects(runWithRateLimitRetry(async () => { calls++; }, { signal: controller.signal }), {
    name: 'AbortError', userCancelled: true,
  });
  assert.equal(calls, 0);
});

test('cancellation from a wait callback cannot queue another dispatch', async t => {
  t.mock.timers.enable({ apis: ['setTimeout'] });
  const controller = new AbortController();
  let calls = 0;
  let notifications = 0;
  await assert.rejects(runWithRateLimitRetry(async () => { calls++; throw limited(60_000); }, {
    signal: controller.signal, onRetryWait: () => { notifications++; controller.abort(); },
  }), { name: 'AbortError', userCancelled: true });
  t.mock.timers.tick(1_000_000);
  assert.equal(calls, 1);
  assert.equal(notifications, 1);
});

test('daily budgets, admission races, auth, configuration and ordinary errors never use provider retry', async () => {
  for (const [code, status] of [
    ['ai_concurrency_limit', 429], ['ai_daily_budget_exceeded', 429],
    ['application_access_denied', 403], ['byo_authentication_failed', 401],
    ['deployment_not_found', 404], ['azure_openai_timeout', 504], ['astra_not_configured', 503],
    ['proxy_not_configured', 503], ['invalid_api_format', 400], ['byo_not_enabled', 403],
    ['deployment_not_allowed', 403],
  ] as const) {
    const error = Object.assign(new Error(code), { code, status });
    assert.equal(isAIRateLimitError(error), false);
    let calls = 0;
    await assert.rejects(runWithRateLimitRetry(async () => { calls++; throw error; }), actual => actual === error);
    assert.equal(calls, 1);
  }
});

test('an unclassified 429 is not assumed to be a retryable provider throttle', async () => {
  for (const code of [undefined, 'http_429', 'unknown_error', '429']) {
    const error = Object.assign(new Error('Unclassified limit.'), { status: 429, code, retryAfterMs: 1000 });
    assert.equal(isAIRateLimitError(error), false);
    let calls = 0;
    await assert.rejects(runWithRateLimitRetry(async () => { calls++; throw error; }), actual => actual === error);
    assert.equal(calls, 1);
  }
});
