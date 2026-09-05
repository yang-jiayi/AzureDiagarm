import test from 'node:test';
import assert from 'node:assert/strict';
import { setImmediate } from 'node:timers/promises';
import {
  AIBudgetQueueError,
  isAIConcurrencyLimitError,
  runAIBudgetQueue,
  validateAIConcurrencyBudget,
  type AIQueueTaskState,
} from '../src/services/aiBudgetQueue';

function deferred<T>() {
  let resolve!: (value: T) => void;
  let reject!: (reason: unknown) => void;
  const promise = new Promise<T>((yes, no) => { resolve = yes; reject = no; });
  return { promise, resolve, reject };
}

async function waitUntil(predicate: () => boolean): Promise<void> {
  for (let attempt = 0; attempt < 100; attempt += 1) {
    if (predicate()) return;
    await setImmediate();
  }
  assert.ok(predicate(), 'queue did not reach the expected dispatch state');
}

const abortError = (error: unknown) => (
  error instanceof Error && error.name === 'AbortError'
  && 'userCancelled' in error && error.userCancelled === true
);
const contentionError = () => Object.assign(new Error('Capacity is occupied.'), {
  code: 'ai_concurrency_limit', status: 429,
});

for (const limit of [1, 2]) {
  test(`five actual model dispatches stay within server cap ${limit} and preserve result order`, async () => {
    let inFlight = 0;
    let peak = 0;
    const dispatched: number[] = [];
    const gates = Array.from({ length: 5 }, () => deferred<string>());
    const states: Array<AIQueueTaskState<string>['status']> = [];
    const run = runAIBudgetQueue(gates.map((gate, index) => async () => {
      assert.ok(inFlight < limit, 'the next request would be rejected by the bounded server');
      inFlight += 1;
      peak = Math.max(peak, inFlight);
      dispatched.push(index);
      try { return await gate.promise; }
      finally { inFlight -= 1; }
    }), {
      getBudget: async () => ({ concurrentLimit: limit, concurrentRequests: inFlight }),
      onStateChange: (index, state) => { states[index] = state.status; },
    });
    await waitUntil(() => dispatched.length === limit);
    assert.deepEqual(states.slice(0, limit), Array(limit).fill('running'));
    assert.deepEqual(states.slice(limit), Array(5 - limit).fill('pending'));
    for (let index = 0; index < gates.length; index += 1) {
      gates[index].resolve(`model-${index}`);
      await waitUntil(() => dispatched.length >= Math.min(gates.length, index + 1 + limit));
    }
    assert.deepEqual(await run, gates.map((_, index) => ({ status: 'fulfilled', value: `model-${index}` })));
    assert.equal(peak, limit);
    assert.equal(inFlight, 0);
    assert.deepEqual(dispatched, [0, 1, 2, 3, 4]);
    assert.deepEqual(states, Array(5).fill('success'));
  });
}

test('other callers occupying a slot are included in the admission budget', async () => {
  let external = 1;
  let inFlight = 0;
  const started: number[] = [];
  const gates = Array.from({ length: 4 }, () => deferred<number>());
  const run = runAIBudgetQueue(gates.map((gate, index) => async () => {
    assert.ok(inFlight + external < 2);
    inFlight += 1;
    started.push(index);
    try { return await gate.promise; }
    finally { inFlight -= 1; }
  }), { getBudget: async () => ({ concurrentLimit: 2, concurrentRequests: inFlight + external }) });
  await waitUntil(() => started.length === 1);
  await setImmediate();
  assert.deepEqual(started, [0]);
  external = 0;
  gates[0].resolve(0);
  await waitUntil(() => started.length === 3);
  gates[1].resolve(1);
  gates[2].resolve(2);
  await waitUntil(() => started.length === 4);
  gates[3].resolve(3);
  assert.equal((await run).filter(result => result.status === 'fulfilled').length, 4);
});

test('a lagging budget snapshot cannot bypass the local in-flight limit', async () => {
  const gates = Array.from({ length: 4 }, () => deferred<number>());
  let inFlight = 0;
  let started = 0;
  const run = runAIBudgetQueue(gates.map(gate => async () => {
    inFlight += 1;
    started += 1;
    assert.ok(inFlight <= 2);
    try { return await gate.promise; }
    finally { inFlight -= 1; }
  }), { getBudget: async () => ({ concurrentLimit: 2, concurrentRequests: 0 }) });
  await waitUntil(() => started === 2);
  await setImmediate();
  assert.equal(started, 2);
  gates[0].resolve(0);
  await waitUntil(() => started === 3);
  gates[1].resolve(1);
  await waitUntil(() => started === 4);
  gates[2].resolve(2);
  gates[3].resolve(3);
  assert.equal((await run).length, 4);
});

test('one model failure is recorded while later models still execute', async () => {
  const failure = new Error('Model failed.');
  const dispatched: number[] = [];
  const states: string[] = [];
  const results = await runAIBudgetQueue([0, 1, 2, 3].map(index => async () => {
    dispatched.push(index);
    if (index === 1) throw failure;
    return index;
  }), {
    getBudget: async () => ({ concurrentLimit: 1, concurrentRequests: 0 }),
    onStateChange: (index, state) => { states[index] = state.status; },
  });
  assert.deepEqual(dispatched, [0, 1, 2, 3]);
  assert.deepEqual(states, ['success', 'error', 'success', 'success']);
  assert.deepEqual(results[1], { status: 'rejected', reason: failure });
  assert.deepEqual(results[3], { status: 'fulfilled', value: 3 });
});

test('full capacity waits and rechecks before any model is dispatched', async () => {
  let budgetReads = 0;
  let dispatches = 0;
  const results = await runAIBudgetQueue([0, 1, 2].map(index => async () => {
    assert.ok(budgetReads >= 3);
    dispatches += 1;
    return index;
  }), {
    getBudget: async () => ({
      concurrentLimit: 2, concurrentRequests: ++budgetReads < 3 ? 2 : 0,
    }),
    pollIntervalMs: 1,
  });
  assert.equal(dispatches, 3);
  assert.equal(results.length, 3);
});

test('an admission race is requeued with backoff, not reported as model success or failure', async () => {
  let attempts = 0;
  const states: string[] = [];
  const results = await runAIBudgetQueue([
    async () => {
      if (++attempts === 1) throw contentionError();
      return 'model';
    },
  ], {
    getBudget: async () => ({ concurrentLimit: 1, concurrentRequests: 0 }),
    pollIntervalMs: 1,
    onStateChange: (_, state) => { states.push(state.status); },
  });
  assert.equal(attempts, 2);
  assert.deepEqual(states, ['pending', 'running', 'pending', 'running', 'success']);
  assert.deepEqual(results, [{ status: 'fulfilled', value: 'model' }]);
});

test('repeated admission races are bounded and do not prevent another model succeeding', async () => {
  let attempts = 0;
  const results = await runAIBudgetQueue([
    async () => { attempts += 1; throw contentionError(); },
    async () => 'other model',
  ], {
    getBudget: async () => ({ concurrentLimit: 1, concurrentRequests: 0 }),
    pollIntervalMs: 1,
    maxContentionRetries: 2,
  });
  assert.equal(attempts, 3);
  assert.equal(results[0].status, 'rejected');
  if (results[0].status === 'rejected') {
    assert.ok(results[0].reason instanceof AIBudgetQueueError);
    assert.equal(results[0].reason.code, 'contention_limit');
  }
  assert.deepEqual(results[1], { status: 'fulfilled', value: 'other model' });
});

test('ordinary rate limits and daily-budget failures are not mistaken for concurrency contention', async () => {
  let attempts = 0;
  const error = Object.assign(new Error('Daily budget exhausted.'), { code: 'ai_daily_budget_exceeded', status: 429 });
  const results = await runAIBudgetQueue([async () => { attempts += 1; throw error; }], {
    getBudget: async () => ({ concurrentLimit: 2, concurrentRequests: 0 }),
  });
  assert.equal(attempts, 1);
  assert.deepEqual(results, [{ status: 'rejected', reason: error }]);
  assert.equal(isAIConcurrencyLimitError({ status: 429 }), false);
  assert.equal(isAIConcurrencyLimitError({ code: 'proxy_rate_limit_exceeded' }), false);
});

test('invalid budgets fail closed without dispatching any models', async () => {
  for (const value of [
    null, [], {},
    ...[0, -1, 1.5, NaN, Infinity, Number.MAX_SAFE_INTEGER + 1, '2'].map(concurrentLimit => (
      { concurrentLimit, concurrentRequests: 0 }
    )),
    ...[-1, 0.5, NaN, Infinity, Number.MAX_SAFE_INTEGER + 1, '0'].map(concurrentRequests => (
      { concurrentLimit: 2, concurrentRequests }
    )),
  ]) {
    let dispatched = 0;
    await assert.rejects(runAIBudgetQueue([async () => { dispatched += 1; }], {
      getBudget: async () => value,
    }), { name: 'AIBudgetQueueError', code: 'invalid_budget' });
    assert.equal(dispatched, 0);
  }
  assert.deepEqual(validateAIConcurrencyBudget({ concurrentLimit: 1, concurrentRequests: 2 }), {
    concurrentLimit: 1, concurrentRequests: 2,
  }, 'an overcommitted snapshot means wait, not permission to exceed the cap');
});

test('budget-service failures remain explicit and retain their diagnostic cause', async () => {
  const failure = new Error('Budget service offline.');
  let dispatches = 0;
  await assert.rejects(runAIBudgetQueue([async () => { dispatches += 1; }], {
    getBudget: async () => { throw failure; },
  }), (error: unknown) => {
    assert.ok(error instanceof AIBudgetQueueError);
    assert.equal(error.code, 'budget_unavailable');
    assert.equal(error.cause, failure);
    return true;
  });
  assert.equal(dispatches, 0);
});

test('a budget outage after partial progress preserves completed work and stops queued dispatches', async () => {
  let reads = 0;
  const dispatched: number[] = [];
  const states: string[] = [];
  await assert.rejects(runAIBudgetQueue([0, 1, 2].map(index => async () => {
    dispatched.push(index);
    return index;
  }), {
    getBudget: async () => {
      if (++reads > 1) throw new Error('Budget service unavailable.');
      return { concurrentLimit: 1, concurrentRequests: 0 };
    },
    onStateChange: (index, state) => { states[index] = state.status; },
  }), { code: 'budget_unavailable' });
  assert.deepEqual(dispatched, [0]);
  assert.deepEqual(states, ['success', 'pending', 'pending']);
});

test('a budget outage aborts in-flight work and never publishes its late success', async () => {
  const gate = deferred<string>();
  let reads = 0;
  let runningSignal: AbortSignal | undefined;
  let dispatches = 0;
  const terminal: string[] = [];
  const run = runAIBudgetQueue([0, 1, 2].map(() => async signal => {
    dispatches += 1;
    runningSignal = signal;
    return gate.promise;
  }), {
    getBudget: async () => {
      if (++reads > 1) throw new Error('Budget service unavailable.');
      return { concurrentLimit: 2, concurrentRequests: 1 };
    },
    onStateChange: (_, state) => {
      if (state.status === 'success' || state.status === 'error') terminal.push(state.status);
    },
  });
  await assert.rejects(run, { code: 'budget_unavailable' });
  assert.equal(dispatches, 1);
  assert.equal(runningSignal?.aborted, true);
  gate.resolve('late success');
  await setImmediate();
  assert.deepEqual(terminal, []);
});

test('long-running admitted work does not cause blind polling or expire queued models', async t => {
  t.mock.timers.enable({ apis: ['setTimeout'] });
  const gates = Array.from({ length: 3 }, () => deferred<number>());
  const signals: AbortSignal[] = [];
  let reads = 0;
  const run = runAIBudgetQueue(gates.map(gate => async signal => {
    signals.push(signal);
    return gate.promise;
  }), {
    getBudget: async () => { reads += 1; return { concurrentLimit: 1, concurrentRequests: 0 }; },
    maxCapacityWaitMs: 10,
    pollIntervalMs: 1,
  });
  await waitUntil(() => signals.length === 1);
  t.mock.timers.tick(100_000);
  await setImmediate();
  assert.equal(reads, 1);
  assert.equal(signals[0].aborted, false);
  for (let index = 0; index < gates.length; index += 1) {
    gates[index].resolve(index);
    await waitUntil(() => signals.length >= Math.min(gates.length, index + 2));
  }
  assert.equal((await run).length, 3);
});

test('independent queues recover from shared-server admission races without exceeding its cap', async () => {
  const gates = Array.from({ length: 6 }, () => deferred<number>());
  let inFlight = 0;
  let accepted = 0;
  let denied = 0;
  const transport = (index: number) => async () => {
    if (inFlight >= 2) {
      denied += 1;
      throw contentionError();
    }
    inFlight += 1;
    accepted += 1;
    assert.ok(inFlight <= 2);
    try { return await gates[index].promise; }
    finally { inFlight -= 1; }
  };
  const options = {
    getBudget: async () => ({ concurrentLimit: 2, concurrentRequests: inFlight }),
    pollIntervalMs: 1,
  };
  const first = runAIBudgetQueue([0, 1, 2].map(transport), options);
  const second = runAIBudgetQueue([3, 4, 5].map(transport), options);
  await waitUntil(() => denied > 0);
  assert.equal(accepted, 2);
  gates.forEach((gate, index) => gate.resolve(index));
  const results = (await Promise.all([first, second])).flat();
  assert.equal(results.filter(result => result.status === 'fulfilled').length, 6);
  assert.equal(accepted, 6);
  assert.equal(inFlight, 0);
});

test('waiting for external capacity has a finite deadline', async t => {
  t.mock.timers.enable({ apis: ['setTimeout'] });
  let reads = 0;
  const run = runAIBudgetQueue([async () => assert.fail('must not dispatch')], {
    getBudget: async () => { reads += 1; return { concurrentLimit: 2, concurrentRequests: 2 }; },
    pollIntervalMs: 10,
    maxCapacityWaitMs: 30,
  });
  const rejection = assert.rejects(run, { code: 'capacity_timeout' });
  await waitUntil(() => reads === 1);
  await setImmediate();
  t.mock.timers.tick(30);
  await rejection;
  const stoppedAt = reads;
  t.mock.timers.tick(1000);
  await setImmediate();
  assert.equal(reads, stoppedAt);
});

test('a nonresponsive budget endpoint times out and is aborted', async t => {
  t.mock.timers.enable({ apis: ['setTimeout'] });
  let budgetSignal: AbortSignal | undefined;
  const run = runAIBudgetQueue([async () => assert.fail('must not dispatch')], {
    getBudget: signal => { budgetSignal = signal; return new Promise(() => {}); },
    budgetRequestTimeoutMs: 10,
  });
  const rejection = assert.rejects(run, { code: 'budget_unavailable' });
  await waitUntil(() => !!budgetSignal);
  t.mock.timers.tick(10);
  await rejection;
  assert.equal(budgetSignal?.aborted, true);
});

for (const limit of [1, 2]) {
  test(`cancellation at cap ${limit} stops queued work and suppresses late success/error events`, async () => {
    const controller = new AbortController();
    const gates = Array.from({ length: 5 }, () => deferred<number>());
    const signals: AbortSignal[] = [];
    const terminal: string[] = [];
    const run = runAIBudgetQueue(gates.map(gate => async signal => {
      signals.push(signal);
      return gate.promise;
    }), {
      signal: controller.signal,
      getBudget: async () => ({ concurrentLimit: limit, concurrentRequests: 0 }),
      onStateChange: (_, state) => {
        if (state.status === 'success' || state.status === 'error') terminal.push(state.status);
      },
    });
    const rejection = assert.rejects(run, abortError);
    await waitUntil(() => signals.length === limit);
    controller.abort();
    await rejection;
    assert.ok(signals.every(signal => signal.aborted));
    gates[0].resolve(0);
    if (limit === 2) gates[1].reject(new Error('Late provider failure.'));
    await setImmediate();
    assert.equal(signals.length, limit);
    assert.deepEqual(terminal, []);
  });
}

test('cancellation during budget lookup does not dispatch after a late budget response', async () => {
  const controller = new AbortController();
  const budget = deferred<unknown>();
  let signal: AbortSignal | undefined;
  let dispatches = 0;
  const run = runAIBudgetQueue([async () => { dispatches += 1; }], {
    signal: controller.signal,
    getBudget: current => { signal = current; return budget.promise; },
  });
  const rejection = assert.rejects(run, abortError);
  await waitUntil(() => !!signal);
  controller.abort();
  await rejection;
  assert.equal(signal?.aborted, true);
  budget.resolve({ concurrentLimit: 2, concurrentRequests: 0 });
  await setImmediate();
  assert.equal(dispatches, 0);
});

test('already-cancelled queues perform no reads, dispatches, or status updates', async () => {
  const controller = new AbortController();
  controller.abort();
  await assert.rejects(runAIBudgetQueue([async () => assert.fail('dispatch')], {
    signal: controller.signal,
    getBudget: async () => assert.fail('budget lookup'),
    onStateChange: () => assert.fail('status update'),
  }), abortError);
});

test('cancellation from a running-state callback prevents even that transport dispatch', async () => {
  const controller = new AbortController();
  let dispatches = 0;
  await assert.rejects(runAIBudgetQueue([0, 1, 2].map(() => async () => {
    dispatches += 1;
  }), {
    signal: controller.signal,
    getBudget: async () => ({ concurrentLimit: 2, concurrentRequests: 0 }),
    onStateChange: (_, state) => { if (state.status === 'running') controller.abort(); },
  }), abortError);
  assert.equal(dispatches, 0);
});

test('normal completion removes the external cancellation listener', async t => {
  const controller = new AbortController();
  const added = t.mock.method(controller.signal, 'addEventListener');
  const removed = t.mock.method(controller.signal, 'removeEventListener');
  await runAIBudgetQueue([async () => 'done'], {
    signal: controller.signal,
    getBudget: async () => ({ concurrentLimit: 1, concurrentRequests: 0 }),
  });
  assert.equal(added.mock.callCount(), 1);
  assert.equal(removed.mock.callCount(), 1);
  assert.equal(added.mock.calls[0].arguments[1], removed.mock.calls[0].arguments[1]);
});

test('cancellation during capacity waiting removes timers and stops budget polling', async t => {
  t.mock.timers.enable({ apis: ['setTimeout'] });
  const controller = new AbortController();
  let reads = 0;
  const run = runAIBudgetQueue([async () => assert.fail('dispatch')], {
    signal: controller.signal,
    getBudget: async () => { reads += 1; return { concurrentLimit: 1, concurrentRequests: 1 }; },
    pollIntervalMs: 10,
  });
  const rejection = assert.rejects(run, abortError);
  await waitUntil(() => reads === 1);
  controller.abort();
  await rejection;
  t.mock.timers.tick(100_000);
  await setImmediate();
  assert.equal(reads, 1);
});

test('empty queues do not read the budget and invalid timer configuration is explicit', async () => {
  assert.deepEqual(await runAIBudgetQueue([], { getBudget: async () => assert.fail('budget lookup') }), []);
  for (const config of [
    { pollIntervalMs: 0 }, { maxCapacityWaitMs: -1 }, { budgetRequestTimeoutMs: Infinity },
    { pollIntervalMs: 2_147_483_648 }, { maxContentionRetries: 0.5 },
  ]) {
    await assert.rejects(runAIBudgetQueue([], {
      getBudget: async () => assert.fail('budget lookup'), ...config,
    }), { code: 'invalid_configuration' });
  }
});
