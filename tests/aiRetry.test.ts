import test from 'node:test';
import assert from 'node:assert/strict';
import type { RuntimeModelOverride } from '../src/services/aiModelRuntime.ts';
import {
  buildRetryOverride,
  degradeOverrideForRetry,
  downshiftReasoningEffort,
  isRetryableAIFailure,
  runWithCompactRetry,
} from '../src/services/aiRetry.ts';

test('proxy timeout codes are retryable', () => {
  for (const code of ['azure_openai_timeout', 'byo_timeout', 'edge_origin_unavailable']) {
    assert.equal(isRetryableAIFailure(Object.assign(new Error('boom'), { code })), true, code);
  }
});

test('abort and timeout errors are retryable', () => {
  const abort = new Error('The operation was aborted');
  abort.name = 'AbortError';
  assert.equal(isRetryableAIFailure(abort), true);

  const timeout = new Error('timed out');
  timeout.name = 'TimeoutError';
  assert.equal(isRetryableAIFailure(timeout), true);
});

test('transient HTTP statuses are retryable but client errors are not', () => {
  for (const status of [429, 502, 503, 504]) {
    assert.equal(isRetryableAIFailure(Object.assign(new Error('http'), { status })), true, String(status));
  }
  for (const status of [400, 401, 403, 404, 422]) {
    assert.equal(isRetryableAIFailure(Object.assign(new Error('http'), { status })), false, String(status));
  }
});

test('the user-visible timeout message is recognised', () => {
  assert.equal(
    isRetryableAIFailure(new Error('The AI provider is taking too long to respond.')),
    true,
  );
  assert.equal(isRetryableAIFailure(new Error('Invalid API key')), false);
  assert.equal(isRetryableAIFailure(null), false);
  assert.equal(isRetryableAIFailure('a string'), false);
});

test('reasoning effort is downshifted only when it is above low', () => {
  assert.equal(downshiftReasoningEffort('xhigh'), 'low');
  assert.equal(downshiftReasoningEffort('high'), 'low');
  assert.equal(downshiftReasoningEffort('medium'), 'low');
  assert.equal(downshiftReasoningEffort('low'), 'low');
  assert.equal(downshiftReasoningEffort('minimal'), 'minimal');
  assert.equal(downshiftReasoningEffort('none'), 'none');
});

test('retry overrides keep the caller model but lower the effort', () => {
  const override = degradeOverrideForRetry({ model: 'gpt-5.6-sol', reasoningEffort: 'high' });
  assert.equal(override?.model, 'gpt-5.6-sol');
  assert.equal(override?.reasoningEffort, 'low');

  // An already-cheap override is returned untouched, and "no override" stays
  // undefined so the feature defaults keep applying.
  const cheap = { model: 'gpt-5.6-sol', reasoningEffort: 'minimal' } as const;
  assert.equal(degradeOverrideForRetry(cheap), cheap);
  assert.equal(degradeOverrideForRetry(undefined), undefined);
});

const OVERRIDE: RuntimeModelOverride = { model: 'gpt-5.6-sol', reasoningEffort: 'high' };

test('a retry override is produced even when the caller passed none', () => {
  const explicit = buildRetryOverride('blueprint', OVERRIDE);
  assert.deepEqual(explicit, { model: 'gpt-5.6-sol', reasoningEffort: 'low' });

  // Falling back to the feature settings must still lower the effort, otherwise
  // the retry repeats exactly the request that timed out.
  const implicit = buildRetryOverride('blueprint', undefined);
  assert.ok(['none', 'minimal', 'low'].includes(implicit.reasoningEffort));
  assert.ok(implicit.model);
});

test('the compact retry runs the second attempt exactly once, cheaply', async () => {
  const calls: Array<{ compact: boolean; override?: RuntimeModelOverride }> = [];
  const result = await runWithCompactRetry({
    transportFeature: 'architectureGeneration',
    override: OVERRIDE,
    label: 'Blueprint generation',
    attempt: async (compact, override) => {
      calls.push({ compact, override });
      if (!compact) throw Object.assign(new Error('slow'), { code: 'azure_openai_timeout' });
      return 'compact-result';
    },
  });

  assert.equal(result, 'compact-result');
  assert.equal(calls.length, 2, 'exactly one retry');
  assert.equal(calls[0].compact, false);
  assert.equal(calls[0].override, OVERRIDE, 'the first attempt uses the caller override verbatim');
  assert.equal(calls[1].compact, true, 'the retry drops the few-shot exemplars');
  assert.equal(calls[1].override?.reasoningEffort, 'low');
  assert.equal(calls[1].override?.model, 'gpt-5.6-sol', 'the user keeps their model');
});

test('non-retryable failures are rethrown untouched and never retried', async () => {
  let attempts = 0;
  const authError = Object.assign(new Error('Invalid API key'), { status: 401 });

  await assert.rejects(
    runWithCompactRetry({
      transportFeature: 'architectureGeneration',
      override: OVERRIDE,
      label: 'Blueprint generation',
      attempt: async () => {
        attempts += 1;
        throw authError;
      },
    }),
    (error: unknown) => error === authError,
  );
  assert.equal(attempts, 1, 'the user must not be charged for a pointless second call');
});

test('payload-level failures opt in through isRetryable', async () => {
  class ResponseError extends Error {}
  let attempts = 0;

  const result = await runWithCompactRetry({
    transportFeature: 'architectureGeneration',
    override: OVERRIDE,
    label: 'Blueprint generation',
    isRetryable: (error) => error instanceof ResponseError,
    attempt: async (compact) => {
      attempts += 1;
      if (!compact) throw new ResponseError('truncated JSON');
      return 'ok';
    },
  });

  assert.equal(result, 'ok');
  assert.equal(attempts, 2);
});

test('a failing retry reports that a retry already happened', async () => {
  let attempts = 0;
  await assert.rejects(
    runWithCompactRetry({
      transportFeature: 'architectureGeneration',
      override: OVERRIDE,
      label: 'Blueprint generation',
      attempt: async () => {
        attempts += 1;
        throw Object.assign(new Error('still too slow'), { code: 'azure_openai_timeout' });
      },
    }),
    (error: unknown) => {
      assert.ok(error instanceof Error);
      assert.match(error.message, /Blueprint generation failed after an automatic retry: still too slow/);
      return true;
    },
  );
  assert.equal(attempts, 2, 'retries are bounded to one — no infinite loop');
});

