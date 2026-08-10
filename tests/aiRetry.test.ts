import test from 'node:test';
import assert from 'node:assert/strict';
import type { RuntimeModelOverride } from '../src/services/aiModelRuntime.ts';
import {
  buildRetryOverride,
  degradeOverrideForRetry,
  downshiftReasoningEffort,
  isRetryableAIFailure,
  runWithCompactRetry,
  safeParseModelJson,
  ModelJsonError,
  ModelJsonErrorKind,
  MODEL_JSON_ERROR_MESSAGES,
} from '../src/services/aiRetry.ts';

/** Console stub so the intentional parse diagnostics don't spam the test log. */
const silentLogger = { error() {}, warn() {} };

test('proxy timeout codes are retryable', () => {
  for (const code of ['azure_openai_timeout', 'byo_timeout', 'edge_origin_unavailable']) {
    assert.equal(isRetryableAIFailure(Object.assign(new Error('boom'), { code })), true, code);
  }
});

test('abort and timeout errors are retryable, but a user cancellation is not', () => {
  // An internal (timeout-driven) abort surfaces as an AbortError and SHOULD be
  // retried — a compact retry often finishes inside the budget.
  const abort = new Error('The operation was aborted');
  abort.name = 'AbortError';
  assert.equal(isRetryableAIFailure(abort), true);

  const timeout = new Error('timed out');
  timeout.name = 'TimeoutError';
  assert.equal(isRetryableAIFailure(timeout), true);

  // A user-initiated cancel ALSO surfaces as an AbortError, so it must be told
  // apart by the explicit `userCancelled` flag — never retried.
  const cancelled = Object.assign(new Error('Generation cancelled.'), {
    name: 'AbortError',
    userCancelled: true,
  });
  assert.equal(isRetryableAIFailure(cancelled), false);

  // The flag wins even when a retryable proxy code is also present.
  const cancelledWithCode = Object.assign(new Error('Generation cancelled.'), {
    name: 'AbortError',
    userCancelled: true,
    code: 'azure_openai_timeout',
  });
  assert.equal(isRetryableAIFailure(cancelledWithCode), false);
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

test('a failing retry preserves the original typed error and flags the retry', async () => {
  let attempts = 0;
  const original = Object.assign(new Error('still too slow'), { code: 'azure_openai_timeout' });
  await assert.rejects(
    runWithCompactRetry({
      transportFeature: 'architectureGeneration',
      override: OVERRIDE,
      label: 'Blueprint generation',
      attempt: async () => {
        attempts += 1;
        throw original;
      },
    }),
    (error: unknown) => {
      // MEDIUM 5: the ORIGINAL error must survive so downstream UI can still
      // classify (`.code`) and localise it — not a flattened generic Error.
      assert.equal(error, original, 'the original typed error instance is rethrown');
      assert.equal((error as { code?: string }).code, 'azure_openai_timeout', 'the code survives');
      assert.equal(error instanceof Error && error.message, 'still too slow', 'the message is not mangled');
      assert.equal((error as { retried?: boolean }).retried, true, 'the retry is annotated');
      return true;
    },
  );
  assert.equal(attempts, 2, 'retries are bounded to one — no infinite loop');
});

// ── safeParseModelJson (HIGH 4) ─────────────────────────────────────────────

test('safeParseModelJson parses plain valid JSON', () => {
  assert.deepEqual(safeParseModelJson('{"services":[{"id":"a"}]}'), { services: [{ id: 'a' }] });
  assert.deepEqual(safeParseModelJson('[1,2,3]'), [1, 2, 3]);
});

test('safeParseModelJson strips ```json fences', () => {
  const raw = '```json\n{"a":1}\n```';
  assert.deepEqual(safeParseModelJson(raw, { logger: silentLogger }), { a: 1 });
});

test('safeParseModelJson strips bare ``` fences (no language tag)', () => {
  const raw = '```\n{"a":1}\n```';
  assert.deepEqual(safeParseModelJson(raw, { logger: silentLogger }), { a: 1 });
});

test('safeParseModelJson handles a language tag other than json', () => {
  const raw = '```JSON\n{"a":1}\n```';
  assert.deepEqual(safeParseModelJson(raw, { logger: silentLogger }), { a: 1 });
});

test('safeParseModelJson extracts JSON after leading prose', () => {
  const raw = 'Sure! Here is the architecture you asked for:\n{"services":["x"]}\nHope that helps.';
  assert.deepEqual(safeParseModelJson(raw, { logger: silentLogger }), { services: ['x'] });
});

test('safeParseModelJson ignores braces inside strings when balancing', () => {
  const raw = 'prefix {"label":"a } b { c","n":1} suffix';
  assert.deepEqual(safeParseModelJson(raw, { logger: silentLogger }), { label: 'a } b { c', n: 1 });
});

test('safeParseModelJson throws a retryable Truncated error for cut-off JSON', () => {
  const raw = '{"services":[{"id":"a"},{"id":"b"';
  assert.throws(
    () => safeParseModelJson(raw, { logger: silentLogger }),
    (error: unknown) => {
      assert.ok(error instanceof ModelJsonError);
      assert.equal(error.kind, ModelJsonErrorKind.Truncated);
      assert.equal(error.retryable, true);
      assert.equal(error.message, MODEL_JSON_ERROR_MESSAGES[ModelJsonErrorKind.Truncated]);
      assert.doesNotMatch(error.message, /Unexpected|token|JSON\.parse|position/i, 'no raw parser detail leaks');
      return true;
    },
  );
  assert.equal(isRetryableAIFailure(new ModelJsonError(ModelJsonErrorKind.Truncated, 'x', { retryable: true })), true);
});

test('safeParseModelJson throws a non-retryable Refusal error for apology prose', () => {
  for (const raw of [
    "I'm sorry, but I can't help with that request.",
    'As an AI, I am unable to assist with generating this architecture.',
    'I cannot comply with this request due to content policy.',
  ]) {
    assert.throws(
      () => safeParseModelJson(raw, { logger: silentLogger }),
      (error: unknown) => {
        assert.ok(error instanceof ModelJsonError, raw);
        assert.equal(error.kind, ModelJsonErrorKind.Refusal, raw);
        assert.equal(error.retryable, false, raw);
        assert.equal(error.message, MODEL_JSON_ERROR_MESSAGES[ModelJsonErrorKind.Refusal]);
        return true;
      },
    );
  }
  assert.equal(isRetryableAIFailure(new ModelJsonError(ModelJsonErrorKind.Refusal, 'x', { retryable: false })), false);
});

// Issue 1 — a refusal that happens to contain a stray bracket must still be
// classified as a non-retryable Refusal, never as retryable Unparseable and
// never silently returned as `{}` / `[1]` data.
test('safeParseModelJson classifies braced/bracketed refusals as non-retryable Refusal', () => {
  for (const raw of [
    "I'm sorry, but I can't help with that request.",
    "I'm sorry, I can't create that. For example {foo}.",
    'As an AI, I am unable to do this. snippet: { }.',
    'I cannot comply. See sources [1] and [2].',
  ]) {
    assert.throws(
      () => safeParseModelJson(raw, { logger: silentLogger }),
      (error: unknown) => {
        assert.ok(error instanceof ModelJsonError, raw);
        assert.equal(error.kind, ModelJsonErrorKind.Refusal, raw);
        assert.equal(error.retryable, false, raw);
        assert.equal(error.message, MODEL_JSON_ERROR_MESSAGES[ModelJsonErrorKind.Refusal]);
        return true;
      },
    );
  }
});

// Issue 2 — bare JSON primitives are valid JSON but useless to every caller;
// they must be rejected as retryable Unparseable instead of being returned as
// `T` and later dereferenced into a raw TypeError.
test('safeParseModelJson rejects bare JSON primitives as retryable Unparseable', () => {
  for (const raw of ['null', '123', '"a string"', 'true']) {
    assert.throws(
      () => safeParseModelJson(raw, { logger: silentLogger }),
      (error: unknown) => {
        assert.ok(error instanceof ModelJsonError, raw);
        assert.equal(error.kind, ModelJsonErrorKind.Unparseable, raw);
        assert.equal(error.retryable, true, raw);
        assert.equal(error.message, MODEL_JSON_ERROR_MESSAGES[ModelJsonErrorKind.Unparseable]);
        return true;
      },
    );
  }
});

test('safeParseModelJson throws a retryable Empty error for blank input', () => {
  for (const raw of ['', '   ', '\n\t ']) {
    assert.throws(
      () => safeParseModelJson(raw, { logger: silentLogger }),
      (error: unknown) => {
        assert.ok(error instanceof ModelJsonError);
        assert.equal(error.kind, ModelJsonErrorKind.Empty);
        assert.equal(error.retryable, true);
        assert.equal(error.message, MODEL_JSON_ERROR_MESSAGES[ModelJsonErrorKind.Empty]);
        return true;
      },
    );
  }
});

test('safeParseModelJson throws Unparseable for non-JSON, non-refusal prose', () => {
  assert.throws(
    () => safeParseModelJson('the quick brown fox jumps over the lazy dog', { logger: silentLogger }),
    (error: unknown) => {
      assert.ok(error instanceof ModelJsonError);
      assert.equal(error.kind, ModelJsonErrorKind.Unparseable);
      return true;
    },
  );
});

test('safeParseModelJson keeps raw parser detail on .detail, never on .message', () => {
  const raw = '{"a": broken,,,}';
  try {
    safeParseModelJson(raw, { logger: silentLogger });
    assert.fail('expected a ModelJsonError');
  } catch (error) {
    assert.ok(error instanceof ModelJsonError);
    // The raw payload is preserved for logging but must not appear in the
    // user-facing message.
    assert.ok(error.detail && error.detail.includes('broken'));
    assert.doesNotMatch(error.message, /broken/);
  }
});

