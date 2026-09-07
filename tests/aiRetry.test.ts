import test from 'node:test';
import assert from 'node:assert/strict';
import type { RuntimeModelOverride } from '../src/services/aiModelRuntime.ts';
import {
  isRetryableAIFailure,
  isAIRateLimitError,
  runWithCompactRetry,
  safeParseModelJson,
  ModelJsonError,
  ModelJsonErrorKind,
  MODEL_JSON_ERROR_MESSAGES,
} from '../src/services/aiRetry.ts';

/** Console stub so the intentional parse diagnostics don't spam the test log. */
const silentLogger = { error() {}, warn() {} };

test('proxy timeout codes are retryable', () => {
  for (const code of ['azure_openai_timeout', 'edge_origin_unavailable']) {
    assert.equal(isRetryableAIFailure(Object.assign(new Error('boom'), { code })), true, code);
  }
});

test('BYO throttles and transient failures share managed retry classification without fallback', () => {
  for (const code of ['byo_timeout', 'byo_unavailable', 'byo_connection_failed', 'byo_rate_limited']) {
    const error = Object.assign(new Error('Retired provider'), { code });
    assert.equal(isAIRateLimitError(error), code === 'byo_rate_limited', code);
    assert.equal(isRetryableAIFailure(error), true, code);
  }
});

test('proxy configuration failures are not transient provider failures despite HTTP 503', () => {
  for (const code of ['astra_not_configured', 'proxy_not_configured']) {
    const error = Object.assign(new Error('The proxy is not configured.'), {
      code, source: 'proxy', status: 503,
    });
    assert.equal(isAIRateLimitError(error), false, code);
    assert.equal(isRetryableAIFailure(error), false, code);
  }
});

test('abort and timeout errors are retryable, but a user cancellation is not', () => {
  // Internal timeouts can be retried explicitly, but never at lower quality.
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

test('capacity contention cannot trigger a compact retry, including a custom payload classifier', async () => {
  const error = Object.assign(new Error('Wait for AI capacity.'), {
    code: 'ai_concurrency_limit', status: 429,
  });
  assert.equal(isRetryableAIFailure(error), false);
  let attempts = 0;
  await assert.rejects(runWithCompactRetry({
    transportFeature: 'architectureGeneration',
    label: 'Comparison generation',
    isRetryable: () => true,
    attempt: async () => { attempts += 1; throw error; },
  }), (actual: unknown) => actual === error);
  assert.equal(attempts, 1);
  assert.equal('retried' in error, false);
});

for (const code of ['azure_openai_rate_limited', 'byo_rate_limited', 'proxy_rate_limit_exceeded', 'http_429']) {
  test(`${code} never triggers a compact quality downgrade`, async () => {
    const error = Object.assign(new Error('Rate limit reached.'), { code, status: 429, retryAfterMs: 60_000 });
    const attempts: Array<{ compact: boolean; override?: RuntimeModelOverride }> = [];
    const requested: RuntimeModelOverride = { model: 'gpt-6-astra', reasoningEffort: 'max' };
    await assert.rejects(runWithCompactRetry({
      transportFeature: 'architectureGeneration',
      override: requested,
      label: 'Astra MAX generation',
      isRetryable: () => true,
      attempt: async (compact, override) => { attempts.push({ compact, override }); throw error; },
    }), (actual: unknown) => actual === error);
    assert.deepEqual(attempts, [{ compact: false, override: requested }]);
  });
}

for (const [code, status] of [
  ['ai_daily_budget_exceeded', 429], ['ai_budget_busy', 503],
  ['ai_budget_unavailable', 503], ['ai_budget_timeout', 504],
] as const) {
  test(`${code} never becomes a compact retry even with a permissive payload classifier`, async () => {
    const error = Object.assign(new Error('Application budget rejected the request.'), { source: 'budget', code, status });
    let attempts = 0;
    assert.equal(isRetryableAIFailure(error), false);
    await assert.rejects(runWithCompactRetry({
      transportFeature: 'architectureGeneration', label: 'Budget rejection',
      isRetryable: () => true,
      attempt: async () => { attempts++; throw error; },
    }), actual => actual === error);
    assert.equal(attempts, 1);
  });
}

for (const status of [500, 502]) {
  test(`a provider 500 carried by HTTP ${status} is not automatically replayed as a lower-quality compact request`, async () => {
    const error = Object.assign(new Error('Provider internal error.'), {
      source: 'azure_openai', code: 'azure_openai_unavailable', status, upstreamStatus: 500, upstreamCode: 'server_error',
    });
    const attempts: Array<{ compact: boolean; override?: RuntimeModelOverride }> = [];
    const requested: RuntimeModelOverride = { model: 'gpt-6-astra', reasoningEffort: 'max' };
    await assert.rejects(runWithCompactRetry({
      transportFeature: 'architectureGeneration', override: requested, label: 'Unknown-usage 500',
      attempt: async (compact, override) => { attempts.push({ compact, override }); throw error; },
    }), actual => actual === error);
    assert.deepEqual(attempts, [{ compact: false, override: requested }]);
  });
}

test('the user-visible timeout message is recognised', () => {
  assert.equal(
    isRetryableAIFailure(new Error('The AI provider is taking too long to respond.')),
    true,
  );
  assert.equal(isRetryableAIFailure(new Error('Invalid API key')), false);
  assert.equal(isRetryableAIFailure(null), false);
  assert.equal(isRetryableAIFailure('a string'), false);
});

const OVERRIDE: RuntimeModelOverride = { model: 'gpt-6-astra', reasoningEffort: 'high' };

test('a timeout preserves requested quality and requires an explicit user retry', async () => {
  const calls: Array<{ compact: boolean; override?: RuntimeModelOverride }> = [];
  const failure = Object.assign(new Error('slow'), { code: 'azure_openai_timeout' });
  await assert.rejects(runWithCompactRetry({
    transportFeature: 'architectureGeneration',
    override: OVERRIDE,
    label: 'Blueprint generation',
    attempt: async (compact, override) => {
      calls.push({ compact, override });
      if (!compact) throw failure;
      return 'compact-result';
    },
  }), actual => actual === failure);

  assert.equal(calls.length, 1, 'no hidden lower-quality replay');
  assert.equal(calls[0].compact, false);
  assert.equal(calls[0].override, OVERRIDE, 'the first attempt uses the caller override verbatim');
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

test('payload classifiers cannot authorize a silent quality downgrade', async () => {
  class ResponseError extends Error {}
  let attempts = 0;
  const failure = new ResponseError('truncated JSON');

  await assert.rejects(runWithCompactRetry({
    transportFeature: 'architectureGeneration',
    override: OVERRIDE,
    label: 'Blueprint generation',
    isRetryable: (error) => error instanceof ResponseError,
    attempt: async (compact) => {
      attempts += 1;
      if (!compact) throw failure;
      return 'ok';
    },
  }), actual => actual === failure);

  assert.equal(attempts, 1);
});

test('terminal failures preserve the typed error without claiming an automatic replay', async () => {
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
      assert.equal((error as { retried?: boolean }).retried, undefined, 'no retry is claimed');
      return true;
    },
  );
  assert.equal(attempts, 1, 'no additional generation is dispatched automatically');
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
