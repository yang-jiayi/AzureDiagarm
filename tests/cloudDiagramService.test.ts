import test, { afterEach } from 'node:test';
import assert from 'node:assert/strict';
import {
  CloudDiagramApiError,
  getCloudDiagram,
  getCloudDiagramRetryDelay,
} from '../src/services/cloudDiagramService.ts';

const originalFetch = globalThis.fetch;

afterEach(() => {
  globalThis.fetch = originalFetch;
});

test('cloud requests expose Retry-After guidance from throttled responses', async () => {
  globalThis.fetch = (async () => new Response(
    JSON.stringify({ error: 'Please retry later.' }),
    {
      status: 429,
      headers: {
        'Content-Type': 'application/json',
        'Retry-After': '3',
      },
    },
  )) as typeof fetch;

  await assert.rejects(
    getCloudDiagram('diagram-a'),
    (error: unknown) => {
      assert.ok(error instanceof CloudDiagramApiError);
      assert.equal(error.status, 429);
      assert.equal(error.retryAfterMs, 3_000);
      assert.equal(getCloudDiagramRetryDelay(error, 15_000), 3_000);
      return true;
    },
  );
});

test('cloud requests abort with a structured timeout error', async () => {
  globalThis.fetch = ((_input, init) => new Promise((_resolve, reject) => {
    init?.signal?.addEventListener('abort', () => {
      reject(new DOMException('Aborted', 'AbortError'));
    }, { once: true });
  })) as typeof fetch;

  await assert.rejects(
    getCloudDiagram('diagram-a', { timeoutMs: 10 }),
    (error: unknown) => {
      assert.ok(error instanceof CloudDiagramApiError);
      assert.equal(error.status, 408);
      assert.equal(error.code, 'REQUEST_TIMEOUT');
      return true;
    },
  );
});

test('external cancellation remains distinct from a request timeout', async () => {
  globalThis.fetch = ((_input, init) => new Promise((_resolve, reject) => {
    init?.signal?.addEventListener('abort', () => {
      reject(new DOMException('Aborted', 'AbortError'));
    }, { once: true });
  })) as typeof fetch;
  const controller = new AbortController();
  const pending = getCloudDiagram('diagram-a', {
    signal: controller.signal,
    timeoutMs: 1_000,
  });
  controller.abort();

  await assert.rejects(
    pending,
    (error: unknown) => {
      assert.ok(error instanceof CloudDiagramApiError);
      assert.equal(error.status, 499);
      assert.equal(error.code, 'REQUEST_ABORTED');
      return true;
    },
  );
});
