const assert = require('node:assert/strict');
const test = require('node:test');
const { cancelErrorResponseBody, readBoundedText } = require('./bounded-response');

function responseFromChunks(chunks, headers = {}) {
  const body = new ReadableStream({
    start(controller) {
      for (const chunk of chunks) {
        controller.enqueue(new TextEncoder().encode(chunk));
      }
      controller.close();
    },
  });
  return new Response(body, { headers });
}

test('bounded response reader returns content up to the byte limit', async () => {
  const response = responseFromChunks(['safe ', 'documentation']);

  assert.equal(await readBoundedText(response, 18), 'safe documentation');
});

test('bounded response reader rejects an oversized streamed body', async () => {
  const response = responseFromChunks(['1234', '5678']);

  await assert.rejects(readBoundedText(response, 7), /Response exceeded its size limit/);
});

test('bounded response reader rejects an oversized declared body before reading', async () => {
  const response = responseFromChunks(['small'], { 'content-length': '8' });

  await assert.rejects(readBoundedText(response, 7), /Response exceeded its size limit/);
});

test('error response cancellation releases the body and logs only status when cancellation fails', async () => {
  let cancelCalled = false;
  const body = new ReadableStream({
    start(controller) {
      controller.enqueue(new TextEncoder().encode('response-body-canary'));
    },
    cancel() {
      cancelCalled = true;
      throw new Error('exception-canary');
    },
  });
  const response = new Response(body, { status: 401 });
  const logs = [];
  const originalError = console.error;
  console.error = (...args) => logs.push(args.join(' '));
  try {
    await cancelErrorResponseBody(response, 'test-operation');
  } finally {
    console.error = originalError;
  }

  assert.equal(cancelCalled, true);
  assert.equal(response.status, 401);
  assert.deepEqual(logs, [
    '[test-operation] failed to cancel upstream error body (upstream_status=401)',
  ]);
  assert.doesNotMatch(logs.join(' '), /response-body-canary|exception-canary/);
});
