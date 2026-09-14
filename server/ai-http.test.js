// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.
const assert = require('node:assert/strict');
const https = require('node:https');
const { EventEmitter } = require('node:events');
const { PassThrough } = require('node:stream');
const { test } = require('node:test');
const { awaitWithSignal, fetchLongAIResponse, MAX_RESPONSE_BYTES } = require('./ai-http');

test('ignored cancellation and late rejection are bounded without unhandled work', async () => {
  let rejectLate;
  const work = new Promise((_resolve, reject) => { rejectLate = reject; });
  const controller = new AbortController();
  const pending = awaitWithSignal(work, controller.signal);
  const reason = new Error('cancelled');
  controller.abort(reason);
  await assert.rejects(pending, error => error === reason);
  rejectLate(new Error('late rejection'));
  await new Promise(resolve => setImmediate(resolve));
});

function mockRequest(t, status, run) {
  t.mock.method(https, 'request', (_url, options, callback) => {
    const request = new EventEmitter();
    const stream = new PassThrough();
    stream.statusCode = status;
    stream.headers = { 'content-type': 'application/json', 'x-request-id': 'native-test' };
    request.end = body => {
      assert.equal(options.headers['Content-Length'], Buffer.byteLength(body));
      request.emit('socket', { setKeepAlive: (enabled, delay) => {
        assert.equal(enabled, true);
        assert.equal(delay, 60000);
      } });
      queueMicrotask(() => { callback(stream); run(stream, request, options); });
    };
    return request;
  });
}

test('native HTTPS consumes bounded bodies, preserves headers, and uses the supplied cancellation signal', async t => {
  const controller = new AbortController();
  mockRequest(t, 200, (stream, _request, options) => {
    assert.equal(options.signal, controller.signal);
    assert.equal(options.headers['Accept-Encoding'], 'identity');
    assert.equal(options.agent.options.keepAlive, false);
    stream.end('{"output_text":"日本語"}');
  });
  const result = await fetchLongAIResponse('https://example.openai.azure.com/', {
    method: 'POST', headers: {}, body: '日本語', signal: controller.signal,
  });
  assert.equal(result.ok, true);
  assert.equal(result.headers.get('x-request-id'), 'native-test');
  assert.equal(JSON.parse(await result.text()).output_text, '日本語');
});

test('native HTTPS rejects redirects without consuming or forwarding them', async t => {
  mockRequest(t, 302, stream => assert.equal(stream.destroyed, true));
  await assert.rejects(fetchLongAIResponse('https://example.openai.azure.com/', {
    method: 'POST', headers: {}, body: '{}',
  }), /redirects are not allowed/);
});

test('native HTTPS rejects oversized bodies rather than retaining unbounded data', async t => {
  mockRequest(t, 200, stream => stream.end(Buffer.alloc(MAX_RESPONSE_BYTES + 1)));
  const result = await fetchLongAIResponse('https://example.openai.azure.com/', { method: 'POST', headers: {}, body: '{}' });
  await assert.rejects(result.text(), /size limit/);
});

test('native response-stream failures remain rejected after headers', async t => {
  mockRequest(t, 200, stream => stream.destroy(new Error('connection reset')));
  const result = await fetchLongAIResponse('https://example.openai.azure.com/', { method: 'POST', headers: {}, body: '{}' });
  await assert.rejects(result.text(), /connection reset/);
});
