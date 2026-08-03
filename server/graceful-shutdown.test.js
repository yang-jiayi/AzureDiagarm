// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

const assert = require('node:assert/strict');
const http = require('node:http');
const { test } = require('node:test');
const { createGracefulShutdown } = require('./graceful-shutdown');

test('graceful shutdown drains an active request before completing', async () => {
  let requestStarted;
  const started = new Promise((resolve) => {
    requestStarted = resolve;
  });
  let releaseRequest;
  const release = new Promise((resolve) => {
    releaseRequest = resolve;
  });

  const server = http.createServer(async (_req, res) => {
    requestStarted();
    await release;
    res.end('done');
  });
  await new Promise((resolve) => server.listen(0, '127.0.0.1', resolve));
  const { port } = server.address();

  let exitCode = null;
  let resolveExit;
  const exited = new Promise((resolve) => {
    resolveExit = resolve;
  });
  const shutdown = createGracefulShutdown(server, {
    timeoutMs: 5_000,
    logger: { info() {}, error() {} },
    exit(code) {
      exitCode = code;
      resolveExit();
    },
  });

  const responsePromise = fetch(`http://127.0.0.1:${port}`);
  await started;
  assert.equal(shutdown('SIGTERM'), true);
  assert.equal(shutdown('SIGTERM'), false);
  await new Promise(resolve => setTimeout(resolve, 20));
  assert.equal(exitCode, null);

  releaseRequest();
  const response = await responsePromise;
  assert.equal(await response.text(), 'done');
  await exited;
  assert.equal(exitCode, 0);
});
