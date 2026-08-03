// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

const assert = require('node:assert/strict');
const { test } = require('node:test');
const express = require('express');
const { createReadinessHandler } = require('./readiness');

test('readiness distinguishes configuration and shutdown state from liveness', async (t) => {
  let configured = false;
  let shuttingDown = false;
  const app = express();
  app.get('/readyz', createReadinessHandler({
    isConfigured: () => configured,
    isShuttingDown: () => shuttingDown,
  }));
  const server = await new Promise((resolve) => {
    const listening = app.listen(0, '127.0.0.1', () => resolve(listening));
  });
  t.after(() => new Promise((resolve, reject) => {
    server.close(error => (error ? reject(error) : resolve()));
  }));
  const { port } = server.address();
  const url = `http://127.0.0.1:${port}/readyz`;

  let response = await fetch(url);
  assert.equal(response.status, 503);
  assert.equal(await response.text(), 'not ready\n');

  configured = true;
  response = await fetch(url);
  assert.equal(response.status, 200);
  assert.equal(await response.text(), 'ready\n');

  shuttingDown = true;
  response = await fetch(url);
  assert.equal(response.status, 503);
});
