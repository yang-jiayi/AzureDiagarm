// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

const assert = require('node:assert/strict');
const { test } = require('node:test');
const express = require('express');
const { asyncHandler, createErrorHandler } = require('./async-handler');

function silentLogger() {
  const entries = [];
  return {
    entries,
    error: (...args) => entries.push(['error', ...args]),
    warn: (...args) => entries.push(['warn', ...args]),
  };
}

async function startServer(configure) {
  const app = express();
  configure(app);
  const server = await new Promise((resolve) => {
    const listening = app.listen(0, '127.0.0.1', () => resolve(listening));
  });
  const { port } = server.address();
  return {
    baseUrl: `http://127.0.0.1:${port}`,
    close: () => new Promise((resolve, reject) => {
      server.close((error) => (error ? reject(error) : resolve()));
    }),
  };
}

test('asyncHandler forwards rejected promises instead of crashing the process', async () => {
  const logger = silentLogger();
  const server = await startServer((app) => {
    app.get('/boom', asyncHandler(async () => {
      throw new Error('async failure');
    }));
    app.use(createErrorHandler(logger));
  });

  try {
    const response = await fetch(`${server.baseUrl}/boom`);
    assert.equal(response.status, 500);
    assert.deepEqual(await response.json(), { error: 'Internal server error' });
    assert.equal(logger.entries[0][0], 'error');
  } finally {
    await server.close();
  }
});

test('asyncHandler forwards synchronous throws', async () => {
  const server = await startServer((app) => {
    app.get('/boom', asyncHandler(() => {
      throw new Error('sync failure');
    }));
    app.use(createErrorHandler(silentLogger()));
  });

  try {
    const response = await fetch(`${server.baseUrl}/boom`);
    assert.equal(response.status, 500);
  } finally {
    await server.close();
  }
});

test('asyncHandler works for middleware that calls next', async () => {
  const server = await startServer((app) => {
    app.get(
      '/ok',
      asyncHandler(async (_req, _res, next) => { next(); }),
      (_req, res) => res.json({ ok: true }),
    );
    app.use(createErrorHandler(silentLogger()));
  });

  try {
    const response = await fetch(`${server.baseUrl}/ok`);
    assert.equal(response.status, 200);
    assert.deepEqual(await response.json(), { ok: true });
  } finally {
    await server.close();
  }
});

test('the error handler preserves body-parser client error statuses', async () => {
  const logger = silentLogger();
  const server = await startServer((app) => {
    app.use(express.json({ limit: '1kb' }));
    app.post('/echo', (req, res) => res.json(req.body));
    app.use(createErrorHandler(logger));
  });

  try {
    const malformed = await fetch(`${server.baseUrl}/echo`, {
      method: 'POST',
      headers: { 'content-type': 'application/json' },
      body: '{"broken":',
    });
    assert.equal(malformed.status, 400);
    assert.deepEqual(await malformed.json(), { error: 'Invalid request' });

    const oversized = await fetch(`${server.baseUrl}/echo`, {
      method: 'POST',
      headers: { 'content-type': 'application/json' },
      body: JSON.stringify({ padding: 'x'.repeat(4096) }),
    });
    assert.equal(oversized.status, 413);

    assert.ok(logger.entries.every((entry) => entry[0] === 'warn'));
  } finally {
    await server.close();
  }
});

test('the error handler ignores out-of-range status hints', async () => {
  const server = await startServer((app) => {
    app.get('/boom', asyncHandler(async () => {
      const error = new Error('bogus status');
      error.status = 99;
      throw error;
    }));
    app.use(createErrorHandler(silentLogger()));
  });

  try {
    const response = await fetch(`${server.baseUrl}/boom`);
    assert.equal(response.status, 500);
  } finally {
    await server.close();
  }
});
