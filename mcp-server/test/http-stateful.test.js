// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

import assert from 'node:assert/strict';
import { once } from 'node:events';
import { createServer } from 'node:net';
import { spawn } from 'node:child_process';
import test from 'node:test';
import { fileURLToPath } from 'node:url';

const entryPoint = fileURLToPath(new URL('../dist/index.js', import.meta.url));

async function allocatePort() {
  const server = createServer();
  await new Promise((resolve, reject) => {
    server.once('error', reject);
    server.listen(0, '127.0.0.1', resolve);
  });
  const address = server.address();
  const port = typeof address === 'object' && address ? address.port : 0;
  await new Promise((resolve) => server.close(resolve));
  return port;
}

async function startStatefulServer() {
  const port = await allocatePort();
  const child = spawn(process.execPath, [entryPoint, '--http'], {
    env: {
      ...process.env,
      MCP_AUTH_TOKEN: '',
      MCP_HTTP_HOST: '127.0.0.1',
      MCP_HTTP_PORT: String(port),
      MCP_HTTP_STATELESS: 'false',
      MCP_SESSION_MAX: '1',
      MCP_SESSION_IDLE_SECONDS: '60',
      MCP_SESSION_TTL_SECONDS: '120',
      MCP_SESSION_GC_SECONDS: '60',
    },
    stdio: ['ignore', 'ignore', 'pipe'],
  });

  let stderr = '';
  const ready = new Promise((resolve, reject) => {
    const timeout = setTimeout(
      () => reject(new Error(`Timed out waiting for MCP server.\n${stderr}`)),
      10_000,
    );
    child.stderr.setEncoding('utf8');
    child.stderr.on('data', (chunk) => {
      stderr += chunk;
      if (stderr.includes('azure-diagram-builder listening')) {
        clearTimeout(timeout);
        resolve();
      }
    });
    child.once('exit', (code, signal) => {
      clearTimeout(timeout);
      reject(new Error(`MCP server exited before startup (${code ?? signal}).\n${stderr}`));
    });
  });

  await ready;
  return {
    url: `http://127.0.0.1:${port}`,
    async close() {
      if (child.exitCode !== null) return;
      child.kill('SIGTERM');
      await Promise.race([
        once(child, 'exit'),
        new Promise((_, reject) => setTimeout(
          () => reject(new Error(`Timed out stopping MCP server.\n${stderr}`)),
          5_000,
        )),
      ]);
    },
  };
}

function initializeBody() {
  return JSON.stringify({
    jsonrpc: '2.0',
    method: 'initialize',
    id: 1,
    params: {
      protocolVersion: '2025-03-26',
      capabilities: {},
      clientInfo: { name: 'stateful-regression-test', version: '1.0.0' },
    },
  });
}

test('stateful: rejected initialize releases its pending reservation', async () => {
  const server = await startStatefulServer();
  try {
    const rejected = await fetch(`${server.url}/mcp`, {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json',
        Accept: 'application/json',
      },
      body: initializeBody(),
    });
    assert.equal(rejected.status, 406);

    const health = await fetch(`${server.url}/healthz`).then((response) => response.json());
    assert.equal(health.sessions, 0);
    assert.equal(health.pendingReservations, 0);

    const valid = await fetch(`${server.url}/mcp`, {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json',
        Accept: 'application/json, text/event-stream',
      },
      body: initializeBody(),
    });
    assert.equal(valid.status, 200);
    assert.ok(valid.headers.get('mcp-session-id'));
  } finally {
    await server.close();
  }
});
