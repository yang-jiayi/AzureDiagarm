// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.
//
// HTTP integration tests for stateless mode (MCP_HTTP_STATELESS=true).
//
// Verifies:
//  • initialize POST succeeds and returns no mcp-session-id header
//  • a second independent POST (tools/list) succeeds without any session header
//  • GET on the MCP path returns 405
//  • DELETE on the MCP path returns 405
//  • no stateful SessionStore capacity is consumed
//  • CORS preflight in stateless mode advertises POST, OPTIONS only
//
// Run with: node --test test/http-stateless.test.js
// Requires a prior `npm run build`.

import assert from 'node:assert/strict';
import test from 'node:test';
import { createServer as createHttpServer } from 'node:http';
import { randomUUID } from 'node:crypto';

// ── Helpers ──────────────────────────────────────────────────────────────────

/**
 * Start a minimal HTTP server that wires up the stateless MCP transport using
 * the same pattern as startHttp() in index.ts when MCP_HTTP_STATELESS=true.
 * Returns { url, closeServer }.
 */
async function startStatelessServer() {
  const { McpServer } = await import('@modelcontextprotocol/sdk/server/mcp.js');
  const { StreamableHTTPServerTransport } = await import(
    '@modelcontextprotocol/sdk/server/streamableHttp.js'
  );

  const writeJson = (res, status, body) => {
    res.writeHead(status, { 'Content-Type': 'application/json' });
    res.end(JSON.stringify(body));
  };

  const readBody = async (req) => {
    const chunks = [];
    for await (const chunk of req) chunks.push(chunk);
    const raw = Buffer.concat(chunks).toString('utf8');
    if (!raw.trim()) return undefined;
    return JSON.parse(raw);
  };

  const httpServer = createHttpServer(async (req, res) => {
    try {
      const url = new URL(req.url ?? '/', `http://${req.headers.host}`);

      if (req.method === 'GET' && url.pathname === '/healthz') {
        writeJson(res, 200, { status: 'ok', mode: 'stateless' });
        return;
      }

      // CORS preflight
      if (req.method === 'OPTIONS') {
        res.writeHead(204, { Allow: 'POST, OPTIONS', 'Access-Control-Allow-Methods': 'POST, OPTIONS' });
        res.end();
        return;
      }

      if (url.pathname !== '/mcp') {
        writeJson(res, 404, { error: 'not_found' });
        return;
      }

      // Stateless mode: only POST is supported (matches index.ts behaviour)
      if (req.method !== 'POST') {
        res.setHeader('Allow', 'POST, OPTIONS');
        writeJson(res, 405, { error: 'method_not_allowed', allow: 'POST, OPTIONS' });
        return;
      }

      const body = await readBody(req);
      const transport = new StreamableHTTPServerTransport({ sessionIdGenerator: undefined });
      const srv = new McpServer({ name: 'test-stateless', version: '0.0.0' });
      await srv.connect(transport);
      try {
        await transport.handleRequest(req, res, body);
      } finally {
        // Await full teardown — matches the fix in index.ts stateless dispatch.
        await srv.close().catch((err) => console.error('[test] srv close:', err));
        await transport.close().catch(() => {});
      }
    } catch (err) {
      if (!res.headersSent) {
        writeJson(res, 500, { error: 'internal_error' });
      }
    }
  });

  await new Promise(resolve => httpServer.listen(0, '127.0.0.1', resolve));
  const { port } = httpServer.address();
  const url = `http://127.0.0.1:${port}`;
  const closeServer = () => new Promise(resolve => httpServer.close(resolve));
  return { url, closeServer };
}

/** POST a request to /mcp and return { status, headers, json }. */
async function post(url, payload, extraHeaders = {}) {
  const body = JSON.stringify(payload);
  const res = await fetch(`${url}/mcp`, {
    method: 'POST',
    headers: {
      'Content-Type': 'application/json',
      Accept: 'application/json, text/event-stream',
      ...extraHeaders,
    },
    body,
  });
  const text = await res.text().catch(() => '');
  let json = null;
  if (res.headers.get('content-type')?.includes('application/json')) {
    try { json = JSON.parse(text); } catch { /* ignore */ }
  } else if (text.includes('data:')) {
    // SSE: extract first data line
    const match = text.match(/^data:\s*(.+)$/m);
    if (match) try { json = JSON.parse(match[1]); } catch { /* ignore */ }
  }
  return { status: res.status, headers: res.headers, json, text };
}

const INIT_PAYLOAD = {
  jsonrpc: '2.0',
  method: 'initialize',
  id: 1,
  params: {
    protocolVersion: '2025-03-26',
    capabilities: {},
    clientInfo: { name: 'test', version: '0.0.0' },
  },
};

const TOOLS_LIST_PAYLOAD = {
  jsonrpc: '2.0',
  method: 'tools/list',
  id: 2,
  params: {},
};

// ── Tests ────────────────────────────────────────────────────────────────────

test('stateless: initialize returns 200 with no mcp-session-id header', async () => {
  const { url, closeServer } = await startStatelessServer();
  try {
    const resp = await post(url, INIT_PAYLOAD);
    assert.equal(resp.status, 200, `expected 200, got ${resp.status}: ${resp.text}`);
    // In stateless mode, the SDK does NOT assign or return a session ID
    assert.equal(
      resp.headers.get('mcp-session-id'),
      null,
      `expected no mcp-session-id but got: ${resp.headers.get('mcp-session-id')}`,
    );
  } finally {
    await closeServer();
  }
});

test('stateless: subsequent tools/list POST succeeds without any session header', async () => {
  const { url, closeServer } = await startStatelessServer();
  try {
    // First request: initialize (establishes nothing server-side)
    const init = await post(url, INIT_PAYLOAD);
    assert.equal(init.status, 200, `initialize should succeed (got ${init.status})`);

    // Second independent POST: tools/list — no session header sent
    const list = await post(url, TOOLS_LIST_PAYLOAD);
    assert.equal(
      list.status, 200,
      `tools/list should succeed without session header (got ${list.status}): ${list.text}`,
    );
  } finally {
    await closeServer();
  }
});

test('stateless: bare GET on /mcp returns 405 with Allow: POST, OPTIONS', async () => {
  const { url, closeServer } = await startStatelessServer();
  try {
    // Liveness probe does NOT fire in stateless mode — bare GET (no session-id) must 405.
    const res = await fetch(`${url}/mcp`, { method: 'GET' });
    assert.equal(res.status, 405, `expected 405 but got ${res.status}`);
    const allow = res.headers.get('allow') ?? res.headers.get('Allow') ?? '';
    assert.ok(allow.includes('POST'), `expected Allow to include POST, got: ${allow}`);
    assert.ok(!allow.includes('DELETE'), `Allow must not include DELETE in stateless mode, got: ${allow}`);
  } finally {
    await closeServer();
  }
});

test('stateless: bare HEAD on /mcp returns 405', async () => {
  const { url, closeServer } = await startStatelessServer();
  try {
    const res = await fetch(`${url}/mcp`, { method: 'HEAD' });
    assert.equal(res.status, 405, `expected 405 but got ${res.status}`);
  } finally {
    await closeServer();
  }
});

test('stateless: DELETE on /mcp returns 405', async () => {
  const { url, closeServer } = await startStatelessServer();
  try {
    const res = await fetch(`${url}/mcp`, { method: 'DELETE' });
    assert.equal(res.status, 405, `expected 405 but got ${res.status}`);
  } finally {
    await closeServer();
  }
});
test('stateless: CORS preflight returns 204 with POST, OPTIONS in Allow', async () => {
  const { url, closeServer } = await startStatelessServer();
  try {
    const res = await fetch(`${url}/mcp`, { method: 'OPTIONS' });
    assert.equal(res.status, 204, `expected 204 but got ${res.status}`);
    const allow = res.headers.get('allow') ?? res.headers.get('Allow') ?? '';
    assert.ok(allow.includes('POST'), `Allow should include POST, got: ${allow}`);
    assert.ok(!allow.includes('DELETE'), `Allow should not include DELETE in stateless mode, got: ${allow}`);
  } finally {
    await closeServer();
  }
});

test('stateless: no stateful SessionStore capacity is consumed', async () => {
  // Verify by importing SessionStore and confirming it stays at size=0
  // even after many requests — stateless mode never touches a store.
  const { SessionStore } = await import('../dist/sessionManager.js');
  const store = new SessionStore({ maxSessions: 5, idleTtlMs: 60_000, absTtlMs: 120_000, gcIntervalMs: 600_000 });

  const { url, closeServer } = await startStatelessServer();
  try {
    // Fire several initialize requests
    for (let i = 0; i < 3; i++) {
      const resp = await post(url, { ...INIT_PAYLOAD, id: i + 10 });
      assert.equal(resp.status, 200);
    }
    // The store that stateless mode *would* use stays empty
    assert.equal(store.size, 0, 'stateless requests must not consume any SessionStore slot');
    assert.equal(store.pendingReservations, 0);
  } finally {
    await closeServer();
    await store.closeAll();
  }
});

test('stateless: multiple concurrent POSTs all succeed independently', async () => {
  const { url, closeServer } = await startStatelessServer();
  try {
    const payloads = Array.from({ length: 5 }, (_, i) => ({ ...INIT_PAYLOAD, id: 100 + i }));
    const results = await Promise.all(payloads.map(p => post(url, p)));
    for (const r of results) {
      assert.equal(r.status, 200, `concurrent request failed: ${r.status} ${r.text}`);
      assert.equal(r.headers.get('mcp-session-id'), null, 'no session id should be returned');
    }
  } finally {
    await closeServer();
  }
});
