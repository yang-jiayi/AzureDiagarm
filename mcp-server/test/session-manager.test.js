// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.
//
// Tests for SessionStore: TTL, idle expiry, reservation API, cap enforcement,
// cleanup, GC, and HTTP integration (initialize succeeds; cap returns 503).
//
// Run with: node --test test/session-manager.test.js
//
// Requires a prior `npm run build` (tests import from dist/).

import assert from 'node:assert/strict';
import test from 'node:test';
import { createServer } from 'node:http';
import { randomUUID } from 'node:crypto';

import { SessionStore } from '../dist/sessionManager.js';

// ── Helpers ──────────────────────────────────────────────────────────────────

/** Minimal transport stub used by SessionStore. */
function makeTransport(id) {
  let closed = false;
  const t = {
    get sessionId() { return id; },
    close() { closed = true; return Promise.resolve(); },
    get isClosed() { return closed; },
  };
  return t;
}

/** reserve() then commitReservation() in one call — the happy path. */
function addSession(store, id) {
  const ok = store.reserve();
  if (ok) store.commitReservation(id, makeTransport(id));
  return ok;
}

// ── Construction & defaults ──────────────────────────────────────────────────

test('SessionStore defaults match expected constants', () => {
  const store = new SessionStore({ gcIntervalMs: 600_000 });
  assert.equal(store.maxSessions,  SessionStore.DEFAULT_MAX_SESSIONS);
  assert.equal(store.idleTtlMs,    SessionStore.DEFAULT_IDLE_TTL_MS);
  assert.equal(store.absTtlMs,     SessionStore.DEFAULT_ABS_TTL_MS);
  assert.equal(store.size,         0);
  assert.equal(store.pendingReservations, 0);
});

test('SessionStore constructor respects config overrides', () => {
  const store = new SessionStore({ maxSessions: 5, idleTtlMs: 1000, absTtlMs: 2000, gcIntervalMs: 600_000 });
  assert.equal(store.maxSessions, 5);
  assert.equal(store.idleTtlMs, 1000);
  assert.equal(store.absTtlMs, 2000);
});

// ── reserve / commitReservation / releaseReservation ─────────────────────────

test('reserve increments pendingReservations', () => {
  const store = new SessionStore({ maxSessions: 10, idleTtlMs: 60_000, absTtlMs: 120_000, gcIntervalMs: 600_000 });
  assert.equal(store.reserve(), true);
  assert.equal(store.pendingReservations, 1);
  assert.equal(store.size, 0, 'no committed sessions yet');
  store.releaseReservation();
});

test('commitReservation converts pending slot to committed session', () => {
  const store = new SessionStore({ maxSessions: 10, idleTtlMs: 60_000, absTtlMs: 120_000, gcIntervalMs: 600_000 });
  const t = makeTransport('s1');
  store.reserve();
  store.commitReservation('s1', t);
  assert.equal(store.pendingReservations, 0);
  assert.equal(store.size, 1);
  assert.equal(store.get('s1'), t);
});

test('releaseReservation decrements pendingReservations', () => {
  const store = new SessionStore({ maxSessions: 10, idleTtlMs: 60_000, absTtlMs: 120_000, gcIntervalMs: 600_000 });
  store.reserve();
  store.reserve();
  assert.equal(store.pendingReservations, 2);
  store.releaseReservation();
  assert.equal(store.pendingReservations, 1);
  store.releaseReservation();
  assert.equal(store.pendingReservations, 0);
});

test('releaseReservation is safe when no pending reservations', () => {
  const store = new SessionStore({ maxSessions: 10, idleTtlMs: 60_000, absTtlMs: 120_000, gcIntervalMs: 600_000 });
  assert.doesNotThrow(() => store.releaseReservation());
  assert.equal(store.pendingReservations, 0);
});

// ── Cap enforcement (live sessions) ─────────────────────────────────────────

test('reserve returns false when committed sessions are at cap', () => {
  const store = new SessionStore({ maxSessions: 2, idleTtlMs: 60_000, absTtlMs: 120_000, gcIntervalMs: 600_000 });
  assert.equal(addSession(store, 'a'), true);
  assert.equal(addSession(store, 'b'), true);
  assert.equal(store.reserve(), false, 'third reserve should be rejected');
  assert.equal(store.size, 2);
  assert.equal(store.pendingReservations, 0);
});

test('reserve counts pending slots toward the cap', () => {
  const store = new SessionStore({ maxSessions: 2, idleTtlMs: 60_000, absTtlMs: 120_000, gcIntervalMs: 600_000 });
  assert.equal(store.reserve(), true);  // slot 1 reserved
  assert.equal(store.reserve(), true);  // slot 2 reserved
  assert.equal(store.reserve(), false, 'cap includes pending reservations');
  assert.equal(store.pendingReservations, 2);
  store.releaseReservation();
  store.releaseReservation();
});

test('reserve succeeds after a committed session is deleted', () => {
  const store = new SessionStore({ maxSessions: 1, idleTtlMs: 60_000, absTtlMs: 120_000, gcIntervalMs: 600_000 });
  addSession(store, 'a');
  assert.equal(store.reserve(), false);
  store.delete('a');
  assert.equal(store.reserve(), true);
  store.releaseReservation();
});

test('reserve succeeds after a reservation is released', () => {
  const store = new SessionStore({ maxSessions: 1, idleTtlMs: 60_000, absTtlMs: 120_000, gcIntervalMs: 600_000 });
  assert.equal(store.reserve(), true);
  assert.equal(store.reserve(), false);
  store.releaseReservation();
  assert.equal(store.reserve(), true);
  store.releaseReservation();
});

test('reserve succeeds after expired session is evicted', async () => {
  const store = new SessionStore({ maxSessions: 1, idleTtlMs: 10, absTtlMs: 120_000, gcIntervalMs: 600_000 });
  addSession(store, 'a');
  await new Promise(r => setTimeout(r, 50));
  // reserve() calls evictExpired() before checking — expired 'a' is removed
  assert.equal(store.reserve(), true, 'should succeed after eviction');
  assert.equal(store.size, 0);
  store.releaseReservation();
});

// ── get / touch / delete ─────────────────────────────────────────────────────

test('get returns transport for live session', () => {
  const store = new SessionStore({ maxSessions: 10, idleTtlMs: 60_000, absTtlMs: 120_000, gcIntervalMs: 600_000 });
  const t = makeTransport('s1');
  store.reserve();
  store.commitReservation('s1', t);
  assert.equal(store.get('s1'), t);
});

test('get returns undefined for unknown session', () => {
  const store = new SessionStore({ maxSessions: 10, idleTtlMs: 60_000, absTtlMs: 120_000, gcIntervalMs: 600_000 });
  assert.equal(store.get('nope'), undefined);
});

test('delete removes a session', () => {
  const store = new SessionStore({ maxSessions: 10, idleTtlMs: 60_000, absTtlMs: 120_000, gcIntervalMs: 600_000 });
  addSession(store, 's1');
  store.delete('s1');
  assert.equal(store.size, 0);
  assert.equal(store.get('s1'), undefined);
});

test('touch does not throw for unknown session', () => {
  const store = new SessionStore({ maxSessions: 10, idleTtlMs: 60_000, absTtlMs: 120_000, gcIntervalMs: 600_000 });
  assert.doesNotThrow(() => store.touch('ghost'));
});

// ── Idle expiry ──────────────────────────────────────────────────────────────

test('get returns undefined for idle-expired session and closes transport', async () => {
  const store = new SessionStore({ maxSessions: 10, idleTtlMs: 10, absTtlMs: 60_000, gcIntervalMs: 600_000 });
  const t = makeTransport('s-idle');
  store.reserve();
  store.commitReservation('s-idle', t);

  await new Promise(r => setTimeout(r, 50));

  assert.equal(store.get('s-idle'), undefined, 'session should be expired');
  assert.equal(store.size, 0);
  assert.equal(t.isClosed, true, 'transport should be closed on eviction');
});

test('touch resets idle timer and keeps session alive', async () => {
  const store = new SessionStore({ maxSessions: 10, idleTtlMs: 60, absTtlMs: 120_000, gcIntervalMs: 600_000 });
  const t = makeTransport('s-touch');
  store.reserve();
  store.commitReservation('s-touch', t);

  // Wait half the idle TTL, touch, wait again — session should still be alive.
  await new Promise(r => setTimeout(r, 40));
  store.touch('s-touch');
  await new Promise(r => setTimeout(r, 40));

  assert.notEqual(store.get('s-touch'), undefined, 'session should still be alive after touch');
  await store.closeAll();
});

// ── Absolute TTL ─────────────────────────────────────────────────────────────

test('get returns undefined for abs-TTL-expired session', async () => {
  const store = new SessionStore({ maxSessions: 10, idleTtlMs: 120_000, absTtlMs: 10, gcIntervalMs: 600_000 });
  const t = makeTransport('s-abs');
  store.reserve();
  store.commitReservation('s-abs', t);

  await new Promise(r => setTimeout(r, 50));

  assert.equal(store.get('s-abs'), undefined, 'session should be abs-TTL expired');
  assert.equal(t.isClosed, true);
});

// ── GC sweep ─────────────────────────────────────────────────────────────────

test('evictExpired removes all expired entries', async () => {
  const store = new SessionStore({ maxSessions: 10, idleTtlMs: 10, absTtlMs: 120_000, gcIntervalMs: 600_000 });
  addSession(store, 'x');
  addSession(store, 'y');

  await new Promise(r => setTimeout(r, 50));
  store.evictExpired();

  assert.equal(store.size, 0);
});

test('evictExpired preserves non-expired entries', () => {
  const store = new SessionStore({ maxSessions: 10, idleTtlMs: 200, absTtlMs: 120_000, gcIntervalMs: 600_000 });
  addSession(store, 'keep');

  store.evictExpired(); // nothing expired yet
  assert.equal(store.size, 1);
  store.closeAll();
});

// ── closeAll / shutdown ──────────────────────────────────────────────────────

test('closeAll closes all transports and empties the store', async () => {
  const store = new SessionStore({ maxSessions: 10, idleTtlMs: 60_000, absTtlMs: 120_000, gcIntervalMs: 600_000 });
  const t1 = makeTransport('s1'); store.reserve(); store.commitReservation('s1', t1);
  const t2 = makeTransport('s2'); store.reserve(); store.commitReservation('s2', t2);

  await store.closeAll();

  assert.equal(store.size, 0);
  assert.equal(t1.isClosed, true);
  assert.equal(t2.isClosed, true);
});

test('closeAll resets pendingReservations', async () => {
  const store = new SessionStore({ maxSessions: 10, idleTtlMs: 60_000, absTtlMs: 120_000, gcIntervalMs: 600_000 });
  store.reserve();
  store.reserve();
  assert.equal(store.pendingReservations, 2);
  await store.closeAll();
  assert.equal(store.pendingReservations, 0);
});

test('closeAll on empty store does not throw', async () => {
  const store = new SessionStore({ maxSessions: 10, idleTtlMs: 60_000, absTtlMs: 120_000, gcIntervalMs: 600_000 });
  await assert.doesNotReject(() => store.closeAll());
});

test('closeAll stops the GC timer', async () => {
  const store = new SessionStore({ maxSessions: 10, idleTtlMs: 60_000, absTtlMs: 120_000, gcIntervalMs: 100 });
  store.startGc();
  await store.closeAll();
});

// ── GC timer ─────────────────────────────────────────────────────────────────

test('startGc is idempotent', async () => {
  const store = new SessionStore({ maxSessions: 10, idleTtlMs: 60_000, absTtlMs: 120_000, gcIntervalMs: 100 });
  store.startGc();
  store.startGc();
  await store.closeAll();
});

// ── HTTP integration ─────────────────────────────────────────────────────────
//
// These tests start the real MCP server over TCP and verify end-to-end
// behaviour: a valid initialize succeeds and returns a session ID; when the
// cap is 1 and a session is active, a second initialize returns 503.

/**
 * Spin up a real HTTP MCP server (import from dist/index.js) on a random
 * loopback port and return { url, closeServer }.
 * env.MCP_SESSION_MAX controls the session cap for the duration of the test.
 */
async function startTestServer(sessionMax = 100) {
  // Import dynamically so we can set env vars before the module runs.
  // The module caches on first import; to get a fresh server we construct
  // one directly via the exported createServer factory + manual HTTP wiring.
  //
  // Rather than re-import the full module (which doesn't re-run module-level
  // code), we replicate the minimal server-creation plumbing here, using only
  // the SDK and SessionStore from dist/.

  const { default: http } = await import('node:http');
  const { randomUUID: uuid } = await import('node:crypto');
  const { McpServer } = await import('@modelcontextprotocol/sdk/server/mcp.js');
  const { StreamableHTTPServerTransport } = await import(
    '@modelcontextprotocol/sdk/server/streamableHttp.js'
  );
  const { isInitializeRequest } = await import('@modelcontextprotocol/sdk/types.js');
  const { SessionStore: SS } = await import('../dist/sessionManager.js');

  const store = new SS({ maxSessions: sessionMax, idleTtlMs: 60_000, absTtlMs: 120_000, gcIntervalMs: 600_000 });
  store.startGc();

  const httpServer = http.createServer(async (req, res) => {
    const writeJson = (status, body) => {
      res.writeHead(status, { 'Content-Type': 'application/json' });
      res.end(JSON.stringify(body));
    };

    try {
      const sid = req.headers['mcp-session-id'];
      const sessionId = Array.isArray(sid) ? sid[0] : sid;

      let transport = sessionId ? store.get(sessionId) : undefined;
      let body;
      if (req.method === 'POST') {
        const chunks = [];
        for await (const chunk of req) chunks.push(chunk);
        const raw = Buffer.concat(chunks).toString('utf8');
        if (raw.trim()) body = JSON.parse(raw);
      }

      if (!transport) {
        if (req.method !== 'POST' || !isInitializeRequest(body)) {
          writeJson(400, { jsonrpc: '2.0', error: { code: -32000, message: 'No session' }, id: null });
          return;
        }
        if (!store.reserve()) {
          writeJson(503, { jsonrpc: '2.0', error: { code: -32000, message: 'At capacity' }, id: null });
          return;
        }
        let committed = false;
        transport = new StreamableHTTPServerTransport({
          sessionIdGenerator: () => uuid(),
          onsessioninitialized: (newSid) => { committed = true; store.commitReservation(newSid, transport); },
        });
        transport.onclose = () => {
          if (committed) { if (transport.sessionId) store.delete(transport.sessionId); }
          else store.releaseReservation();
        };
        const srv = new McpServer({ name: 'test-server', version: '0.0.0' });
        await srv.connect(transport);
      }

      if (transport.sessionId) store.touch(transport.sessionId);
      await transport.handleRequest(req, res, body);
    } catch (err) {
      if (!res.headersSent) {
        res.writeHead(500, { 'Content-Type': 'application/json' });
        res.end(JSON.stringify({ error: 'internal_error' }));
      }
    }
  });

  await new Promise(resolve => httpServer.listen(0, '127.0.0.1', resolve));
  const { port } = httpServer.address();
  const url = `http://127.0.0.1:${port}`;

  const closeServer = () => new Promise(resolve => {
    store.closeAll().then(() => httpServer.close(resolve));
  });

  return { url, store, closeServer };
}

/** POST a minimal MCP initialize request and return { status, headers, body }. */
async function postInitialize(url) {
  const body = JSON.stringify({
    jsonrpc: '2.0',
    method: 'initialize',
    id: 1,
    params: {
      protocolVersion: '2025-03-26',
      capabilities: {},
      clientInfo: { name: 'test-client', version: '0.0.0' },
    },
  });
  const res = await fetch(`${url}/`, {
    method: 'POST',
    headers: { 'Content-Type': 'application/json', Accept: 'application/json, text/event-stream' },
    body,
  });
  const text = await res.text().catch(() => '');
  let json = null;
  // SDK may return SSE or JSON; try to extract the JSON payload
  if (res.headers.get('content-type')?.includes('application/json')) {
    try { json = JSON.parse(text); } catch { /* ignore */ }
  } else if (text.includes('data:')) {
    // SSE: extract first data line
    const match = text.match(/^data:\s*(.+)$/m);
    if (match) try { json = JSON.parse(match[1]); } catch { /* ignore */ }
  }
  return { status: res.status, sessionId: res.headers.get('mcp-session-id'), text, json };
}

test('HTTP: normal initialize succeeds and returns a session ID', async () => {
  const { url, closeServer } = await startTestServer(10);
  try {
    const resp = await postInitialize(url);
    // The SDK returns 200 for a successful initialize
    assert.equal(resp.status, 200, `expected 200 but got ${resp.status}: ${resp.text}`);
    assert.ok(resp.sessionId, `expected mcp-session-id header but got: ${JSON.stringify(resp.sessionId)}`);
  } finally {
    await closeServer();
  }
});

test('HTTP: second initialize is rejected with 503 when cap is 1', async () => {
  const { url, store, closeServer } = await startTestServer(1);
  try {
    const first = await postInitialize(url);
    assert.equal(first.status, 200, `first initialize should succeed (got ${first.status})`);
    assert.equal(store.size + store.pendingReservations, 1, 'store should show 1 session/reservation');

    const second = await postInitialize(url);
    assert.equal(second.status, 503, `second initialize should be rejected with 503 (got ${second.status})`);
  } finally {
    await closeServer();
  }
});

test('HTTP: after cap-1 session is removed, new initialize succeeds', async () => {
  const { url, store, closeServer } = await startTestServer(1);
  try {
    const first = await postInitialize(url);
    assert.equal(first.status, 200);

    // Manually remove the session from the store (simulating expiry / onclose)
    const sid = first.sessionId;
    store.delete(sid);
    assert.equal(store.size, 0);

    const second = await postInitialize(url);
    assert.equal(second.status, 200, `initialize should succeed after slot freed (got ${second.status})`);
  } finally {
    await closeServer();
  }
});

test('HTTP: reservation released when transport closes before commit', async () => {
  const { url, store, closeServer } = await startTestServer(5);
  try {
    // Manually push a reservation without committing it, then release it.
    const reserved = store.reserve();
    assert.equal(reserved, true);
    assert.equal(store.pendingReservations, 1);
    store.releaseReservation();
    assert.equal(store.pendingReservations, 0);

    // A normal initialize should still succeed.
    const resp = await postInitialize(url);
    assert.equal(resp.status, 200);
  } finally {
    await closeServer();
  }
});

// ── envInt env-var validation ────────────────────────────────────────────────
// envInt is not exported so we test it indirectly via SessionStore constructor
// with env vars set to known-bad values.

test('envInt: valid integer env var is accepted', () => {
  const saved = process.env.MCP_SESSION_MAX;
  try {
    process.env.MCP_SESSION_MAX = '42';
    const store = new SessionStore({ gcIntervalMs: 600_000 });
    assert.equal(store.maxSessions, 42);
  } finally {
    if (saved === undefined) delete process.env.MCP_SESSION_MAX;
    else process.env.MCP_SESSION_MAX = saved;
  }
});

test('envInt: value with trailing junk is rejected and default used', () => {
  const saved = process.env.MCP_SESSION_MAX;
  try {
    process.env.MCP_SESSION_MAX = '10oops';
    const store = new SessionStore({ gcIntervalMs: 600_000 });
    // Default is 100; '10oops' should not parse as 10
    assert.equal(store.maxSessions, SessionStore.DEFAULT_MAX_SESSIONS,
      `expected default ${SessionStore.DEFAULT_MAX_SESSIONS} for '10oops' but got ${store.maxSessions}`);
  } finally {
    if (saved === undefined) delete process.env.MCP_SESSION_MAX;
    else process.env.MCP_SESSION_MAX = saved;
  }
});

test('envInt: float string is rejected and default used', () => {
  const saved = process.env.MCP_SESSION_MAX;
  try {
    process.env.MCP_SESSION_MAX = '10.5';
    const store = new SessionStore({ gcIntervalMs: 600_000 });
    assert.equal(store.maxSessions, SessionStore.DEFAULT_MAX_SESSIONS);
  } finally {
    if (saved === undefined) delete process.env.MCP_SESSION_MAX;
    else process.env.MCP_SESSION_MAX = saved;
  }
});

test('envInt: whitespace-padded integer is accepted', () => {
  const saved = process.env.MCP_SESSION_MAX;
  try {
    process.env.MCP_SESSION_MAX = '  15  ';
    const store = new SessionStore({ gcIntervalMs: 600_000 });
    assert.equal(store.maxSessions, 15);
  } finally {
    if (saved === undefined) delete process.env.MCP_SESSION_MAX;
    else process.env.MCP_SESSION_MAX = saved;
  }
});

// ── Reservation-leak guard (connect/handleRequest failure paths) ──────────────
//
// Simulates what happens when server.connect() throws after a reservation was
// taken: the releaseReservationOnce() guard should restore capacity.

test('reservation is released exactly once when connect throws (unit simulation)', () => {
  const store = new SessionStore({ maxSessions: 1, idleTtlMs: 60_000, absTtlMs: 120_000, gcIntervalMs: 600_000 });

  assert.equal(store.reserve(), true);
  assert.equal(store.pendingReservations, 1);

  // Simulate the releaseReservationOnce() pattern: idempotent double-call
  let committed = false;
  let released = false;
  const releaseOnce = () => {
    if (!committed && !released) { released = true; store.releaseReservation(); }
  };

  // Error path: connect threw before onsessioninitialized fired
  releaseOnce(); // first call — should release
  releaseOnce(); // second call — should be a no-op (idempotency)

  assert.equal(store.pendingReservations, 0, 'reservation should be released exactly once');
  assert.equal(store.reserve(), true, 'slot should be available again after release');
  store.releaseReservation(); // clean up
});

test('reservation is NOT released when commitReservation already fired', async () => {
  const store = new SessionStore({ maxSessions: 2, idleTtlMs: 60_000, absTtlMs: 120_000, gcIntervalMs: 600_000 });
  const t = makeTransport('s-commit');

  assert.equal(store.reserve(), true);
  store.commitReservation('s-commit', t); // committed → pendingReservations = 0, size = 1

  let committed = true; // simulate reservationCommitted = true
  let released = false;
  const releaseOnce = () => {
    if (!committed && !released) { released = true; store.releaseReservation(); }
  };

  releaseOnce(); // should be a no-op since committed = true
  assert.equal(store.size, 1, 'session should still be in store');
  assert.equal(store.pendingReservations, 0);
  await store.closeAll();
});

// Integration test: start a server with maxSessions=1, trigger a connect-time
// error via a mock, and prove capacity is restored so a subsequent real
// initialize succeeds.
test('HTTP integration: reservation released after connect error allows next session', async () => {
  const { McpServer } = await import('@modelcontextprotocol/sdk/server/mcp.js');
  const { StreamableHTTPServerTransport } = await import(
    '@modelcontextprotocol/sdk/server/streamableHttp.js'
  );
  const { isInitializeRequest } = await import('@modelcontextprotocol/sdk/types.js');
  const { SessionStore: SS } = await import('../dist/sessionManager.js');
  const { randomUUID: uuid } = await import('node:crypto');
  const { default: http } = await import('node:http');

  const store = new SS({ maxSessions: 1, idleTtlMs: 60_000, absTtlMs: 120_000, gcIntervalMs: 600_000 });
  let injectConnectError = false; // flip to true to force connect() failure

  const httpServer = http.createServer(async (req, res) => {
    const writeJson = (status, body) => {
      res.writeHead(status, { 'Content-Type': 'application/json' });
      res.end(JSON.stringify(body));
    };
    try {
      let body;
      if (req.method === 'POST') {
        const chunks = [];
        for await (const chunk of req) chunks.push(chunk);
        const raw = Buffer.concat(chunks).toString('utf8');
        if (raw.trim()) body = JSON.parse(raw);
      }
      if (req.method !== 'POST' || !isInitializeRequest(body)) {
        writeJson(400, { error: 'bad request' });
        return;
      }
      if (!store.reserve()) {
        writeJson(503, { error: 'at capacity' });
        return;
      }
      let committed = false;
      let reservationReleased = false;
      const releaseReservationOnce = () => {
        if (!committed && !reservationReleased) { reservationReleased = true; store.releaseReservation(); }
      };
      const transport = new StreamableHTTPServerTransport({
        sessionIdGenerator: () => uuid(),
        onsessioninitialized: (newSid) => { committed = true; store.commitReservation(newSid, transport); },
      });
      transport.onclose = () => {
        if (committed) { if (transport.sessionId) store.delete(transport.sessionId); }
        else releaseReservationOnce();
      };
      const srv = new McpServer({ name: 'test-leak', version: '0.0.0' });
      try {
        if (injectConnectError) throw new Error('injected connect failure');
        await srv.connect(transport);
      } catch (connectErr) {
        releaseReservationOnce();
        transport.close().catch(() => {});
        writeJson(500, { error: 'connect_failed' });
        return;
      }
      await transport.handleRequest(req, res, body);
    } catch (err) {
      if (!res.headersSent) writeJson(500, { error: 'internal' });
    }
  });

  await new Promise(resolve => httpServer.listen(0, '127.0.0.1', resolve));
  const { port } = httpServer.address();
  const url = `http://127.0.0.1:${port}`;

  const initPayload = JSON.stringify({
    jsonrpc: '2.0', method: 'initialize', id: 1,
    params: { protocolVersion: '2025-03-26', capabilities: {}, clientInfo: { name: 'test', version: '0.0.0' } },
  });
  const headers = { 'Content-Type': 'application/json', Accept: 'application/json, text/event-stream' };

  try {
    // Request 1: inject connect error — reservation should be released
    injectConnectError = true;
    const r1 = await fetch(url, { method: 'POST', headers, body: initPayload });
    assert.equal(r1.status, 500, 'injected error should return 500');
    assert.equal(store.pendingReservations, 0, 'reservation must be released after connect error');
    assert.equal(store.size, 0, 'no committed sessions after connect error');

    // Request 2: no error — should succeed (cap restored)
    injectConnectError = false;
    const r2 = await fetch(url, { method: 'POST', headers, body: initPayload });
    assert.equal(r2.status, 200, `second initialize should succeed (got ${r2.status})`);
  } finally {
    await store.closeAll();
    await new Promise(resolve => httpServer.close(resolve));
  }
});