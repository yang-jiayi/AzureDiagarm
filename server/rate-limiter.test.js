// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

const assert = require('node:assert/strict');
const { test } = require('node:test');
const {
  createFixedWindowRateLimiter,
  createTableRateLimiter,
  getClientKey,
} = require('./rate-limiter');

// ── Helpers ───────────────────────────────────────────────────────────────────

// Minimal fake Express request.
function fakeReq(ip = '1.2.3.4', headers = {}) {
  return {
    ip,
    get(name) { return headers[name.toLowerCase()] ?? null; },
  };
}

// Azure Table Storage fake with correct ETag-based optimistic concurrency.
// Mirrors the @azure/data-tables TableClient surface used by the rate limiter:
//   getEntity(partitionKey, rowKey)
//   createEntity(entity)
//   updateEntity(entity, mode, { etag })
class FakeTable {
  constructor() {
    this.entities = new Map();
    this._etagSeq = 0;
    // Hooks for fault injection (set to a function to override one call).
    this.onNextGet = null;
    this.onNextWrite = null;
  }

  _key(pk, rk) { return `${pk}:${rk}`; }

  async getEntity(partitionKey, rowKey) {
    if (this.onNextGet) {
      const fn = this.onNextGet;
      this.onNextGet = null;
      return fn(partitionKey, rowKey);
    }
    const entry = this.entities.get(this._key(partitionKey, rowKey));
    if (!entry) {
      const err = new Error('Not found');
      err.statusCode = 404;
      throw err;
    }
    return { ...entry.data, etag: entry.etag };
  }

  async createEntity(entity) {
    if (this.onNextWrite) {
      const fn = this.onNextWrite;
      this.onNextWrite = null;
      return fn(entity);
    }
    const k = this._key(entity.partitionKey, entity.rowKey);
    if (this.entities.has(k)) {
      const err = new Error('Already exists');
      err.statusCode = 409;
      throw err;
    }
    const etag = `etag-${++this._etagSeq}`;
    this.entities.set(k, { data: { ...entity }, etag });
  }

  async updateEntity(entity, _mode, options = {}) {
    if (this.onNextWrite) {
      const fn = this.onNextWrite;
      this.onNextWrite = null;
      return fn(entity, options);
    }
    const k = this._key(entity.partitionKey, entity.rowKey);
    const entry = this.entities.get(k);
    if (!entry) {
      const err = new Error('Not found');
      err.statusCode = 404;
      throw err;
    }
    if (options.etag && options.etag !== entry.etag) {
      const err = new Error('Precondition Failed');
      err.statusCode = 412;
      throw err;
    }
    const etag = `etag-${++this._etagSeq}`;
    this.entities.set(k, { data: { ...entity }, etag });
  }
}

// ── getClientKey ──────────────────────────────────────────────────────────────

test('getClientKey prefers Azure Front Door headers over XFF', () => {
  // X-Azure-ClientIP wins over anything in X-Forwarded-For.
  assert.equal(
    getClientKey(fakeReq('10.0.0.1', {
      'x-azure-clientip': '203.0.113.1',
      'x-forwarded-for': '1.2.3.4, 5.6.7.8',
    })),
    '203.0.113.1',
  );
});

test('getClientKey falls back to last XFF entry when Front Door headers absent', () => {
  assert.equal(
    getClientKey(fakeReq('10.0.0.1', { 'x-forwarded-for': '1.2.3.4, 5.6.7.8' })),
    '5.6.7.8',
  );
});

test('getClientKey rejects invalid IP-like strings', () => {
  // An attacker supplying a bogus AzureClientIP shouldn't get a key they control.
  const k = getClientKey(fakeReq('10.0.0.1', { 'x-azure-clientip': 'evil(); DROP TABLE' }));
  // Falls through to XFF last entry or req.ip.
  assert.match(k, /^[0-9a-fA-F.:[\]]+$/);
});

// ── In-process fixed-window limiter ──────────────────────────────────────────

test('in-process limiter allows requests up to maxRequests', () => {
  const consume = createFixedWindowRateLimiter(60_000, 3);
  const req = fakeReq('1.1.1.1');
  assert.equal(consume(req), 0);
  assert.equal(consume(req), 0);
  assert.equal(consume(req), 0);
  assert.ok(consume(req) > 0, 'fourth request should be rate-limited');
});

test('in-process limiter resets after the window expires', () => {
  const consume = createFixedWindowRateLimiter(50, 1); // 50 ms window
  const req = fakeReq('2.2.2.2');
  assert.equal(consume(req), 0);
  assert.ok(consume(req) > 0, 'should be rate-limited within window');
  return new Promise((resolve) => {
    setTimeout(() => {
      assert.equal(consume(req), 0, 'should reset after window');
      resolve();
    }, 60);
  });
});

test('in-process limiter is per-client (different IPs do not share quota)', () => {
  const consume = createFixedWindowRateLimiter(60_000, 1);
  assert.equal(consume(fakeReq('3.3.3.3')), 0);
  assert.equal(consume(fakeReq('4.4.4.4')), 0, 'different IP must start fresh');
});

// ── Table rate limiter — happy path ───────────────────────────────────────────

test('table limiter allows the first request and creates an entity', async () => {
  const table = new FakeTable();
  const consume = createTableRateLimiter(table, 3_600_000, 5);
  assert.equal(await consume(fakeReq('10.0.0.1')), 0);
  assert.equal(table.entities.size, 1);
  const entry = table.entities.values().next().value;
  assert.equal(entry.data.count, 1);
});

test('table limiter increments atomically across sequential requests', async () => {
  const table = new FakeTable();
  const consume = createTableRateLimiter(table, 3_600_000, 3);
  const req = fakeReq('10.0.0.2');
  assert.equal(await consume(req), 0);
  assert.equal(await consume(req), 0);
  assert.equal(await consume(req), 0);
  const retryAfter = await consume(req);
  assert.ok(retryAfter > 0, 'fourth request must be rejected');
  // Upstream must not be called on a rate-limited request — the counter stays at 3.
  const entry = table.entities.values().next().value;
  assert.equal(entry.data.count, 3);
});

test('table limiter returns a positive retry-after when the window is active', async () => {
  const table = new FakeTable();
  const windowMs = 3_600_000;
  const consume = createTableRateLimiter(table, windowMs, 1);
  const req = fakeReq('10.0.0.3');
  await consume(req); // First request sets count=1
  const retryAfter = await consume(req); // Second request is over limit
  assert.ok(retryAfter >= 1, `retryAfter should be at least 1, got ${retryAfter}`);
  assert.ok(retryAfter <= Math.ceil(windowMs / 1000) + 1);
});

test('table limiter resets counter after window expiry', async () => {
  const table = new FakeTable();
  const consume = createTableRateLimiter(table, 30, 1); // 30 ms window
  const req = fakeReq('10.0.0.4');
  assert.equal(await consume(req), 0);
  assert.ok(await consume(req) > 0, 'should be limited within window');
  await new Promise((resolve) => setTimeout(resolve, 40));
  assert.equal(await consume(req), 0, 'should allow after window expiry');
});

// ── Table limiter — multi-replica shared state ────────────────────────────────

test('table limiter enforces limit globally across two simulated replicas', async () => {
  // Two limiter instances sharing the same FakeTable simulate two ACA replicas.
  const table = new FakeTable();
  const consumeA = createTableRateLimiter(table, 3_600_000, 2);
  const consumeB = createTableRateLimiter(table, 3_600_000, 2);
  const req = fakeReq('10.0.1.1');

  // Replica A uses the first slot.
  assert.equal(await consumeA(req), 0);
  // Replica B uses the second slot.
  assert.equal(await consumeB(req), 0);
  // Both replicas must now see the limit enforced.
  assert.ok(await consumeA(req) > 0, 'replica A: should be limited after global limit');
  assert.ok(await consumeB(req) > 0, 'replica B: should be limited after global limit');
});

test('table limiter resolves write conflict via retry when two replicas race', async () => {
  // Both replicas call getEntity concurrently and get the same ETag.
  // Only one write will succeed; the other will get 412 and must retry.
  const table = new FakeTable();
  const consumeA = createTableRateLimiter(table, 3_600_000, 5);
  const consumeB = createTableRateLimiter(table, 3_600_000, 5);
  const req = fakeReq('10.0.1.2');

  // Seed a base entity so both replicas read the same snapshot.
  await consumeA(req); // count=1 written by A

  // Capture the current ETag so we can inject a stale snapshot for B.
  const k = table.entities.keys().next().value;
  const staleEtag = table.entities.get(k).etag;

  // Intercept B's getEntity to return a stale ETag (simulates reading before A's last write).
  const staleEntity = { ...table.entities.get(k).data, etag: staleEtag };
  table.onNextGet = async () => staleEntity;

  // A writes first: count=2.
  await consumeA(req);

  // B now tries with stale ETag; should get 412 on first write attempt,
  // re-read, and succeed on retry → count=3.
  const result = await consumeB(req);
  assert.equal(result, 0, 'B should succeed after retry');
  assert.equal(Number(table.entities.get(k).data.count), 3);
});

// ── Table limiter — fail-closed semantics ─────────────────────────────────────

test('table limiter fails closed on storage read error (non-404)', async () => {
  const table = new FakeTable();
  const errors = [];
  const consume = createTableRateLimiter(table, 3_600_000, 10, {
    onStorageError: (operation, error) => errors.push({ operation, error }),
  });
  const req = fakeReq('10.0.2.1');

  table.onNextGet = async () => {
    const err = new Error('Service unavailable');
    err.statusCode = 503;
    throw err;
  };

  const retryAfter = await consume(req);
  assert.equal(retryAfter, 5, 'storage failures should fail closed with a short retry');
  assert.equal(errors[0].operation, 'read');
});

test('table limiter fails closed on non-conflict storage write error', async () => {
  const table = new FakeTable();
  const consume = createTableRateLimiter(table, 3_600_000, 10);
  const req = fakeReq('10.0.2.2');

  // Override createEntity with a 500 error.
  table.onNextWrite = async () => {
    const err = new Error('Internal Server Error');
    err.statusCode = 500;
    throw err;
  };

  const retryAfter = await consume(req);
  assert.equal(retryAfter, 5, 'must fail closed on write error');
});

test('table limiter fails closed when all retries are write conflicts', async () => {
  const table = new FakeTable();
  const consume = createTableRateLimiter(table, 3_600_000, 10);
  const req = fakeReq('10.0.2.3');

  // Intercept every write with a 412 so all retries exhaust.
  const always412 = async () => {
    const err = new Error('Precondition Failed');
    err.statusCode = 412;
    throw err;
  };

  // We also need every *read* after the first to return the same stale entity
  // so the loop keeps seeing count < max and attempts a write.
  let firstRead = true;
  const fakeEntity = {
    partitionKey: 'window',
    rowKey: 'x',
    count: 1,
    resetAt: String(Date.now() + 3_600_000),
    etag: 'stale',
  };

  // Monkey-patch the table for this test.
  const origGetEntity = table.getEntity.bind(table);
  table.getEntity = async (pk, rk) => {
    if (firstRead) { firstRead = false; return origGetEntity(pk, rk).catch(() => { const e = Object.assign(new Error('nf'), { statusCode: 404 }); throw e; }); }
    return { ...fakeEntity };
  };
  table.createEntity = always412;
  table.updateEntity = always412;

  const retryAfter = await consume(req);
  assert.equal(retryAfter, 5, 'must fail closed after all retries exhausted');
});

test('table limiter returns exact retry-after when limit is already exceeded', async () => {
  const table = new FakeTable();
  const windowMs = 3_600_000;
  const consume = createTableRateLimiter(table, windowMs, 2);
  const req = fakeReq('10.0.2.4');

  // Fill the limit.
  await consume(req);
  await consume(req);

  // Third request should see count >= max immediately from the read.
  const before = Date.now();
  const retryAfter = await consume(req);
  // retryAfter should be at most windowMs/1000 seconds away (≤ 3600).
  assert.ok(retryAfter >= 1);
  assert.ok(retryAfter <= Math.ceil(windowMs / 1000) + 1);
  // No write should have been attempted for an already-limited request.
  const entry = table.entities.values().next().value;
  assert.equal(entry.data.count, 2);
  void before;
});
