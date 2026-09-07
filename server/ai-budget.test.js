// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.
const assert = require('node:assert/strict');
const { test } = require('node:test');
const {
  MemoryBudgetStore, CosmosBudgetStore, TableBudgetStore, createBudgetManager, budgetIdentity, reservationTokens, actualUsage, hasUnmeteredInput,
} = require('./ai-budget');

function fakeCosmos() {
  const docs = new Map();
  let revision = 0;
  const missing = () => Object.assign(new Error('Not found'), { code: 404 });
  return {
    docs,
    read: async () => ({ resource: { partitionKey: { paths: ['/id'] }, defaultTtl: -1 } }),
    items: {
      async create(doc) {
        if (docs.has(doc.id)) throw Object.assign(new Error('Conflict'), { code: 409 });
        docs.set(doc.id, structuredClone({ ...doc, _etag: String(++revision) }));
      },
    },
    item(id, partitionKey) {
      assert.equal(partitionKey, id);
      return {
        async read() {
          if (!docs.has(id)) throw missing();
          return { resource: structuredClone(docs.get(id)) };
        },
        async replace(doc, options) {
          assert.equal(options.accessCondition.type, 'IfMatch');
          if (docs.get(id)?._etag !== options.accessCondition.condition) {
            throw Object.assign(new Error('Stale ETag'), { code: 412 });
          }
          docs.set(id, structuredClone({ ...doc, _etag: String(++revision) }));
        },
      };
    },
  };
}

test('independent Cosmos-backed replicas atomically enforce token and concurrency limits', async () => {
  const container = fakeCosmos();
  const replicas = [0, 1].map(() => createBudgetManager({
    store: new CosmosBudgetStore(container), dailyTokens: 1000, concurrency: 3,
  }));
  const requests = await Promise.allSettled(Array.from({ length: 15 }, (_, i) => replicas[i % 2].reserve('user', 200)));
  const accepted = requests.filter(request => request.status === 'fulfilled').map(request => request.value);
  assert.equal(accepted.length, 3);
  assert.equal((await replicas[1].status('user')).remainingTokens, 400);
  assert.ok(requests.filter(request => request.status === 'rejected').every(request => request.reason.retryAfter > 0));
  await Promise.all(accepted.map(lease => replicas[0].settle('user', lease, 50)));
  const final = await replicas[1].status('user');
  assert.equal(final.concurrentRequests, 0);
  assert.equal(final.usedTokens, 150);
  assert.equal(final.remainingTokens, 850);
  assert.equal([...container.docs.values()][0].ttl, 259200);
  assert.doesNotMatch(JSON.stringify([...container.docs.values()]), /"user"/);
});

test('daily budget rejects before dispatch; expiry frees concurrency without refunding uncertain usage', async () => {
  let time = Date.parse('2026-09-05T10:00:00Z');
  const manager = createBudgetManager({ store: new MemoryBudgetStore(), dailyTokens: 100, concurrency: 1, now: () => time, leaseMs: 1000 });
  const lease = await manager.reserve('one', 80);
  await assert.rejects(manager.reserve('one', 1), { code: 'ai_concurrency_limit', status: 429 });
  time += 1001;
  assert.equal((await manager.status('one')).concurrentRequests, 0);
  await assert.rejects(manager.reserve('one', 30), { code: 'ai_daily_budget_exceeded', status: 429 });
  await manager.settle('one', lease, 0);
  assert.equal((await manager.status('one')).usedTokens, 80, 'expired lease cannot manufacture a late refund');
  assert.equal((await manager.status('other')).remainingTokens, 100);
  time = Date.parse('2026-09-06T00:00:00Z');
  assert.equal((await manager.status('one')).remainingTokens, 100);
});

test('midnight resets tokens but cannot bypass active concurrency; old usage cannot refund a new day', async () => {
  let time = Date.parse('2026-09-05T23:59:59Z');
  const manager = createBudgetManager({ store: new MemoryBudgetStore(), dailyTokens: 100, concurrency: 2, now: () => time });
  const yesterday = await manager.reserve('user', 80);
  time += 2000;
  const today = await manager.reserve('user', 90);
  await assert.rejects(manager.reserve('user', 1), { code: 'ai_concurrency_limit' });
  await manager.settle('user', yesterday, 1);
  assert.equal((await manager.status('user')).usedTokens, 90);
  await manager.settle('user', today, 15);
  await manager.settle('user', today, 0);
  assert.equal((await manager.status('user')).usedTokens, 15, 'settlement is idempotent');
  time -= 2000;
  assert.equal((await manager.status('user')).usedTokens, 15, 'a replica behind midnight cannot reset yesterday');
});

test('usage over reservation is debited and unknown usage keeps reservation', async () => {
  const manager = createBudgetManager({ store: new MemoryBudgetStore(), dailyTokens: 1000 });
  const a = await manager.reserve('user', 100);
  await manager.settle('user', a, 120);
  const b = await manager.reserve('user', 100);
  await manager.settle('user', b, undefined);
  assert.equal((await manager.status('user')).usedTokens, 220);
  assert.equal((await manager.status('user')).concurrentRequests, 0);
});

test('storage failures do not silently fall back to memory', async () => {
  const manager = createBudgetManager({ store: { async read() { throw new Error('Storage unavailable'); } } });
  await assert.rejects(manager.status('user'), /Storage unavailable/);
  await assert.rejects(manager.reserve('user', 100), /Storage unavailable/);
  await assert.rejects(new CosmosBudgetStore({ read: async () => ({ resource: { partitionKey: { paths: ['/id'] } } }) }).validate(), /enabled TTL/);
  await assert.rejects(new CosmosBudgetStore(fakeCosmos(), {
    getDatabaseAccount: async () => ({ resource: { enableMultipleWritableLocations: true, writableLocations: [{}, {}] } }),
  }).validate(), /exactly one write region/);
});

test('Table shared store requires conditional ETags and expiration deletes are conditional', async () => {
  const writes = [];
  const table = {
    getEntity: async () => ({ document: '{"id":"one"}', etag: 'etag-1' }),
    updateEntity: async (...args) => writes.push(args),
    createEntity: async (...args) => writes.push(args),
    listEntities: () => (async function* () { yield { rowKey: 'one', etag: 'etag-1' }; })(),
    deleteEntity: async (...args) => writes.push(args),
  };
  const store = new TableBudgetStore(table);
  assert.equal((await store.read('one'))._etag, 'etag-1');
  await store.write({ id: 'one', expiresAt: 10 }, 'etag-1');
  assert.deepEqual(writes[0].slice(1), ['Replace', { etag: 'etag-1' }]);
  await store.sweep(20);
  assert.deepEqual(writes[1], ['ai-budget', 'one', { etag: 'etag-1' }]);
  const missing = new TableBudgetStore({
    getEntity: async () => { throw Object.assign(new Error('Not found'), { statusCode: 404 }); },
    listEntities: () => ({ byPage: () => ({ next: async () => { throw new Error('Table unavailable'); } }) }),
  });
  await assert.rejects(missing.read('new-user'), /Table unavailable/);
});

test('public identity is required; reservations include input/output and vision; usage is validated', () => {
  assert.throws(() => budgetIdentity({ get: () => '' }, 'public'), { status: 401 });
  assert.equal(budgetIdentity({}, 'local'), 'local-development');
  assert.ok(reservationTokens({ input: '日本語', max_output_tokens: 100 }, 'responses') > 4100);
  assert.ok(reservationTokens({ input: [{ type: 'input_image', image_url: 'data:image/png;base64,AA==' }], max_output_tokens: 1 }, 'responses') > 131072);
  assert.equal(
    reservationTokens({ input: [{ type: 'input_image', image_url: 'data:image/png;base64,AA==' }], max_output_tokens: 1 }, 'responses'),
    reservationTokens({ input: [{ type: 'input_image', image_url: `data:image/png;base64,${'A'.repeat(100_000)}` }], max_output_tokens: 1 }, 'responses'),
    'binary image transport bytes are covered by the vision reservation, not counted as text',
  );
  assert.equal(actualUsage({ usage: { input_tokens: 10, output_tokens: 20 } }), 30);
  assert.equal(actualUsage({ usage: { prompt_tokens: 10, completion_tokens: 20 } }), 30);
  assert.equal(actualUsage({ usage: { prompt_tokens: 10 } }), undefined);
  assert.equal(actualUsage({ usage: { completion_tokens: 20 } }), undefined);
  assert.equal(actualUsage({ usage: { prompt_tokens: '10', completion_tokens: 20 } }), undefined);
  assert.equal(actualUsage({ usage: { total_tokens: -1 } }), undefined);
  assert.equal(actualUsage({ usage: {} }), undefined);
  assert.equal(hasUnmeteredInput({ input: [{ type: 'input_image', image_url: 'https://remote.example/image' }] }), true);
  assert.equal(hasUnmeteredInput({ input: [{ type: 'item_reference', id: 'stored' }] }), true);
  assert.equal(hasUnmeteredInput({ input: [{ type: 'input_image', image_url: 'data:image/png;base64,AA==' }] }), false);
});

test('Chat Completions inline vision budgets exclude base64 bytes and include the selected output cap', () => {
  const chat = data => ({
    messages: [{ role: 'user', content: [{ type: 'image_url', image_url: { url: `data:image/png;base64,${data}`, detail: 'high' } }] }],
    max_tokens: 32000,
  });
  assert.equal(reservationTokens(chat('AA=='), 'chat-completions'), reservationTokens(chat('A'.repeat(100_000)), 'chat-completions'));
  assert.ok(reservationTokens(chat('AA=='), 'chat-completions') > 163000);
  assert.equal(hasUnmeteredInput(chat('AA==')), false);
});

test('reasoning chat and Anthropic vision/cache usage remain bounded and metered', () => {
  assert.ok(reservationTokens({ max_completion_tokens: 32000, messages: [] }, 'chat-completions') > 36000);
  const anthropic = data => ({
    max_tokens: 1,
    messages: [{ role: 'user', content: [{ type: 'image', source: { type: 'base64', media_type: 'image/png', data } }] }],
  });
  assert.ok(reservationTokens(anthropic('AA=='), 'anthropic-messages') > 131072);
  assert.equal(
    reservationTokens(anthropic('AA=='), 'anthropic-messages'),
    reservationTokens(anthropic('A'.repeat(100000)), 'anthropic-messages'),
  );
  assert.equal(hasUnmeteredInput(anthropic('AA==')), false);
  assert.equal(hasUnmeteredInput({ messages: [{ content: [{ type: 'image', source: { type: 'url', url: 'https://example.test/image' } }] }] }), true);
  assert.equal(hasUnmeteredInput({ messages: [{ content: [{ type: 'document', source: { type: 'file', file_id: 'stored' } }] }] }), true);
  assert.equal(actualUsage({ usage: { input_tokens: 10, output_tokens: 5, cache_read_input_tokens: 20, cache_creation_input_tokens: 30 } }), 65);
});
