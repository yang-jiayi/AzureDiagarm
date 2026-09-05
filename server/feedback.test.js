// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.
const assert = require('node:assert/strict');
const { test } = require('node:test');
const express = require('express');
const { safeContext, createFeedbackService, createFeedbackRouter } = require('./feedback');
const { createAccessControlRouter } = require('./access-control');
const { createOriginGuard } = require('./deployment-security');

const origin = 'https://app.example.com';
const admin = { 'X-MS-CLIENT-PRINCIPAL-NAME': 'admin@example.com', 'X-MS-CLIENT-PRINCIPAL-ID': 'admin-id' };
const member = { 'X-MS-CLIENT-PRINCIPAL-NAME': 'member@example.com', 'X-MS-CLIENT-PRINCIPAL-ID': 'member-id' };
const logger = { error() {}, warn() {}, info() {} };

function fakeTable() {
  const docs = new Map();
  return {
    docs,
    async createEntity(item) { docs.set(item.rowKey, { ...item, etag: 'etag' }); },
    async deleteEntity(_partition, key, options) { assert.ok(options?.etag); docs.delete(key); },
    listEntities(options) {
      const id = options?.queryOptions?.filter?.match(/id eq '([^']+)'/)?.[1];
      return (async function* () {
        for (const doc of docs.values()) if (!id || doc.id === id) yield doc;
      })();
    },
  };
}
async function startServer(service) {
  const app = express();
  const access = createAccessControlRouter({
    enabled: true, adminEmail: 'admin@example.com', publicAppUrl: origin, logger,
    table: { listEntities: () => (async function* () { yield { email: 'member@example.com' }; })() },
  });
  app.use(express.json(), access.requireAllowed, createOriginGuard({ mode: 'public', origin }));
  app.use('/api/feedback', createFeedbackRouter({ service, mode: 'public', adminEmail: 'admin@example.com', logger }));
  const server = await new Promise(resolve => {
    const instance = app.listen(0, '127.0.0.1', () => resolve(instance));
  });
  return { url: `http://127.0.0.1:${server.address().port}/api/feedback`, close: () => new Promise(resolve => server.close(resolve)) };
}

test('metadata defaults off and server drops arbitrary identifiers even with consent', () => {
  const context = { diagramName: 'customer name', userAgent: 'fingerprint', email: 'secret', serviceCount: 8, model: 'gpt-test', url: 'https://user:password@app.example.com/customer?token=secret#private' };
  assert.deepEqual(safeContext({ context }), {});
  assert.deepEqual(safeContext({ includeMetadata: 'true', context }), {});
  assert.deepEqual(safeContext({ includeMetadata: true, context }), { serviceCount: 8, model: 'gpt-test', url: origin });
  assert.deepEqual(safeContext({ includeMetadata: true, context: { url: 'javascript:alert(1)' } }), {});
});

test('feedback uses TTL/retention and signed-in owner/admin deletion, never an arbitrary caller', async (t) => {
  const table = fakeTable();
  const service = createFeedbackService({ table, retentionDays: 30, logger });
  const server = await startServer(service);
  t.after(server.close);
  const body = { rating: 5, comment: 'Great', context: { url: `${origin}/?private=secret`, diagramName: 'customer' } };
  assert.equal((await fetch(server.url, { method: 'POST', headers: { 'Content-Type': 'application/json', Origin: origin }, body: JSON.stringify(body) })).status, 401);
  const response = await fetch(server.url, {
    method: 'POST', headers: { ...member, 'Content-Type': 'application/json', Origin: origin }, body: JSON.stringify(body),
  });
  assert.equal(response.status, 201);
  const receipt = await response.json();
  assert.equal(receipt.canDelete, true);
  const stored = [...table.docs.values()][0];
  assert.ok(Date.parse(stored.expiresAt) > Date.now());
  assert.equal(stored.contextJson, '{}');
  assert.doesNotMatch(JSON.stringify(stored), /member-id|member@example|customer|private/);
  assert.equal((await fetch(`${server.url}/list`, { headers: member })).status, 403);
  assert.equal((await fetch(`${server.url}/list`, { headers: admin })).status, 200);
  assert.equal((await fetch(`${server.url}/${receipt.id}`, { method: 'DELETE', headers: member })).status, 403);
  assert.equal((await fetch(`${server.url}/${receipt.id}`, {
    method: 'DELETE', headers: { ...member, 'X-MS-CLIENT-PRINCIPAL-ID': 'different-id', Origin: origin },
  })).status, 404);
  const deleted = await fetch(`${server.url}/${receipt.id}`, { method: 'DELETE', headers: { ...member, Origin: origin } });
  assert.equal(deleted.status, 200);
  assert.deepEqual(await deleted.json(), { archiveDeleted: true, emailDeleted: false });
  assert.equal(table.docs.size, 0);
});

test('Table retention actually deletes expired and legacy feedback', async () => {
  const table = fakeTable();
  const service = createFeedbackService({ table, retentionDays: 30, legacyRetentionEnabled: true, logger });
  await table.createEntity({ rowKey: 'old', createdAt: '2020-01-01T00:00:00Z' });
  await table.createEntity({ rowKey: 'expired', expiresAt: '2020-01-01T00:00:00Z' });
  await table.createEntity({ rowKey: 'new', createdAt: new Date().toISOString() });
  await service.sweep();
  assert.deepEqual([...table.docs.keys()], ['new']);
});

test('Cosmos writes TTL, rejects disabled TTL and sweeps legacy records', async () => {
  const deleted = [];
  const stored = [];
  const container = {
    read: async () => ({ resource: { partitionKey: { paths: ['/id'] }, defaultTtl: -1 } }),
    items: {
      create: async item => stored.push(item),
      query: () => ({ getAsyncIterator: () => (async function* () { yield { resources: [{ id: 'legacy' }] }; })() }),
    },
    item: id => ({ delete: async () => deleted.push(id) }),
  };
  const service = createFeedbackService({ container, retentionDays: 7, legacyRetentionEnabled: true });
  await service.save({ id: 'new', createdAt: new Date().toISOString() });
  assert.equal(stored[0].ttl, 7 * 86400);
  await service.sweep();
  assert.deepEqual(deleted, ['legacy']);
  await assert.rejects(createFeedbackService({ container: { read: async () => ({ resource: {} }) } }).validate(), /enabled TTL/);
});

test('email-only policy never promises archive deletion and archive errors are not hidden by email', async (t) => {
  const service = createFeedbackService({
    emailClient: { beginSend: async () => ({ pollUntilDone: async () => ({ status: 'Succeeded' }) }) },
    emailSender: 'sender@example.com', emailRecipient: 'recipient@example.com',
  });
  const server = await startServer(service);
  t.after(server.close);
  const policy = await (await fetch(`${server.url}/policy`, { headers: member })).json();
  assert.equal(policy.archiveEnabled, false);
  assert.equal(policy.emailDeletionSupported, false);
  const response = await fetch(`${server.url}/11111111-1111-1111-1111-111111111111`, { method: 'DELETE', headers: { ...member, Origin: origin } });
  assert.equal(response.status, 409);
  const failing = createFeedbackService({
    table: { createEntity: async () => { throw new Error('Archive unavailable'); } },
    emailClient: { beginSend: async () => { throw new Error('Must not send'); } },
  });
  await assert.rejects(failing.save({ createdAt: new Date().toISOString(), context: {} }), /Archive unavailable/);
});

test('legacy retention is opt-in while newly expiring feedback is still removed', async () => {
  const table = fakeTable();
  let cosmosQueries = 0;
  const service = createFeedbackService({
    table, logger,
    container: { items: { query() { cosmosQueries++; throw new Error('Legacy sweep must stay off'); } } },
  });
  await table.createEntity({ rowKey: 'legacy', createdAt: '2020-01-01T00:00:00Z' });
  await table.createEntity({ rowKey: 'unknown-age' });
  await table.createEntity({ rowKey: 'expired-new', expiresAt: '2020-01-01T00:00:00Z' });
  await table.createEntity({ rowKey: 'unexpired', expiresAt: '2099-01-01T00:00:00Z' });
  await service.sweep();
  assert.deepEqual([...table.docs.keys()], ['legacy', 'unknown-age', 'unexpired']);
  assert.equal(cosmosQueries, 0);
  assert.equal(service.legacyRetentionEnabled, false);
});

test('follow-up consent is server-gated and addresses go only to email, never archives', async (t) => {
  const table = fakeTable();
  const delivered = [];
  const service = createFeedbackService({
    table, logger, contactEnabled: true,
    emailClient: {
      async beginSend(message) {
        delivered.push(message);
        return { pollUntilDone: async () => ({ status: 'Succeeded' }) };
      },
    },
    emailSender: 'sender@example.com', emailRecipient: 'recipient@example.com',
  });
  const server = await startServer(service);
  t.after(server.close);
  const send = contact => fetch(server.url, {
    method: 'POST', headers: { ...member, Origin: origin, 'Content-Type': 'application/json' },
    body: JSON.stringify({
      rating: 5, category: 'General', comment: 'A useful feature',
      contact, context: { diagramName: 'private', model: 'private-model' },
    }),
  });
  assert.equal((await send({ consent: true, email: 'invalid' })).status, 400);
  assert.equal((await send({ consent: true, email: ' FollowUp@Example.com ' })).status, 201);
  assert.match(delivered[0].content.plainText, /followup@example\.com/);
  const saved = [...table.docs.values()][0];
  assert.doesNotMatch(JSON.stringify(saved), /followup@example|FollowUp@Example|private|member-id/);
  assert.equal(JSON.parse(saved.contactJson).consent, true);
  assert.equal(JSON.parse(saved.contactJson).followUpStatus, 'new');
  assert.ok(Date.parse(JSON.parse(saved.contactJson).expiresAt) > Date.now());
  assert.equal((await send({ consent: false, email: 'ignored@example.com' })).status, 201);
  assert.doesNotMatch(JSON.stringify([...table.docs.values(), ...delivered]), /ignored@example/);
  const policy = await (await fetch(`${server.url}/policy`, { headers: member })).json();
  assert.equal(policy.contactEnabled, true);
  assert.equal(policy.legacyRetentionEnabled, false);
});

test('contact is unavailable without delivery and failed email never archives a follow-up address', async (t) => {
  for (const [contactEnabled, emailClient, expected] of [
    [false, null, 400],
    [true, null, 503],
    [true, { beginSend: async () => { throw new Error('Email unavailable'); } }, 503],
  ]) {
    const table = fakeTable();
    const service = createFeedbackService({ table, logger, contactEnabled, emailClient });
    const server = await startServer(service);
    t.after(server.close);
    const response = await fetch(server.url, {
      method: 'POST', headers: { ...member, Origin: origin, 'Content-Type': 'application/json' },
      body: JSON.stringify({ rating: 5, contact: { consent: true, email: 'person@example.com' } }),
    });
    assert.equal(response.status, expected);
    assert.equal(table.docs.size, 0);
  }
});

test('Cosmos fallback retains deletion ownership and strips contact addresses', async () => {
  const docs = new Map();
  const container = {
    read: async () => ({ resource: { partitionKey: { paths: ['/id'] }, defaultTtl: -1 } }),
    items: { create: async item => docs.set(item.id, item) },
    item: id => ({
      read: async () => ({ resource: docs.get(id) }),
      delete: async () => docs.delete(id),
    }),
  };
  const service = createFeedbackService({
    logger, container,
    table: {
      createEntity: async () => { throw new Error('Table unavailable'); },
      listEntities: () => (async function* () { throw new Error('Table unavailable'); })(),
    },
  });
  await service.save({
    id: 'fallback', type: 'feedback', createdAt: new Date().toISOString(),
    ownerHash: 'hashed-owner', contact: { consent: false, email: 'never-store@example.com' },
  });
  const saved = await service.find('fallback');
  assert.equal(saved.ownerHash, 'hashed-owner');
  assert.doesNotMatch(JSON.stringify(saved), /never-store@example/);
  assert.equal(saved.ttl, 30 * 86400);
  await service.remove(saved);
  assert.equal(docs.size, 0);
  docs.set('budget', { id: 'budget', type: 'ai-budget' });
  await assert.rejects(service.find('budget'), /Table unavailable/);
});
