// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

const assert = require('node:assert/strict');
const { test } = require('node:test');
const express = require('express');
const { createAzureBlobBackend, createDiagramsRouter } = require('./diagram-api');
const { getPrincipal } = require('./access-control');
const { createErrorHandler } = require('./async-handler');

// In-memory backend that mirrors the semantics the Azure Blob backend exposes:
// create-only writes fail 409 when the blob exists, conditional replaces fail
// 412 on an ETag mismatch, and reads return a fresh deep copy so router-side
// mutations never leak into stored state before an explicit write.
class FakeBackend {
  constructor(options = {}) {
    this.map = new Map();
    this.counter = 0;
    this.locks = new Map();
    this.afterLockOperation = options.afterLockOperation || null;
  }

  nextEtag() {
    this.counter += 1;
    return `"etag-${this.counter}"`;
  }

  async read(name) {
    const record = this.map.get(name);
    if (!record) return null;
    return { value: JSON.parse(record.json), etag: record.etag };
  }

  async create(name, value) {
    if (this.map.has(name)) {
      const error = new Error('BlobAlreadyExists');
      error.statusCode = 409;
      throw error;
    }
    const etag = this.nextEtag();
    this.map.set(name, { json: JSON.stringify(value), etag });
    return { etag };
  }

  async replace(name, value, etag) {
    const record = this.map.get(name);
    if (!record) {
      const error = new Error('BlobNotFound');
      error.statusCode = 404;
      throw error;
    }
    if (record.etag !== etag) {
      const error = new Error('ConditionNotMet');
      error.statusCode = 412;
      throw error;
    }
    const next = this.nextEtag();
    this.map.set(name, { json: JSON.stringify(value), etag: next });
    return { etag: next };
  }

  async put(name, value) {
    const etag = this.nextEtag();
    this.map.set(name, { json: JSON.stringify(value), etag });
    return { etag };
  }

  async remove(name, etag) {
    const record = this.map.get(name);
    if (!record) return false;
    if (etag && record.etag !== etag) {
      const error = new Error('ConditionNotMet');
      error.statusCode = 412;
      throw error;
    }
    return this.map.delete(name);
  }

  async* list(prefix) {
    for (const key of Array.from(this.map.keys())) {
      if (key.startsWith(prefix)) yield key;
    }
  }

  async withLock(name, operation) {
    const previous = this.locks.get(name) || Promise.resolve();
    let release;
    const gate = new Promise((resolve) => {
      release = resolve;
    });
    const queued = previous.then(() => gate);
    this.locks.set(name, queued);
    await previous;
    try {
      const result = await operation();
      if (this.afterLockOperation) await this.afterLockOperation(name, result);
      return result;
    } finally {
      release();
      if (this.locks.get(name) === queued) this.locks.delete(name);
    }
  }
}

class ConcurrentUpdateBeforeDeleteBackend extends FakeBackend {
  constructor() {
    super();
    this.injected = false;
  }

  async replace(name, value, etag) {
    if (value?.kind === 'diagram-deletion-marker-v1' && !this.injected) {
      const record = this.map.get(name);
      if (record) {
        const concurrentValue = JSON.parse(record.json);
        concurrentValue.diagramName = 'Concurrent update';
        concurrentValue.revision += 1;
        this.map.set(name, { json: JSON.stringify(concurrentValue), etag: this.nextEtag() });
        this.injected = true;
      }
    }
    return super.replace(name, value, etag);
  }
}

class FailVersionCleanupOnceBackend extends FakeBackend {
  constructor() {
    super();
    this.failed = false;
  }

  async remove(name, etag) {
    if (!etag && name.includes('/versions/') && !this.failed) {
      this.failed = true;
      const error = new Error('Transient cleanup failure');
      error.statusCode = 503;
      throw error;
    }
    return super.remove(name, etag);
  }
}

class BlockingVersionWriteBackend extends FakeBackend {
  constructor() {
    super();
    this.versionWriteStarted = new Promise((resolve) => {
      this.signalVersionWriteStarted = resolve;
    });
    this.releaseVersionWrite = null;
    this.versionWriteGate = new Promise((resolve) => {
      this.releaseVersionWrite = resolve;
    });
  }

  async put(name, value) {
    if (name.includes('/versions/')) {
      this.signalVersionWriteStarted();
      await this.versionWriteGate;
    }
    return super.put(name, value);
  }
}

const SILENT = { error() {}, warn() {}, info() {}, log() {} };

async function startServer({ backend }) {
  const app = express();
  app.use('/api/diagrams', express.json({ limit: '12mb' }));
  app.use('/api/diagrams', createDiagramsRouter({
    backend,
    getPrincipal,
    publicUrl: 'https://app.example',
    logger: SILENT,
  }));
  app.use(createErrorHandler(SILENT));
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

const USER_A = {
  'X-MS-CLIENT-PRINCIPAL-NAME': 'alice@example.com',
  'X-MS-CLIENT-PRINCIPAL-ID': '11111111-1111-1111-1111-111111111111',
};
const USER_B = {
  'X-MS-CLIENT-PRINCIPAL-NAME': 'bob@example.com',
  'X-MS-CLIENT-PRINCIPAL-ID': '22222222-2222-2222-2222-222222222222',
};

function samplePayload(overrides = {}) {
  return { nodes: [{ id: 'n1' }, { id: 'n2' }], edges: [{ id: 'e1', source: 'n1', target: 'n2' }], ...overrides };
}

async function call(baseUrl, method, path, { headers = {}, body, ifMatch } = {}) {
  const finalHeaders = { ...headers };
  if (body !== undefined) finalHeaders['Content-Type'] = 'application/json';
  if (ifMatch) finalHeaders['If-Match'] = ifMatch;
  const res = await fetch(`${baseUrl}${path}`, {
    method,
    headers: finalHeaders,
    body: body === undefined ? undefined : JSON.stringify(body),
  });
  let json = null;
  const text = await res.text();
  if (text) {
    try { json = JSON.parse(text); } catch { json = text; }
  }
  return { status: res.status, etag: res.headers.get('etag'), json };
}

async function createDoc(baseUrl, headers, name = 'My Diagram') {
  const res = await call(baseUrl, 'POST', '/api/diagrams', {
    headers,
    body: { diagramName: name, payload: samplePayload() },
  });
  assert.equal(res.status, 201, `create failed: ${JSON.stringify(res.json)}`);
  return res;
}
test('authentication is required for diagram routes', async (t) => {
  const server = await startServer({ backend: new FakeBackend() });
  t.after(server.close);
  const res = await call(server.baseUrl, 'GET', '/api/diagrams');
  assert.equal(res.status, 401);
});

test('returns 503 when storage is not configured', async (t) => {
  const server = await startServer({ backend: null });
  t.after(server.close);
  const res = await call(server.baseUrl, 'GET', '/api/diagrams', { headers: USER_A });
  assert.equal(res.status, 503);
});

test('create, list, and get a document', async (t) => {
  const server = await startServer({ backend: new FakeBackend() });
  t.after(server.close);

  const created = await createDoc(server.baseUrl, USER_A, 'Alpha');
  assert.ok(created.etag, 'create returns an ETag header');
  const id = created.json.document.id;
  assert.equal(created.json.document.diagramName, 'Alpha');
  assert.equal(created.json.document.revision, 1);
  assert.equal(created.json.document.owner.email, 'alice@example.com');

  const list = await call(server.baseUrl, 'GET', '/api/diagrams', { headers: USER_A });
  assert.equal(list.status, 200);
  assert.equal(list.json.documents.length, 1);
  assert.equal(list.json.documents[0].id, id);
  assert.equal(list.json.documents[0].serviceCount, 2);
  assert.equal(list.json.documents[0].access, 'owner');

  const got = await call(server.baseUrl, 'GET', `/api/diagrams/${id}`, { headers: USER_A });
  assert.equal(got.status, 200);
  assert.ok(got.etag);
  assert.equal(got.json.document.access, 'owner');
});

test('create rejects invalid payloads', async (t) => {
  const server = await startServer({ backend: new FakeBackend() });
  t.after(server.close);

  const noArrays = await call(server.baseUrl, 'POST', '/api/diagrams', {
    headers: USER_A, body: { diagramName: 'x', payload: { foo: 1 } },
  });
  assert.equal(noArrays.status, 400);

  const noName = await call(server.baseUrl, 'POST', '/api/diagrams', {
    headers: USER_A, body: { diagramName: '   ', payload: samplePayload() },
  });
  assert.equal(noName.status, 400);
});

test('create replays an idempotent request without duplicating the document', async (t) => {
  const server = await startServer({ backend: new FakeBackend() });
  t.after(server.close);
  const headers = {
    ...USER_A,
    'Idempotency-Key': 'create-request-1234567890',
  };
  const body = { diagramName: 'Idempotent', payload: samplePayload() };

  const first = await call(server.baseUrl, 'POST', '/api/diagrams', { headers, body });
  assert.equal(first.status, 201);

  const replay = await call(server.baseUrl, 'POST', '/api/diagrams', { headers, body });
  assert.equal(replay.status, 200);
  assert.equal(replay.json.document.id, first.json.document.id);

  const list = await call(server.baseUrl, 'GET', '/api/diagrams', { headers: USER_A });
  assert.equal(list.status, 200);
  assert.equal(list.json.documents.length, 1);
});

test('create rejects reusing an idempotency key with different content', async (t) => {
  const server = await startServer({ backend: new FakeBackend() });
  t.after(server.close);
  const headers = {
    ...USER_A,
    'Idempotency-Key': 'create-request-abcdefghij',
  };

  const first = await call(server.baseUrl, 'POST', '/api/diagrams', {
    headers,
    body: { diagramName: 'First', payload: samplePayload() },
  });
  assert.equal(first.status, 201);

  const mismatch = await call(server.baseUrl, 'POST', '/api/diagrams', {
    headers,
    body: {
      diagramName: 'Second',
      payload: samplePayload({ nodes: [{ id: 'different' }], edges: [] }),
    },
  });
  assert.equal(mismatch.status, 409);
});

test('owner isolation: another user cannot access a document by id', async (t) => {
  const server = await startServer({ backend: new FakeBackend() });
  t.after(server.close);

  const created = await createDoc(server.baseUrl, USER_A);
  const id = created.json.document.id;

  const bGet = await call(server.baseUrl, 'GET', `/api/diagrams/${id}`, { headers: USER_B });
  assert.equal(bGet.status, 404);

  const bPut = await call(server.baseUrl, 'PUT', `/api/diagrams/${id}`, {
    headers: USER_B, ifMatch: created.etag, body: { diagramName: 'hax', payload: samplePayload() },
  });
  assert.equal(bPut.status, 404);

  const bList = await call(server.baseUrl, 'GET', '/api/diagrams', { headers: USER_B });
  assert.equal(bList.json.documents.length, 0);
});

test('update requires If-Match and enforces optimistic concurrency', async (t) => {
  const server = await startServer({ backend: new FakeBackend() });
  t.after(server.close);

  const created = await createDoc(server.baseUrl, USER_A);
  const id = created.json.document.id;

  const missing = await call(server.baseUrl, 'PUT', `/api/diagrams/${id}`, {
    headers: USER_A, body: { diagramName: 'v2', payload: samplePayload() },
  });
  assert.equal(missing.status, 428);

  const ok = await call(server.baseUrl, 'PUT', `/api/diagrams/${id}`, {
    headers: USER_A, ifMatch: created.etag, body: { diagramName: 'v2', payload: samplePayload({ nodes: [{ id: 'n1' }] }) },
  });
  assert.equal(ok.status, 200);
  assert.equal(ok.json.document.diagramName, 'v2');
  assert.equal(ok.json.document.revision, 2);
  assert.notEqual(ok.etag, created.etag);

  const stale = await call(server.baseUrl, 'PUT', `/api/diagrams/${id}`, {
    headers: USER_A, ifMatch: created.etag, body: { diagramName: 'v3', payload: samplePayload() },
  });
  assert.equal(stale.status, 412);
});

test('update preserves server-owned fields', async (t) => {
  const server = await startServer({ backend: new FakeBackend() });
  t.after(server.close);

  const created = await createDoc(server.baseUrl, USER_A);
  const id = created.json.document.id;
  const createdAt = created.json.document.createdAt;

  const comment = await call(server.baseUrl, 'POST', `/api/diagrams/${id}/comments`, {
    headers: USER_A, body: { message: 'first comment' },
  });
  assert.equal(comment.status, 201);

  const updated = await call(server.baseUrl, 'PUT', `/api/diagrams/${id}`, {
    headers: USER_A, ifMatch: comment.etag, body: { diagramName: 'renamed', payload: samplePayload() },
  });
  assert.equal(updated.status, 200);
  assert.equal(updated.json.document.createdAt, createdAt);
  assert.equal(updated.json.document.comments.length, 1);
  assert.equal(updated.json.document.owner.email, 'alice@example.com');
});

test('immutable versions can be created, listed and fetched', async (t) => {
  const server = await startServer({ backend: new FakeBackend() });
  t.after(server.close);

  const created = await createDoc(server.baseUrl, USER_A);
  const id = created.json.document.id;

  const snap = await call(server.baseUrl, 'POST', `/api/diagrams/${id}/versions`, {
    headers: USER_A, body: { notes: 'baseline' },
  });
  assert.equal(snap.status, 201);
  const versionId = snap.json.version.versionId;
  assert.equal(snap.json.version.notes, 'baseline');
  assert.equal(snap.json.version.sourceRevision, 1);

  const list = await call(server.baseUrl, 'GET', `/api/diagrams/${id}/versions`, { headers: USER_A });
  assert.equal(list.status, 200);
  assert.equal(list.json.versions.length, 1);
  assert.equal(list.json.versions[0].versionId, versionId);

  const one = await call(server.baseUrl, 'GET', `/api/diagrams/${id}/versions/${versionId}`, { headers: USER_A });
  assert.equal(one.status, 200);
  assert.equal(one.json.version.versionId, versionId);
  assert.ok(Array.isArray(one.json.version.payload.nodes));
});

test('version creation waits for lock completion before sending success', async (t) => {
  const backend = new FakeBackend({
    afterLockOperation() {
      const error = new Error('Lock ownership could not be confirmed');
      error.statusCode = 503;
      throw error;
    },
  });
  const server = await startServer({ backend });
  t.after(server.close);

  const created = await createDoc(server.baseUrl, USER_A);
  const res = await call(
    server.baseUrl,
    'POST',
    `/api/diagrams/${created.json.document.id}/versions`,
    { headers: USER_A, body: { notes: 'late lock failure' } },
  );
  assert.equal(res.status, 503);
  assert.deepEqual(res.json, { error: 'Diagram storage is temporarily unavailable' });
});

test('comments are added and bounded to the document', async (t) => {
  const server = await startServer({ backend: new FakeBackend() });
  t.after(server.close);

  const created = await createDoc(server.baseUrl, USER_A);
  const id = created.json.document.id;

  const added = await call(server.baseUrl, 'POST', `/api/diagrams/${id}/comments`, {
    headers: USER_A, body: { message: 'looks good' },
  });
  assert.equal(added.status, 201);
  assert.equal(added.json.document.comments.length, 1);
  assert.equal(added.json.document.comments[0].authorEmail, 'alice@example.com');
  assert.equal(added.json.document.comments[0].message, 'looks good');
  assert.equal(added.json.document.revision, 2);
  assert.ok(added.etag);

  const empty = await call(server.baseUrl, 'POST', `/api/diagrams/${id}/comments`, {
    headers: USER_A, body: { message: '' },
  });
  assert.equal(empty.status, 400);
});

test('sharing: viewer can read but not edit, editor can edit', async (t) => {
  const backend = new FakeBackend();
  const server = await startServer({ backend });
  t.after(server.close);

  const created = await createDoc(server.baseUrl, USER_A);
  const id = created.json.document.id;

  // Create a viewer share.
  const viewerShare = await call(server.baseUrl, 'POST', `/api/diagrams/${id}/shares`, {
    headers: USER_A, body: { role: 'viewer' },
  });
  assert.equal(viewerShare.status, 201);
  const viewerToken = viewerShare.json.token;
  assert.match(viewerToken, /^[A-Za-z0-9_-]{43}$/);
  assert.equal(viewerShare.json.url, `https://app.example/#share-${viewerToken}`);
  assert.equal(viewerShare.json.share.tokenHash, undefined);
  assert.ok(viewerShare.json.share.shareId);

  // A different authenticated user resolves the share.
  const sharedGet = await call(server.baseUrl, 'GET', `/api/diagrams/shared/${viewerToken}`, { headers: USER_B });
  assert.equal(sharedGet.status, 200);
  assert.equal(sharedGet.json.document.access, 'shared');
  assert.equal(sharedGet.json.document.role, 'viewer');
  assert.equal(sharedGet.json.document.revision, 2);
  assert.ok(sharedGet.etag);

  const ownerGetThroughShare = await call(
    server.baseUrl,
    'GET',
    `/api/diagrams/shared/${viewerToken}`,
    { headers: USER_A },
  );
  assert.equal(ownerGetThroughShare.status, 200);
  assert.equal(ownerGetThroughShare.json.document.access, 'owner');
  assert.equal(ownerGetThroughShare.json.document.role, 'owner');
  assert.equal(ownerGetThroughShare.json.document.owner.id, USER_A['X-MS-CLIENT-PRINCIPAL-ID']);
  assert.equal(ownerGetThroughShare.json.document.shares.length, 1);

  // Viewer cannot modify.
  const viewerPut = await call(server.baseUrl, 'PUT', `/api/diagrams/shared/${viewerToken}`, {
    headers: USER_B, ifMatch: sharedGet.etag, body: { diagramName: 'nope', payload: samplePayload() },
  });
  assert.equal(viewerPut.status, 403);

  // Viewer can comment.
  const viewerComment = await call(server.baseUrl, 'POST', `/api/diagrams/shared/${viewerToken}/comments`, {
    headers: USER_B, body: { message: 'from a viewer' },
  });
  assert.equal(viewerComment.status, 201);
  assert.equal(viewerComment.json.document.comments.at(-1).authorEmail, 'bob@example.com');
  assert.equal(viewerComment.json.document.revision, 3);

  // Create an editor share and confirm edit access.
  const editorShare = await call(server.baseUrl, 'POST', `/api/diagrams/${id}/shares`, {
    headers: USER_A, body: { role: 'editor' },
  });
  const editorToken = editorShare.json.token;
  const editorGet = await call(server.baseUrl, 'GET', `/api/diagrams/shared/${editorToken}`, { headers: USER_B });
  assert.equal(editorGet.json.document.role, 'editor');
  assert.equal(editorGet.json.document.revision, 4);
  const editorPut = await call(server.baseUrl, 'PUT', `/api/diagrams/shared/${editorToken}`, {
    headers: USER_B, ifMatch: editorGet.etag, body: { diagramName: 'edited by editor', payload: samplePayload() },
  });
  assert.equal(editorPut.status, 200);
  assert.equal(editorPut.json.document.diagramName, 'edited by editor');

  // Editor may snapshot; viewer may not.
  const editorSnap = await call(server.baseUrl, 'POST', `/api/diagrams/shared/${editorToken}/versions`, {
    headers: USER_B, body: { notes: 'editor snapshot' },
  });
  assert.equal(editorSnap.status, 201);
  const viewerSnap = await call(server.baseUrl, 'POST', `/api/diagrams/shared/${viewerToken}/versions`, {
    headers: USER_B, body: {},
  });
  assert.equal(viewerSnap.status, 403);
});

test('shared version creation waits for lock completion before sending success', async (t) => {
  const backend = new FakeBackend();
  const server = await startServer({ backend });
  t.after(server.close);

  const created = await createDoc(server.baseUrl, USER_A);
  const share = await call(
    server.baseUrl,
    'POST',
    `/api/diagrams/${created.json.document.id}/shares`,
    { headers: USER_A, body: { role: 'editor' } },
  );
  backend.afterLockOperation = () => {
    const error = new Error('Lock ownership could not be confirmed');
    error.statusCode = 503;
    throw error;
  };

  const res = await call(
    server.baseUrl,
    'POST',
    `/api/diagrams/shared/${share.json.token}/versions`,
    { headers: USER_B, body: { notes: 'late lock failure' } },
  );
  assert.equal(res.status, 503);
  assert.deepEqual(res.json, { error: 'Diagram storage is temporarily unavailable' });
});

test('shared routes require an authenticated app user', async (t) => {
  const server = await startServer({ backend: new FakeBackend() });
  t.after(server.close);

  const created = await createDoc(server.baseUrl, USER_A);
  const id = created.json.document.id;
  const share = await call(server.baseUrl, 'POST', `/api/diagrams/${id}/shares`, {
    headers: USER_A, body: { role: 'viewer' },
  });
  const anon = await call(server.baseUrl, 'GET', `/api/diagrams/shared/${share.json.token}`);
  assert.equal(anon.status, 401);
});

test('raw share tokens are never persisted or listed', async (t) => {
  const backend = new FakeBackend();
  const server = await startServer({ backend });
  t.after(server.close);

  const created = await createDoc(server.baseUrl, USER_A);
  const id = created.json.document.id;
  const share = await call(server.baseUrl, 'POST', `/api/diagrams/${id}/shares`, {
    headers: USER_A, body: { role: 'editor' },
  });
  const token = share.json.token;

  for (const record of backend.map.values()) {
    assert.ok(!record.json.includes(token), 'raw token must not be stored in any blob');
  }

  const shares = await call(server.baseUrl, 'GET', `/api/diagrams/${id}/shares`, { headers: USER_A });
  assert.equal(shares.status, 200);
  assert.equal(shares.json.shares.length, 1);
  assert.equal(shares.json.shares[0].tokenHash, undefined);
  assert.equal(shares.json.shares[0].role, 'editor');
  assert.ok(shares.json.shares[0].shareId);
});

test('revoking a share stops the token from resolving', async (t) => {
  const server = await startServer({ backend: new FakeBackend() });
  t.after(server.close);

  const created = await createDoc(server.baseUrl, USER_A);
  const id = created.json.document.id;
  const share = await call(server.baseUrl, 'POST', `/api/diagrams/${id}/shares`, {
    headers: USER_A, body: { role: 'viewer' },
  });
  const token = share.json.token;
  const shareId = share.json.share.shareId;

  const before = await call(server.baseUrl, 'GET', `/api/diagrams/shared/${token}`, { headers: USER_B });
  assert.equal(before.status, 200);

  const revoke = await call(server.baseUrl, 'DELETE', `/api/diagrams/${id}/shares/${shareId}`, { headers: USER_A });
  assert.equal(revoke.status, 204);
  const current = await call(server.baseUrl, 'GET', `/api/diagrams/${id}`, { headers: USER_A });
  assert.equal(current.json.document.revision, 3);

  const after = await call(server.baseUrl, 'GET', `/api/diagrams/shared/${token}`, { headers: USER_B });
  assert.equal(after.status, 404);
});

test('deleting a document requires If-Match and cleans up shares and versions', async (t) => {
  const backend = new FakeBackend();
  const server = await startServer({ backend });
  t.after(server.close);

  const created = await createDoc(server.baseUrl, USER_A);
  const id = created.json.document.id;

  await call(server.baseUrl, 'POST', `/api/diagrams/${id}/versions`, { headers: USER_A, body: {} });
  const share = await call(server.baseUrl, 'POST', `/api/diagrams/${id}/shares`, {
    headers: USER_A, body: { role: 'viewer' },
  });
  const token = share.json.token;

  const noMatch = await call(server.baseUrl, 'DELETE', `/api/diagrams/${id}`, { headers: USER_A });
  assert.equal(noMatch.status, 428);

  // Re-read current etag (share creation bumped it).
  const current = await call(server.baseUrl, 'GET', `/api/diagrams/${id}`, { headers: USER_A });
  const del = await call(server.baseUrl, 'DELETE', `/api/diagrams/${id}`, {
    headers: USER_A, ifMatch: current.etag,
  });
  assert.equal(del.status, 204);

  // Every blob for the owner/document prefix and the share index is gone.
  const remaining = Array.from(backend.map.keys());
  assert.equal(remaining.length, 0, `expected no residual blobs, found: ${remaining.join(', ')}`);

  const resolve = await call(server.baseUrl, 'GET', `/api/diagrams/shared/${token}`, { headers: USER_B });
  assert.equal(resolve.status, 404);
});

test('delete waits for lock completion before sending success', async (t) => {
  const backend = new FakeBackend({
    afterLockOperation() {
      const error = new Error('Lock ownership could not be confirmed');
      error.statusCode = 503;
      throw error;
    },
  });
  const server = await startServer({ backend });
  t.after(server.close);

  const created = await createDoc(server.baseUrl, USER_A);
  const res = await call(server.baseUrl, 'DELETE', `/api/diagrams/${created.json.document.id}`, {
    headers: USER_A,
    ifMatch: created.etag,
  });
  assert.equal(res.status, 503);
  assert.deepEqual(res.json, { error: 'Diagram storage is temporarily unavailable' });
});

test('delete cannot discard an update that wins after the initial read', async (t) => {
  const backend = new ConcurrentUpdateBeforeDeleteBackend();
  const server = await startServer({ backend });
  t.after(server.close);

  const created = await createDoc(server.baseUrl, USER_A);
  const id = created.json.document.id;
  const deleted = await call(server.baseUrl, 'DELETE', `/api/diagrams/${id}`, {
    headers: USER_A,
    ifMatch: created.etag,
  });
  assert.equal(deleted.status, 412);

  const current = await call(server.baseUrl, 'GET', `/api/diagrams/${id}`, { headers: USER_A });
  assert.equal(current.status, 200);
  assert.equal(current.json.document.diagramName, 'Concurrent update');
  assert.equal(current.json.document.revision, 2);
});

test('failed delete cleanup is inaccessible and can be retried', async (t) => {
  const backend = new FailVersionCleanupOnceBackend();
  const server = await startServer({ backend });
  t.after(server.close);

  const created = await createDoc(server.baseUrl, USER_A);
  const id = created.json.document.id;
  const snapshot = await call(server.baseUrl, 'POST', `/api/diagrams/${id}/versions`, {
    headers: USER_A,
    body: { notes: 'cleanup retry' },
  });
  const versionId = snapshot.json.version.versionId;

  const firstDelete = await call(server.baseUrl, 'DELETE', `/api/diagrams/${id}`, {
    headers: USER_A,
    ifMatch: created.etag,
  });
  assert.equal(firstDelete.status, 503);

  const document = await call(server.baseUrl, 'GET', `/api/diagrams/${id}`, { headers: USER_A });
  assert.equal(document.status, 404);
  const version = await call(
    server.baseUrl,
    'GET',
    `/api/diagrams/${id}/versions/${versionId}`,
    { headers: USER_A },
  );
  assert.equal(version.status, 404);

  const retry = await call(server.baseUrl, 'DELETE', `/api/diagrams/${id}`, {
    headers: USER_A,
    ifMatch: created.etag,
  });
  assert.equal(retry.status, 204);
  assert.equal(backend.map.size, 0);
});

test('delete waits for an in-flight snapshot and removes it without leaving orphaned blobs', async (t) => {
  const backend = new BlockingVersionWriteBackend();
  const server = await startServer({ backend });
  t.after(server.close);

  const created = await createDoc(server.baseUrl, USER_A);
  const id = created.json.document.id;
  const snapshotPromise = call(server.baseUrl, 'POST', `/api/diagrams/${id}/versions`, {
    headers: USER_A,
    body: { notes: 'concurrent snapshot' },
  });
  await backend.versionWriteStarted;

  const deletePromise = call(server.baseUrl, 'DELETE', `/api/diagrams/${id}`, {
    headers: USER_A,
    ifMatch: created.etag,
  });
  let deleteSettled = false;
  void deletePromise.finally(() => {
    deleteSettled = true;
  });
  await new Promise((resolve) => setTimeout(resolve, 50));
  assert.equal(deleteSettled, false, 'delete must wait for the snapshot lock');
  backend.releaseVersionWrite();

  const [snapshot, deleted] = await Promise.all([snapshotPromise, deletePromise]);
  assert.equal(snapshot.status, 201);
  assert.equal(deleted.status, 204);
  assert.equal(backend.map.size, 0);
});

test('Azure Blob backend retries container creation after a transient failure', async () => {
  let createAttempts = 0;
  let uploads = 0;
  const containerClient = {
    createIfNotExists() {
      createAttempts += 1;
      if (createAttempts === 1) return Promise.reject(new Error('Transient initialization failure'));
      return Promise.resolve({});
    },
    getBlockBlobClient() {
      return {
        async upload() {
          uploads += 1;
          return { etag: '"etag-upload"' };
        },
      };
    },
  };
  const backend = createAzureBlobBackend({ containerClient });

  await assert.rejects(
    backend.put('owners/example/current.json', { ok: false }),
    /Transient initialization failure/,
  );
  const result = await backend.put('owners/example/current.json', { ok: true });
  assert.equal(result.etag, '"etag-upload"');
  assert.equal(createAttempts, 2);
  assert.equal(uploads, 1);
});

test('Azure Blob backend does not turn a successful operation into a retry after release failure', async () => {
  const warnings = [];
  const leaseClient = {
    async acquireLease() {},
    async renewLease() {},
    async releaseLease() {
      throw new Error('Release failed');
    },
  };
  const containerClient = {
    async createIfNotExists() {},
    getBlockBlobClient() {
      return {
        async upload() {
          return { etag: '"etag-upload"' };
        },
        getBlobLeaseClient() {
          return leaseClient;
        },
      };
    },
  };
  const backend = createAzureBlobBackend({
    containerClient,
    logger: {
      error() {},
      warn(...args) {
        warnings.push(args.join(' '));
      },
    },
  });

  const result = await backend.withLock('locks/document.lock', async () => 'saved');
  assert.equal(result, 'saved');
  assert.equal(warnings.length, 1);
  assert.match(warnings[0], /release failed after a successful operation/i);
});

test('Azure Blob backend returns a transient failure when renewal and release both fail', async () => {
  let signalRenewalAttempt;
  const renewalAttempted = new Promise((resolve) => {
    signalRenewalAttempt = resolve;
  });
  const leaseClient = {
    async acquireLease() {},
    async renewLease() {
      signalRenewalAttempt();
      throw new Error('Renewal failed');
    },
    async releaseLease() {
      throw new Error('Release failed');
    },
  };
  const containerClient = {
    async createIfNotExists() {},
    getBlockBlobClient() {
      return {
        async upload() {
          return { etag: '"etag-upload"' };
        },
        getBlobLeaseClient() {
          return leaseClient;
        },
      };
    },
  };
  const backend = createAzureBlobBackend({
    containerClient,
    leaseRenewalIntervalMs: 1,
    logger: { error() {}, warn() {} },
  });

  await assert.rejects(
    backend.withLock('locks/document.lock', async () => {
      await renewalAttempted;
      return 'saved';
    }),
    (error) => error.statusCode === 503 && /ownership could not be confirmed/i.test(error.message),
  );
});

test('malformed share tokens are rejected as not found', async (t) => {
  const server = await startServer({ backend: new FakeBackend() });
  t.after(server.close);
  const res = await call(server.baseUrl, 'GET', '/api/diagrams/shared/not-a-valid-token', { headers: USER_A });
  assert.equal(res.status, 404);
});
