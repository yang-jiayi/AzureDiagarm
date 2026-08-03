// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

const assert = require('node:assert/strict');
const { test } = require('node:test');
const express = require('express');
const {
  createAccessControlRouter,
  getAccessControlConfiguration,
} = require('./access-control');

class FakeTable {
  constructor() {
    this.entities = new Map();
  }

  listEntities() {
    const values = Array.from(this.entities.values());
    return {
      async *[Symbol.asyncIterator]() {
        for (const value of values) yield value;
      },
    };
  }

  async createEntity(entity) {
    const key = `${entity.partitionKey}:${entity.rowKey}`;
    if (this.entities.has(key)) {
      const error = new Error('Entity exists');
      error.statusCode = 409;
      throw error;
    }
    this.entities.set(key, entity);
  }

  async deleteEntity(partitionKey, rowKey) {
    const key = `${partitionKey}:${rowKey}`;
    if (!this.entities.delete(key)) {
      const error = new Error('Entity not found');
      error.statusCode = 404;
      throw error;
    }
  }
}

async function startServer(options) {
  const app = express();
  app.use(express.json());
  app.use('/api/access', createAccessControlRouter(options));
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

const ADMIN_HEADERS = {
  'X-MS-CLIENT-PRINCIPAL-NAME': 'yangjiayi@msft.jp',
  'X-MS-CLIENT-PRINCIPAL-ID': '11111111-1111-1111-1111-111111111111',
};
const MEMBER_HEADERS = {
  'X-MS-CLIENT-PRINCIPAL-NAME': 'member@example.com',
  'X-MS-CLIENT-PRINCIPAL-ID': '22222222-2222-2222-2222-222222222222',
};
const APP_ORIGIN = 'https://azurediagarm.mssql.biz';

test('access-control configuration reports mandatory readiness dependencies', () => {
  assert.deepEqual(
    getAccessControlConfiguration({ enabled: false }).missing,
    [],
  );

  const incomplete = getAccessControlConfiguration({
    enabled: true,
    adminEmail: 'not-an-email',
    publicAppUrl: 'ftp://not-an-http-origin.example',
  });
  assert.equal(incomplete.configured, false);
  assert.deepEqual(incomplete.missing, ['administrator', 'public URL', 'access store']);

  const complete = getAccessControlConfiguration({
    enabled: true,
    adminEmail: 'yangjiayi@msft.jp',
    publicAppUrl: APP_ORIGIN,
    table: new FakeTable(),
  });
  assert.equal(complete.configured, true);
  assert.deepEqual(complete.missing, []);
});

test('access control enforces whitelist and administrator-only management', async (t) => {
  const table = new FakeTable();
  const server = await startServer({
    enabled: true,
    adminEmail: 'yangjiayi@msft.jp',
    publicAppUrl: APP_ORIGIN,
    table,
    cacheTtlMs: 1_000,
    logger: { info() {}, error() {} },
  });
  t.after(server.close);

  let response = await fetch(`${server.baseUrl}/api/access/check`);
  assert.equal(response.status, 401);

  response = await fetch(`${server.baseUrl}/api/access/check`, {
    headers: { 'X-MS-CLIENT-PRINCIPAL-NAME': 'yangjiayi@msft.jp' },
  });
  assert.equal(response.status, 401);

  response = await fetch(`${server.baseUrl}/api/access/check`, { headers: MEMBER_HEADERS });
  assert.equal(response.status, 403);

  response = await fetch(`${server.baseUrl}/api/access/check`, { headers: ADMIN_HEADERS });
  assert.equal(response.status, 204);

  response = await fetch(`${server.baseUrl}/api/access/me`, { headers: ADMIN_HEADERS });
  assert.equal(response.status, 200);
  assert.deepEqual(await response.json(), {
    enabled: true,
    authenticated: true,
    email: 'yangjiayi@msft.jp',
    isAdmin: true,
    allowed: true,
  });

  response = await fetch(`${server.baseUrl}/api/access/users`, { headers: MEMBER_HEADERS });
  assert.equal(response.status, 403);

  response = await fetch(`${server.baseUrl}/api/access/users`, {
    method: 'POST',
    headers: {
      ...ADMIN_HEADERS,
      Origin: 'https://attacker.example',
      'Content-Type': 'application/json',
    },
    body: JSON.stringify({ email: 'member@example.com' }),
  });
  assert.equal(response.status, 403);

  response = await fetch(`${server.baseUrl}/api/access/users`, {
    method: 'POST',
    headers: {
      ...ADMIN_HEADERS,
      Origin: APP_ORIGIN,
      'Content-Type': 'application/json',
    },
    body: JSON.stringify({ email: 'Member@Example.com' }),
  });
  assert.equal(response.status, 201);

  response = await fetch(`${server.baseUrl}/api/access/check`, { headers: MEMBER_HEADERS });
  assert.equal(response.status, 204);

  response = await fetch(`${server.baseUrl}/api/access/users`, { headers: ADMIN_HEADERS });
  assert.equal(response.status, 200);
  const list = await response.json();
  assert.deepEqual(list.users.map((user) => user.email), [
    'yangjiayi@msft.jp',
    'member@example.com',
  ]);
  assert.equal(list.users[0].immutable, true);

  response = await fetch(`${server.baseUrl}/api/access/users`, {
    method: 'DELETE',
    headers: {
      ...ADMIN_HEADERS,
      Origin: APP_ORIGIN,
      'Content-Type': 'application/json',
    },
    body: JSON.stringify({ email: 'member@example.com' }),
  });
  assert.equal(response.status, 204);

  response = await fetch(`${server.baseUrl}/api/access/check`, { headers: MEMBER_HEADERS });
  assert.equal(response.status, 403);

  response = await fetch(`${server.baseUrl}/api/access/users`, {
    method: 'DELETE',
    headers: {
      ...ADMIN_HEADERS,
      Origin: APP_ORIGIN,
      'Content-Type': 'application/json',
    },
    body: JSON.stringify({ email: 'yangjiayi@msft.jp' }),
  });
  assert.equal(response.status, 400);
});

test('disabled access control remains transparent', async (t) => {
  const server = await startServer({ enabled: false });
  t.after(server.close);

  let response = await fetch(`${server.baseUrl}/api/access/check`);
  assert.equal(response.status, 204);

  response = await fetch(`${server.baseUrl}/api/access/me`);
  assert.equal(response.status, 200);
  assert.deepEqual(await response.json(), {
    enabled: false,
    authenticated: false,
    email: null,
    isAdmin: false,
    allowed: true,
  });
});

test('access control accepts the aggregate Easy Auth client principal header', async (t) => {
  const table = new FakeTable();
  const server = await startServer({
    enabled: true,
    adminEmail: 'yangjiayi@msft.jp',
    publicAppUrl: APP_ORIGIN,
    table,
    logger: { info() {}, error() {} },
  });
  t.after(server.close);

  const encodedPrincipal = Buffer.from(JSON.stringify({
    auth_typ: 'aad',
    claims: [
      { typ: 'preferred_username', val: 'yangjiayi@msft.jp' },
      {
        typ: 'http://schemas.microsoft.com/identity/claims/objectidentifier',
        val: '33333333-3333-3333-3333-333333333333',
      },
    ],
  })).toString('base64');

  let response = await fetch(`${server.baseUrl}/api/access/check`, {
    headers: { 'X-MS-CLIENT-PRINCIPAL': encodedPrincipal },
  });
  assert.equal(response.status, 204);

  response = await fetch(`${server.baseUrl}/api/access/check`, {
    headers: { 'X-MS-CLIENT-PRINCIPAL': 'not-base64-json' },
  });
  assert.equal(response.status, 401);
});

test('access control maps Entra B2B guest UPNs back to their invited email', async (t) => {
  const table = new FakeTable();
  const server = await startServer({
    enabled: true,
    adminEmail: 'yangjiayi@msft.jp',
    publicAppUrl: APP_ORIGIN,
    table,
    logger: { info() {}, error() {} },
  });
  t.after(server.close);

  let response = await fetch(`${server.baseUrl}/api/access/users`, {
    method: 'POST',
    headers: {
      ...ADMIN_HEADERS,
      Origin: APP_ORIGIN,
      'Content-Type': 'application/json',
    },
    body: JSON.stringify({ email: 'maaya_ishida@microsoft.com' }),
  });
  assert.equal(response.status, 201);

  response = await fetch(`${server.baseUrl}/api/access/check`, {
    headers: {
      'X-MS-CLIENT-PRINCIPAL-NAME':
        'maaya_ishida_microsoft.com#EXT#@exampletenant.onmicrosoft.com',
      'X-MS-CLIENT-PRINCIPAL-ID': '44444444-4444-4444-4444-444444444444',
    },
  });
  assert.equal(response.status, 204);

  const encodedPrincipal = Buffer.from(JSON.stringify({
    auth_typ: 'aad',
    claims: [
      {
        typ: 'preferred_username',
        val: 'maaya_ishida_microsoft.com#EXT#@exampletenant.onmicrosoft.com',
      },
      { typ: 'email', val: 'maaya_ishida@microsoft.com' },
      {
        typ: 'http://schemas.microsoft.com/identity/claims/objectidentifier',
        val: '44444444-4444-4444-4444-444444444444',
      },
    ],
  })).toString('base64');

  response = await fetch(`${server.baseUrl}/api/access/me`, {
    headers: { 'X-MS-CLIENT-PRINCIPAL': encodedPrincipal },
  });
  assert.equal(response.status, 200);
  assert.deepEqual(await response.json(), {
    enabled: true,
    authenticated: true,
    email: 'maaya_ishida@microsoft.com',
    isAdmin: false,
    allowed: true,
  });
});
