// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

const assert = require('node:assert/strict');
const { test } = require('node:test');
const { ArmKeyVaultAccessStore } = require('./arm-key-vault-access-store');

const RESOURCE_ID =
  '/subscriptions/00000000-0000-0000-0000-000000000000/resourceGroups/example-rg/providers/Microsoft.KeyVault/vaults/example-access-kv';
const ROW_KEY = 'a'.repeat(64);

function jsonResponse(status, body) {
  return {
    ok: status >= 200 && status < 300,
    status,
    async json() {
      return body;
    },
  };
}

test('ARM Key Vault store lists only enabled access entries', async () => {
  const scopes = [];
  const credential = {
    async getToken(scope) {
      scopes.push(scope);
      return { token: 'test-token' };
    },
  };
  const store = new ArmKeyVaultAccessStore(RESOURCE_ID, credential, async () => jsonResponse(200, {
    value: [
      {
        name: ROW_KEY,
        tags: {
          kind: 'azurediagarm-access-v1',
          email: 'member@example.com',
          addedAt: '2026-01-01T00:00:00.000Z',
          addedBy: 'admin@example.com',
        },
        properties: {
          contentType: 'application/vnd.azurediagarm.access',
          attributes: { enabled: true },
        },
      },
      {
        name: 'b'.repeat(64),
        tags: { kind: 'azurediagarm-access-v1', email: 'disabled@example.com' },
        properties: {
          contentType: 'application/vnd.azurediagarm.access',
          attributes: { enabled: false },
        },
      },
      {
        name: 'c'.repeat(64),
        tags: { email: 'unrelated@example.com' },
        properties: {
          contentType: 'application/vnd.other',
          attributes: { enabled: true },
        },
      },
    ],
  }));

  const entities = [];
  for await (const entity of store.listEntities()) entities.push(entity);

  assert.deepEqual(entities, [{
    email: 'member@example.com',
    addedAt: '2026-01-01T00:00:00.000Z',
    addedBy: 'admin@example.com',
  }]);
  assert.deepEqual(scopes, ['https://management.azure.com/.default']);
});

test('ARM Key Vault store writes and revokes deterministic secret resources', async () => {
  const requests = [];
  const credential = {
    async getToken() {
      return { token: 'test-token' };
    },
  };
  const store = new ArmKeyVaultAccessStore(RESOURCE_ID, credential, async (url, init) => {
    requests.push({ url, init });
    if (!init?.method) {
      return jsonResponse(200, {
        tags: {
          kind: 'azurediagarm-access-v1',
          email: 'member@example.com',
          addedAt: '2026-01-01T00:00:00.000Z',
          addedBy: 'admin@example.com',
        },
        properties: {
          contentType: 'application/vnd.azurediagarm.access',
          attributes: { enabled: true },
        },
      });
    }
    return jsonResponse(200, {});
  });

  await store.createEntity({
    rowKey: ROW_KEY,
    email: 'member@example.com',
    addedAt: '2026-01-01T00:00:00.000Z',
    addedBy: 'admin@example.com',
  });
  await store.deleteEntity('allowed', ROW_KEY);

  assert.equal(requests[0].init.method, 'PUT');
  assert.equal(JSON.parse(requests[0].init.body).properties.attributes.enabled, true);
  assert.equal(requests[1].init.method, undefined);
  assert.equal(requests[2].init.method, 'PUT');
  assert.equal(JSON.parse(requests[2].init.body).properties.attributes.enabled, false);
  assert.match(requests[0].url, new RegExp(`/secrets/${ROW_KEY}\\?api-version=2023-07-01$`));
});

test('ARM Key Vault store rejects continuation links outside the configured vault', async () => {
  const credential = {
    async getToken() {
      return { token: 'test-token' };
    },
  };
  const store = new ArmKeyVaultAccessStore(RESOURCE_ID, credential, async () => jsonResponse(200, {
    value: [],
    nextLink: 'https://attacker.example/secrets?api-version=2023-07-01',
  }));

  await assert.rejects(async () => {
    for await (const _entity of store.listEntities()) {
      // No entries are expected.
    }
  }, /invalid Key Vault continuation URL/);
});

test('ARM Key Vault store rejects non-Key-Vault resource IDs', () => {
  assert.throws(
    () => new ArmKeyVaultAccessStore('https://attacker.example', { getToken() {} }),
    /invalid/,
  );
});
