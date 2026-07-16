// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

const assert = require('node:assert/strict');
const { test } = require('node:test');
const { ArmKeyVaultAccessStore } = require('./arm-key-vault-access-store');

const RESOURCE_ID =
  '/subscriptions/f2c0fe9a-0171-42ed-803d-3e78322545a1/resourceGroups/AzureDiagarm_rg/providers/Microsoft.KeyVault/vaults/azurediagarm-access-kv';

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
        name: 'active',
        tags: {
          email: 'member@example.com',
          addedAt: '2026-01-01T00:00:00.000Z',
          addedBy: 'yangjiayi@msft.jp',
        },
        properties: { attributes: { enabled: true } },
      },
      {
        name: 'disabled',
        tags: { email: 'disabled@example.com' },
        properties: { attributes: { enabled: false } },
      },
      {
        name: 'unrelated',
        properties: { attributes: { enabled: true } },
      },
    ],
  }));

  const entities = [];
  for await (const entity of store.listEntities()) entities.push(entity);

  assert.deepEqual(entities, [{
    email: 'member@example.com',
    addedAt: '2026-01-01T00:00:00.000Z',
    addedBy: 'yangjiayi@msft.jp',
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
          email: 'member@example.com',
          addedAt: '2026-01-01T00:00:00.000Z',
          addedBy: 'yangjiayi@msft.jp',
        },
        properties: { attributes: { enabled: true } },
      });
    }
    return jsonResponse(200, {});
  });

  await store.createEntity({
    rowKey: 'abc123',
    email: 'member@example.com',
    addedAt: '2026-01-01T00:00:00.000Z',
    addedBy: 'yangjiayi@msft.jp',
  });
  await store.deleteEntity('allowed', 'abc123');

  assert.equal(requests[0].init.method, 'PUT');
  assert.equal(JSON.parse(requests[0].init.body).properties.attributes.enabled, true);
  assert.equal(requests[1].init.method, undefined);
  assert.equal(requests[2].init.method, 'PUT');
  assert.equal(JSON.parse(requests[2].init.body).properties.attributes.enabled, false);
  assert.match(requests[0].url, /\/secrets\/abc123\?api-version=2023-07-01$/);
});

test('ARM Key Vault store rejects non-Key-Vault resource IDs', () => {
  assert.throws(
    () => new ArmKeyVaultAccessStore('https://attacker.example', { getToken() {} }),
    /invalid/,
  );
});
