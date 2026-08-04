// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

const KEY_VAULT_RESOURCE_ID_PATTERN =
  /^\/subscriptions\/[0-9a-f-]+\/resourceGroups\/[^/]+\/providers\/Microsoft\.KeyVault\/vaults\/[A-Za-z0-9-]+$/i;
const API_VERSION = '2023-07-01';
const ARM_SCOPE = 'https://management.azure.com/.default';
const ACCESS_ENTRY_CONTENT_TYPE = 'application/vnd.azurediagarm.access';
const ACCESS_ENTRY_KIND = 'azurediagarm-access-v1';
const ACCESS_SECRET_NAME_PATTERN = /^[a-f0-9]{64}$/;

class ArmKeyVaultAccessStore {
  constructor(resourceId, credential, fetchImpl = fetch) {
    const normalizedResourceId = String(resourceId || '').trim().replace(/\/+$/, '');
    if (!KEY_VAULT_RESOURCE_ID_PATTERN.test(normalizedResourceId)) {
      throw new Error('AZURE_ACCESS_KEY_VAULT_RESOURCE_ID is invalid.');
    }
    this.resourceId = normalizedResourceId;
    this.credential = credential;
    this.fetch = fetchImpl;
    this.baseUrl = `https://management.azure.com${normalizedResourceId}/secrets`;
    this.allowedPathPrefix = new URL(this.baseUrl).pathname.toLowerCase();
  }

  assertAllowedUrl(value) {
    const url = new URL(value);
    if (
      url.protocol !== 'https:'
      || url.hostname !== 'management.azure.com'
      || (
        url.pathname.toLowerCase() !== this.allowedPathPrefix
        && !url.pathname.toLowerCase().startsWith(`${this.allowedPathPrefix}/`)
      )
    ) {
      throw new Error('Azure Resource Manager returned an invalid Key Vault continuation URL.');
    }
    return url.toString();
  }

  isAccessEntry(secret) {
    return (
      ACCESS_SECRET_NAME_PATTERN.test(String(secret?.name || ''))
      && secret?.properties?.contentType === ACCESS_ENTRY_CONTENT_TYPE
      && secret?.tags?.kind === ACCESS_ENTRY_KIND
    );
  }

  async request(url, init = {}) {
    const { token } = await this.credential.getToken(ARM_SCOPE);
    const response = await this.fetch(this.assertAllowedUrl(url), {
      ...init,
      headers: {
        Accept: 'application/json',
        Authorization: `Bearer ${token}`,
        ...(init.body ? { 'Content-Type': 'application/json' } : {}),
        ...(init.headers || {}),
      },
      signal: AbortSignal.timeout(15_000),
    });

    if (!response.ok) {
      const error = new Error(`Azure Resource Manager returned ${response.status}.`);
      error.statusCode = response.status;
      throw error;
    }
    if (response.status === 204) return null;
    return response.json();
  }

  listEntities() {
    const store = this;
    return {
      async *[Symbol.asyncIterator]() {
        let nextUrl = `${store.baseUrl}?api-version=${API_VERSION}`;
        while (nextUrl) {
          const page = await store.request(nextUrl);
          for (const secret of page.value || []) {
            if (!store.isAccessEntry(secret)) continue;
            if (secret?.properties?.attributes?.enabled === false) continue;
            const email = secret?.tags?.email;
            if (typeof email !== 'string' || email.length === 0) continue;
            yield {
              email,
              addedAt: typeof secret.tags.addedAt === 'string' ? secret.tags.addedAt : '',
              addedBy: typeof secret.tags.addedBy === 'string' ? secret.tags.addedBy : '',
            };
          }
          nextUrl = typeof page.nextLink === 'string'
            ? store.assertAllowedUrl(page.nextLink)
            : '';
        }
      },
    };
  }

  async createEntity(entity) {
    if (!ACCESS_SECRET_NAME_PATTERN.test(String(entity.rowKey || ''))) {
      throw new Error('The access entry key is invalid.');
    }
    const body = {
      tags: {
        kind: ACCESS_ENTRY_KIND,
        email: entity.email,
        addedAt: entity.addedAt,
        addedBy: entity.addedBy,
      },
      properties: {
        value: 'allowed',
        contentType: ACCESS_ENTRY_CONTENT_TYPE,
        attributes: { enabled: true },
      },
    };
    await this.request(
      `${this.baseUrl}/${encodeURIComponent(entity.rowKey)}?api-version=${API_VERSION}`,
      { method: 'PUT', body: JSON.stringify(body) },
    );
  }

  async deleteEntity(_partitionKey, rowKey) {
    if (!ACCESS_SECRET_NAME_PATTERN.test(String(rowKey || ''))) {
      const error = new Error('The access entry key is invalid.');
      error.statusCode = 404;
      throw error;
    }
    const url = `${this.baseUrl}/${encodeURIComponent(rowKey)}?api-version=${API_VERSION}`;
    const current = await this.request(url);
    if (!this.isAccessEntry({ ...current, name: rowKey })) {
      const error = new Error('The access entry was not found.');
      error.statusCode = 404;
      throw error;
    }
    const body = {
      tags: current.tags || {},
      properties: {
        value: 'revoked',
        contentType: ACCESS_ENTRY_CONTENT_TYPE,
        attributes: { enabled: false },
      },
    };
    await this.request(url, { method: 'PUT', body: JSON.stringify(body) });
  }
}

module.exports = {
  ArmKeyVaultAccessStore,
};
