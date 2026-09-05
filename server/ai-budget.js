// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.
const crypto = require('crypto');
const { getPrincipal } = require('./access-control');

const DOCUMENT_TTL = 3 * 24 * 60 * 60;
const LEASE_MS = 315_000; // Beyond the proxy's 210s deadline, including settlement.
const statusCode = error => Number(error?.statusCode || error?.status || error?.code);
const userKey = id => `ai-budget-${crypto.createHash('sha256').update(id).digest('hex')}`;

class BudgetError extends Error {
  constructor(code, message, retryAfter = 5, status = 429) {
    super(message);
    Object.assign(this, { code, retryAfter, status });
  }
}

class MemoryBudgetStore {
  constructor() { this.documents = new Map(); }
  async read(id) {
    const item = this.documents.get(id);
    return item ? structuredClone(item) : null;
  }
  async write(doc, etag) {
    const current = this.documents.get(doc.id);
    if (current?._etag !== etag) throw Object.assign(new Error('Conflict'), { code: 412 });
    this.documents.set(doc.id, structuredClone({ ...doc, _etag: crypto.randomUUID() }));
    for (const [id, value] of this.documents) {
      if (value.expiresAt <= Date.now()) this.documents.delete(id);
    }
  }
}

class CosmosBudgetStore {
  constructor(container, client) { this.container = container; this.client = client; }
  async read(id) {
    try { return (await this.container.item(id, id).read()).resource || null; }
    catch (error) {
      if (statusCode(error) !== 404) throw error;
      // A missing container must not masquerade as a new user's full balance.
      await this.container.read();
      return null;
    }
  }
  async write(doc, etag) {
    const { _etag, ...body } = doc;
    if (!etag) return this.container.items.create(body);
    return this.container.item(doc.id, doc.id).replace(body, {
      accessCondition: { type: 'IfMatch', condition: etag },
    });
  }
  async validate() {
    if (this.client) {
      const { resource: account } = await this.client.getDatabaseAccount();
      if (!account || account.enableMultipleWritableLocations || account.writableLocations.length !== 1) {
        throw new Error('Atomic AI budgets require a Cosmos account with exactly one write region.');
      }
    }
    const { resource } = await this.container.read();
    if (resource?.partitionKey?.paths?.join() !== '/id' || !resource.defaultTtl) {
      throw new Error('Budget Cosmos container requires partition /id and enabled TTL (defaultTtl=-1).');
    }
  }
}

class TableBudgetStore {
  constructor(table) { this.table = table; }
  async read(id) {
    try {
      const entity = await this.table.getEntity('ai-budget', id);
      return { ...JSON.parse(entity.document), _etag: entity.etag };
    } catch (error) {
      if (statusCode(error) !== 404 || error?.code === 'TableNotFound'
        || error?.details?.odataError?.code === 'TableNotFound') throw error;
      // Confirm the table still exists: a deleted/unavailable table is not a
      // newly authenticated user with an unused balance.
      await this.table.listEntities({
        queryOptions: { filter: "PartitionKey eq 'ai-budget'", select: ['RowKey'] },
      }).byPage({ maxPageSize: 1 }).next();
      return null;
    }
  }
  async write(doc, etag) {
    const { _etag, ...body } = doc;
    const entity = { partitionKey: 'ai-budget', rowKey: doc.id, document: JSON.stringify(body), expiresAt: new Date(doc.expiresAt).toISOString() };
    if (!etag) return this.table.createEntity(entity);
    return this.table.updateEntity(entity, 'Replace', { etag });
  }
  async validate() {
    try { await this.table.createTable(); }
    catch (error) { if (statusCode(error) !== 409) throw error; }
  }
  async sweep(now = Date.now()) {
    for await (const entity of this.table.listEntities({
      queryOptions: { filter: `PartitionKey eq 'ai-budget' and expiresAt lt '${new Date(now).toISOString()}'` },
    })) {
      try { await this.table.deleteEntity('ai-budget', entity.rowKey, { etag: entity.etag }); }
      catch (error) { if (![404, 412].includes(statusCode(error))) throw error; }
    }
  }
}

function createBudgetManager({ store, dailyTokens = 250_000, concurrency = 2, now = Date.now, leaseMs = LEASE_MS }) {
  const snapshot = (doc) => {
    const reservedTokens = Object.values(doc.leases).filter(lease => lease.day === doc.day)
      .reduce((sum, lease) => sum + lease.tokens, 0);
    return {
      available: true, limitTokens: dailyTokens, usedTokens: doc.usedTokens,
      reservedTokens, remainingTokens: Math.max(0, dailyTokens - doc.usedTokens),
      concurrentRequests: Object.keys(doc.leases).length, concurrentLimit: concurrency,
      resetAt: new Date(Date.parse(`${doc.day}T00:00:00Z`) + 86400_000).toISOString(),
    };
  };
  async function mutate(id, change, readOnly = false) {
    for (let attempt = 0; attempt < 12; attempt++) {
      const time = now();
      const day = new Date(time).toISOString().slice(0, 10);
      const previous = await store.read(userKey(id));
      const doc = previous || { id: userKey(id), type: 'ai-budget', day, usedTokens: 0, leases: {} };
      // Clock skew at midnight must not let an older replica reset the day backwards.
      if (doc.day < day) { doc.day = day; doc.usedTokens = 0; }
      // Expired/crashed requests keep their reservation charged: missing usage
      // must not grant a refund that can be exploited by cancelling requests.
      for (const [key, lease] of Object.entries(doc.leases)) {
        if (lease.expiresAt <= time) delete doc.leases[key];
      }
      const result = change(doc, time);
      if (readOnly) return result;
      doc.ttl = DOCUMENT_TTL;
      doc.expiresAt = time + DOCUMENT_TTL * 1000;
      try { await store.write(doc, previous?._etag); return result; }
      catch (error) {
        if (![409, 412].includes(statusCode(error))) throw error;
      }
    }
    throw new BudgetError('ai_budget_busy', 'Budget is busy. Retry in a few seconds.', 2, 503);
  }
  return {
    status: id => mutate(id, snapshot, true),
    reserve: (id, tokens) => mutate(id, (doc, time) => {
      if (!Number.isSafeInteger(tokens) || tokens < 1) throw new Error('Invalid token reservation.');
      if (Object.keys(doc.leases).length >= concurrency) {
        const wait = Math.max(1, Math.ceil((Math.min(...Object.values(doc.leases).map(lease => lease.expiresAt)) - time) / 1000));
        throw new BudgetError('ai_concurrency_limit', 'Too many AI requests are running. Wait for a request to finish, then retry.', wait);
      }
      if (doc.usedTokens + tokens > dailyTokens) {
        const wait = Math.max(1, Math.ceil((Date.parse(snapshot(doc).resetAt) - time) / 1000));
        throw new BudgetError('ai_daily_budget_exceeded', 'Daily token budget cannot cover this request. Reduce the input/output or wait until midnight UTC.', wait);
      }
      const lease = { id: crypto.randomUUID(), tokens, day: doc.day, expiresAt: time + leaseMs };
      doc.usedTokens += tokens;
      doc.leases[lease.id] = lease;
      return lease;
    }),
    settle: (id, lease, actualTokens) => mutate(id, doc => {
      if (!doc.leases[lease.id]) return;
      if (doc.day === lease.day && Number.isSafeInteger(actualTokens) && actualTokens >= 0) {
        doc.usedTokens = Math.max(0, doc.usedTokens + actualTokens - lease.tokens);
      }
      delete doc.leases[lease.id];
    }),
  };
}

function budgetIdentity(req, mode) {
  if (mode === 'local') return 'local-development';
  const principal = req.accessPrincipal || getPrincipal(req);
  if (!principal) throw new BudgetError('authentication_required', 'Sign in to use AI.', 0, 401);
  return principal.id;
}

function reservationTokens(body, format) {
  // UTF-8 bytes conservatively bound text tokens, with protocol overhead.
  // Vision is charged a full context-sized allowance per image; URLs must not
  // fetch unbounded remote input. Reconcile only from trusted upstream usage.
  let images = 0;
  const embeddedImage = value => typeof value === 'string' && /^data:image\/[a-z0-9.+-]+;base64,/i.test(value);
  const json = JSON.stringify(body, (_key, value) => {
    if (value && ['input_image', 'image_url'].includes(value.type)) {
      images++;
      if (embeddedImage(value.image_url)) return { ...value, image_url: '[embedded image]' };
      if (embeddedImage(value.image_url?.url)) {
        return { ...value, image_url: { ...value.image_url, url: '[embedded image]' } };
      }
    }
    if (value?.type === 'image' && value.source?.type === 'base64') {
      images++;
      return { ...value, source: { ...value.source, data: '[embedded image]' } };
    }
    return value;
  });
  const output = format === 'responses'
    ? body.max_output_tokens
    : (body.max_completion_tokens ?? body.max_tokens);
  return Buffer.byteLength(json, 'utf8') + 4096 + images * 131072 + output;
}

function hasUnmeteredInput(body) {
  const pending = [body.input, body.messages];
  while (pending.length) {
    const value = pending.pop();
    if (!value || typeof value !== 'object') continue;
    if (['item_reference', 'input_file', 'file', 'input_audio', 'audio', 'document'].includes(value.type)
      || value.file_id || value.file_url) return true;
    if (value.type === 'input_image' && !/^data:image\/[a-z0-9.+-]+;base64,/i.test(value.image_url || '')) return true;
    if (value.type === 'image_url' && !/^data:image\/[a-z0-9.+-]+;base64,/i.test(value.image_url?.url || '')) return true;
    if (value.type === 'image' && (value.source?.type !== 'base64'
      || typeof value.source.data !== 'string'
      || !/^image\/[a-z0-9.+-]+$/i.test(value.source.media_type || ''))) return true;
    for (const child of Object.values(value)) {
      if (child && typeof child === 'object') pending.push(child);
    }
  }
  return false;
}

function actualUsage(payload) {
  const usage = payload?.usage;
  const total = usage?.total_tokens ?? ((usage?.input_tokens ?? usage?.prompt_tokens)
    + (usage?.output_tokens ?? usage?.completion_tokens)
    + (usage?.cache_creation_input_tokens ?? 0) + (usage?.cache_read_input_tokens ?? 0));
  return Number.isSafeInteger(total) && total >= 0 ? total : undefined;
}

module.exports = {
  BudgetError, MemoryBudgetStore, CosmosBudgetStore, TableBudgetStore,
  createBudgetManager, budgetIdentity, reservationTokens, actualUsage, hasUnmeteredInput,
};
