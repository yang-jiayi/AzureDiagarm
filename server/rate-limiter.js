// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

const crypto = require('crypto');

// ── Client-key derivation ─────────────────────────────────────────────────────
//
// nginx forwards X-Forwarded-For as `$proxy_add_x_forwarded_for`, so the
// *first* entry is whatever the caller sent — using it would let anyone reset
// their own quota (and spoof another caller's) by varying the header.  Azure
// Front Door owns X-Azure-ClientIP / X-Azure-SocketIP and overwrites any
// client-supplied value, and nginx rejects requests that do not carry this
// deployment's X-Azure-FDID, so those headers are trustworthy here.  When they
// are absent the *last* X-Forwarded-For entry is used, because that one is
// appended by the nearest trusted proxy rather than by the caller.

const IP_LIKE = /^[0-9a-fA-F.:[\]]{3,45}$/;

function trustedIp(value) {
  const candidate = String(value || '').trim();
  return IP_LIKE.test(candidate) ? candidate : '';
}

function getClientKey(req) {
  const forwarded = String(req.get('x-forwarded-for') || '').split(',');
  return (
    trustedIp(req.get('x-azure-clientip'))
    || trustedIp(req.get('x-azure-socketip'))
    || trustedIp(forwarded[forwarded.length - 1])
    || trustedIp(req.ip)
    || 'unknown'
  ).slice(0, 128);
}

// ── In-process fixed-window limiter ──────────────────────────────────────────
//
// Safe for single-replica deployments.  With multiple Container Apps replicas
// each process has its own independent counter, so the effective limit is
// maxRequests × <replica count>.  Use createTableRateLimiter for an honest
// global limit across replicas.
//
// Returns a synchronous function (req) => retryAfterSeconds.

function createFixedWindowRateLimiter(windowMs, maxRequests) {
  const clients = new Map();
  return (req) => {
    const now = Date.now();
    for (const [key, value] of clients) {
      if (value.resetAt <= now) clients.delete(key);
    }

    const key = getClientKey(req);
    const current = clients.get(key);
    if (!current || current.resetAt <= now) {
      clients.set(key, { count: 1, resetAt: now + windowMs });
      return 0;
    }
    if (current.count >= maxRequests) {
      return Math.max(1, Math.ceil((current.resetAt - now) / 1000));
    }
    current.count += 1;
    return 0;
  };
}

// ── Shared Azure Table Storage fixed-window limiter ───────────────────────────
//
// Enforces a global per-client limit across any number of Container Apps
// replicas by storing counters in Azure Table Storage.
//
// Algorithm: optimistic concurrency via ETag precondition.
//   1. Read entity (partitionKey=RATE_LIMIT_PARTITION, rowKey=sha256(clientKey)).
//   2. If already over limit → return retryAfter immediately (no write needed).
//   3. Increment count and write back with If-Match: <etag>.
//   4. On 409/412 conflict (another replica beat us) → re-read and retry.
//   5. On any non-conflict storage error, or after MAX_RETRIES conflicts →
//      fail closed (return non-zero retryAfter) so the request is rejected
//      rather than recorded without a count.
//
// Latency cost: one GetEntity + one CreateEntity/UpdateEntity per non-limited
// request (~1–5 ms round-trip within the same Azure region).
//
// tableClient  – @azure/data-tables TableClient pointed at the rate-limit table.
//                The table must exist; call tableClient.createTable() at startup
//                (ignoring 409 if it already exists).
// windowMs     – window length in milliseconds.
// maxRequests  – maximum allowed requests per client per window.
//
// Returns an async function (req) => Promise<retryAfterSeconds>.

const RATE_LIMIT_PARTITION = 'window';
const MAX_RETRIES = 4;

function createTableRateLimiter(tableClient, windowMs, maxRequests, options = {}) {
  const configuredRetryAfter = Number(options.storageErrorRetryAfterSeconds);
  const storageErrorRetryAfter = Number.isFinite(configuredRetryAfter)
    ? Math.max(1, Math.floor(configuredRetryAfter))
    : 5;
  const onStorageError = typeof options.onStorageError === 'function'
    ? options.onStorageError
    : () => {};

  function rowKeyForClient(clientKey) {
    return crypto.createHash('sha256').update(clientKey).digest('hex');
  }

  return async (req) => {
    const clientKey = getClientKey(req);
    const rowKey = rowKeyForClient(clientKey);
    // Returned when we cannot safely record the request.
    const errorRetryAfter = Math.min(
      Math.ceil(windowMs / 1000),
      storageErrorRetryAfter,
    );

    for (let attempt = 0; attempt < MAX_RETRIES; attempt++) {
      const now = Date.now();
      let entity = null;
      let etag = null;

      try {
        entity = await tableClient.getEntity(RATE_LIMIT_PARTITION, rowKey);
        // @azure/data-tables v13 maps the ETag response header onto entity.etag
        etag = entity.etag;
      } catch (err) {
        if (err.statusCode !== 404) {
          // Transient read failure — fail closed: reject the request.
          onStorageError('read', err);
          return errorRetryAfter;
        }
        // 404 → first request from this client; entity will be created below.
      }

      const resetAt = entity ? Number(entity.resetAt) : 0;
      const count = entity ? Number(entity.count) : 0;
      const windowExpired = resetAt <= now;

      // If the current window is active and the counter is already at the
      // ceiling, reject without touching storage.
      if (!windowExpired && count >= maxRequests) {
        return Math.max(1, Math.ceil((resetAt - now) / 1000));
      }

      const newCount = windowExpired ? 1 : count + 1;
      const newResetAt = windowExpired ? now + windowMs : resetAt;
      const newEntity = {
        partitionKey: RATE_LIMIT_PARTITION,
        rowKey,
        count: newCount,
        resetAt: String(newResetAt),
      };

      try {
        if (entity === null) {
          // First ever request from this client — create a fresh entity.
          await tableClient.createEntity(newEntity);
        } else {
          // Update with ETag precondition to detect concurrent writes.
          await tableClient.updateEntity(newEntity, 'Replace', { etag });
        }
        return 0; // Request recorded; not rate-limited.
      } catch (err) {
        if (err.statusCode === 412 || err.statusCode === 409) {
          // Another replica wrote the entity between our read and write;
          // re-read and retry from the top of the loop.
          continue;
        }
        // Any other write failure — fail closed.
        onStorageError('write', err);
        return errorRetryAfter;
      }
    }

    // All retry attempts exhausted due to extreme concurrent pressure from the
    // same client IP.  Fail closed: reject rather than silently skip recording.
    onStorageError(
      'conflict-retries',
      new Error(`Rate-limit counter update conflicted ${MAX_RETRIES} times`),
    );
    return errorRetryAfter;
  };
}

module.exports = {
  createFixedWindowRateLimiter,
  createTableRateLimiter,
  getClientKey,
};
