// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.
const crypto = require('node:crypto');
const express = require('express');
const { budgetIdentity } = require('./ai-budget');
const { deriveOwnerKey } = require('./diagram-api');
const { awaitWithSignal } = require('./ai-http');

const JOB_TIMEOUT_MS = 15 * 60_000;
const RESULT_TTL_MS = 60 * 60_000;
const TOMBSTONE_TTL_MS = 24 * 60 * 60_000;
const JOB_ID_RE = /^[a-f0-9]{8}-[a-f0-9]{4}-4[a-f0-9]{3}-[89ab][a-f0-9]{3}-[a-f0-9]{12}$/;
const ACTIVE = new Set(['queued', 'running', 'cancelling']);
const PREFIX = 'ai-jobs/';

function jobError(code, message, status = 503) {
  return Object.assign(new Error(message), { code, status });
}

function failure(id, code, message, status = 503) {
  return {
    status, headers: { 'X-AzureDiagarm-Request-Id': id },
    body: { error: { source: 'job', code, message, requestId: id } },
  };
}

function createMemoryJobBackend() {
  const values = new Map();
  return {
    async read(name) { return values.has(name) ? structuredClone(values.get(name)) : null; },
    async create(name, value) {
      if (values.has(name)) throw Object.assign(new Error('Conflict'), { statusCode: 409 });
      values.set(name, { value: structuredClone(value), etag: crypto.randomUUID() });
    },
    async replace(name, value, etag) {
      if (values.get(name)?.etag !== etag) throw Object.assign(new Error('Conflict'), { statusCode: 412 });
      values.set(name, { value: structuredClone(value), etag: crypto.randomUUID() });
    },
    async remove(name, etag) {
      if (etag && values.get(name)?.etag !== etag) throw Object.assign(new Error('Conflict'), { statusCode: 412 });
      return values.delete(name);
    },
    async *list(prefix) { for (const name of values.keys()) if (name.startsWith(prefix)) yield name; },
  };
}

function createAIJobs({
  backend, mode = 'local', logger = console, now = Date.now,
  timeoutMs = JOB_TIMEOUT_MS, resultTtlMs = RESULT_TTL_MS,
  heartbeatMs = 5000, workerLeaseMs = 60_000, maxActive = 16,
  storageTimeoutMs = 8000,
  consumeControlRateLimit = () => 0,
} = {}) {
  const router = express.Router();
  const active = new Map();
  const worker = crypto.randomUUID();
  let executor;
  let closed = false;
  let admitting = 0;
  let sweeping = false;
  let sweepIterator;
  const io = operation => {
    const signal = AbortSignal.timeout(storageTimeoutMs);
    return awaitWithSignal(Promise.resolve().then(() => operation(signal)), signal);
  };
  if (backend) {
    const storage = backend;
    backend = {
      read: (...args) => io(() => storage.read(...args)),
      create: (...args) => io(() => storage.create(...args)),
      replace: (...args) => io(() => storage.replace(...args)),
      remove: (...args) => io(() => storage.remove(...args)),
      list: prefix => storage.list(prefix),
      listJobBlobs: storage.listJobBlobs ? () => storage.listJobBlobs() : undefined,
      removeJobVersion: storage.removeJobVersion
        ? (name, version) => io(signal => storage.removeJobVersion(name, version, { abortSignal: signal })) : undefined,
    };
  }
  const log = (event, id, details = {}, level = 'error') =>
    (logger[level] || logger.error).call(logger, `[ai-jobs] ${JSON.stringify({ event, requestId: id, ...details })}`);
  const path = (owner, id) => `${PREFIX}${owner}/${id}.json`;
  const identity = req => deriveOwnerKey(budgetIdentity(req, mode));
  const validateId = id => {
    if (typeof id !== 'string' || !JOB_ID_RE.test(id)) throw jobError('ai_job_invalid_id', 'A valid AI job identifier is required.', 400);
    return id;
  };
  const summary = job => ({
    id: job.id, status: job.status, createdAt: job.createdAt,
    elapsedMs: Math.max(0, (job.finishedAt || now()) - job.createdAt),
    deadlineAt: job.deadlineAt, pollAfterMs: 2000,
  });

  async function mutate(name, change) {
    for (let attempt = 0; attempt < 8; attempt++) {
      const record = await backend.read(name);
      if (!record) throw jobError('ai_job_not_found', 'The AI job was not found for this account.', 404);
      const value = change(record.value);
      if (!value) return record.value;
      try {
        await backend.replace(name, value, record.etag);
        if (ACTIVE.has(record.value.status) && !ACTIVE.has(value.status)) {
          const code = value.result?.body?.error?.code;
          const succeeded = value.status === 'succeeded';
          log(typeof code === 'string' && /^[a-z_]{1,64}$/.test(code) ? code : `ai_job_${value.status}`, value.id, {
            status: value.status, durationMs: Math.max(0, value.finishedAt - value.createdAt),
          }, succeeded ? 'info' : 'error');
        }
        return value;
      } catch (error) {
        if (![409, 412].includes(Number(error.statusCode))) throw error;
      }
    }
    throw jobError('ai_job_busy', 'The AI job status is busy. Try again shortly.');
  }

  async function current(name) {
    return mutate(name, job => {
      if (ACTIVE.has(job.status) && (job.workerExpiresAt <= now() || job.deadlineAt <= now())) {
        const cancelled = job.status === 'cancelling';
        const timedOut = job.deadlineAt <= now();
        return { ...job, status: cancelled ? 'cancelled' : 'failed', finishedAt: now(),
          result: cancelled ? failure(job.id, 'ai_job_cancelled', 'The AI job was cancelled.', 409)
            : timedOut ? failure(job.id, 'ai_job_timeout', 'The AI job exceeded its 15-minute processing limit.', 504)
              : failure(job.id, 'ai_job_interrupted', 'The AI job was interrupted. It was not automatically resubmitted.') };
      }
      if (job.result && job.finishedAt + resultTtlMs <= now()) {
        const { result, ...record } = job;
        return { ...record, status: 'expired' };
      }
      return null;
    });
  }

  function endpoint(handler) {
    return async (req, res) => {
      res.set('Cache-Control', 'no-store');
      try {
        if (!backend || !executor) throw jobError('ai_jobs_unavailable', 'Asynchronous AI storage is not configured.');
        const owner = identity(req);
        await handler(req, res, owner);
      } catch (error) {
        const id = JOB_ID_RE.test(req.params.id || req.get('Idempotency-Key') || '')
          ? req.params.id || req.get('Idempotency-Key') : crypto.randomUUID();
        const jobCode = typeof error.code === 'string' && error.code.startsWith('ai_job');
        log(jobCode ? error.code : 'ai_job_storage_failed', id);
        const allowed = jobCode || error.code === 'authentication_required';
        const result = failure(id, allowed ? error.code : 'ai_jobs_unavailable',
          allowed ? error.message : 'AI job storage is temporarily unavailable.', allowed ? error.status : 503);
        if (!res.destroyed) res.status(result.status).set(result.headers).json(result.body);
      }
    };
  }

  async function run(req, name, job) {
    const controller = new AbortController();
    if (closed) controller.abort(jobError('ai_job_interrupted', 'The AI service is restarting. This job was not automatically resubmitted.'));
    const state = { controller, work: null };
    active.set(name, state);
    let heartbeatTimer;
    let heartbeatWork;
    const deadline = setTimeout(() => controller.abort(
      jobError('ai_job_timeout', 'The AI job exceeded its 15-minute processing limit.', 504),
    ), Math.max(1, job.deadlineAt - now()));
    deadline.unref?.();
    const heartbeat = () => {
      heartbeatWork = mutate(name, value => {
        if (value.worker !== worker || !ACTIVE.has(value.status)) throw jobError('ai_job_interrupted', 'The AI job no longer belongs to this worker.');
        if (value.status === 'cancelling') {
          controller.abort(jobError('ai_job_cancelled', 'The AI job was cancelled.', 409));
          return null;
        }
        return { ...value, status: 'running', workerExpiresAt: now() + workerLeaseMs };
      }).catch(() => {
        log('ai_job_heartbeat_failed', job.id);
        controller.abort(jobError('ai_job_interrupted', 'The AI job could not renew its ownership.'));
      }).finally(() => {
        if (!controller.signal.aborted) {
          heartbeatTimer = setTimeout(heartbeat, heartbeatMs);
          heartbeatTimer.unref?.();
        }
      });
      return heartbeatWork;
    };
    const work = (async () => {
      let result;
      try {
        await heartbeat();
        if (controller.signal.aborted) throw controller.signal.reason;
        // Only metadata and the eventual result are persisted. req.body and
        // any BYO credential belong exclusively to this live worker.
        result = await executor(req, {
          requestId: job.id, signal: controller.signal, timeoutMs, longRunning: true,
        });
      } catch (error) {
        log('ai_job_execution_failed', job.id);
        result = failure(job.id, 'ai_job_interrupted', 'The AI job was interrupted. It was not automatically resubmitted.');
      } finally {
        const reason = controller.signal.reason;
        controller.abort();
        clearTimeout(deadline);
        clearTimeout(heartbeatTimer);
        await heartbeatWork;
        if (reason?.code?.startsWith('ai_job')) {
          result = failure(job.id, reason.code, reason.message, reason.status);
        }
        try {
          await mutate(name, value => {
            if (!ACTIVE.has(value.status) || value.worker !== worker) return null;
            const cancelled = value.status === 'cancelling' || reason?.code === 'ai_job_cancelled';
            return {
              ...value, finishedAt: now(),
              status: cancelled ? 'cancelled' : result.status >= 200 && result.status < 300 ? 'succeeded' : 'failed',
              result: cancelled ? failure(job.id, 'ai_job_cancelled', 'The AI job was cancelled.', 409) : result,
            };
          });
        } catch {
          // A dead worker/failed write stays recoverable as an interrupted job,
          // never as a second inference dispatch.
          log('ai_job_result_write_failed', job.id);
        }
        active.delete(name);
      }
    })();
    state.work = work;
    await work;
  }

  const submit = endpoint(async (req, res, owner) => {
    if (closed) throw jobError('ai_jobs_unavailable', 'The AI service is restarting. Try again shortly.');
    const id = validateId(req.get('Idempotency-Key'));
    const name = path(owner, id);
    const body = JSON.stringify(req.body);
    if (!body || Buffer.byteLength(body) > 12 * 1024 * 1024) {
      throw jobError('ai_job_invalid_request', 'A bounded AI request body is required.', 400);
    }
    const fingerprint = crypto.createHash('sha256').update(body).digest('hex');
    let record = await backend.read(name);
    if (!record) {
      const retryAfter = await consumeControlRateLimit(req);
      if (retryAfter > 0 || active.size + admitting >= maxActive) {
        res.set('Retry-After', String(retryAfter || 5));
        throw jobError('ai_job_busy', 'AI job admission is busy. Try again shortly.', 429);
      }
      const job = {
        id, fingerprint, worker, status: 'queued', createdAt: now(),
        deadlineAt: now() + timeoutMs, workerExpiresAt: now() + workerLeaseMs,
        deleteAfter: now() + TOMBSTONE_TTL_MS,
      };
      admitting++;
      try {
        await backend.create(name, job);
        // Starting execution is independent of the submission HTTP lifetime.
        const rateHeaders = Object.fromEntries(['x-forwarded-for', 'x-azure-clientip', 'x-azure-socketip']
          .map(header => [header, req.get(header)]));
        const context = {
          body: req.body, ip: req.ip, accessPrincipal: { id: budgetIdentity(req, mode) },
          get: header => rateHeaders[header.toLowerCase()],
        };
        void run(context, name, job);
      } catch (error) {
        if (![409, 412].includes(Number(error.statusCode))) throw error;
      } finally { admitting--; }
      record = await backend.read(name);
    }
    if (!record) throw jobError('ai_jobs_unavailable', 'The accepted AI job could not be read.');
    if (record.value.fingerprint !== null && record.value.fingerprint !== fingerprint) {
      throw jobError('ai_job_conflict', 'This job identifier belongs to a different request. Submit a new job explicitly.', 409);
    }
    const job = await current(name);
    res.set('Location', `/api/openai/jobs/${id}`);
    res.set('Retry-After', '2');
    res.set('X-AzureDiagarm-Request-Id', id);
    return res.status(202).json({ job: summary(job) });
  });

  router.get('/:id', endpoint(async (req, res, owner) => {
    const id = validateId(req.params.id);
    const job = await current(path(owner, id));
    res.set('X-AzureDiagarm-Request-Id', id);
    res.set('Retry-After', '2');
    res.json({ job: summary(job) });
  }));
  router.get('/:id/result', endpoint(async (req, res, owner) => {
    const id = validateId(req.params.id);
    const job = await current(path(owner, id));
    if (ACTIVE.has(job.status)) return res.status(202).json({ job: summary(job) });
    if (!job.result) throw jobError('ai_job_expired', 'The AI job result has expired. Submit a new job explicitly.', 410);
    res.status(job.result.status).set(job.result.headers).type('application/json').send(job.result.body);
  }));
  router.delete('/:id', endpoint(async (req, res, owner) => {
    const id = validateId(req.params.id);
    const name = path(owner, id);
    const record = await backend.read(name);
    if (!record) {
      const retryAfter = await consumeControlRateLimit(req);
      if (retryAfter > 0) throw jobError('ai_job_busy', 'AI job cancellation is busy. Try again shortly.', 429);
      try {
        await backend.create(name, {
          id, fingerprint: null, status: 'cancelled', createdAt: now(), finishedAt: now(),
          deadlineAt: now(), deleteAfter: now() + TOMBSTONE_TTL_MS,
          result: failure(id, 'ai_job_cancelled', 'The AI job was cancelled.', 409),
        });
      } catch (error) {
        if (![409, 412].includes(Number(error.statusCode))) throw error;
      }
    }
    const job = await mutate(name, value => ACTIVE.has(value.status) ? { ...value, status: 'cancelling' } : null);
    active.get(name)?.controller.abort(jobError('ai_job_cancelled', 'The AI job was cancelled.', 409));
    res.json({ job: summary(job) });
  }));

  async function sweep() {
    if (!backend || closed || sweeping) return;
    sweeping = true;
    const startedAt = Date.now();
    let removals = [];
    const flushRemovals = async () => {
      const pending = removals;
      removals = [];
      const results = await Promise.allSettled(pending.map(remove => Promise.resolve().then(remove)));
      const failed = results.find(result => result.status === 'rejected');
      if (failed) throw failed.reason;
    };
    try {
      sweepIterator ||= (backend.listJobBlobs ? backend.listJobBlobs() : backend.list(PREFIX))[Symbol.asyncIterator]();
      for (let count = 0; count < 2000 && Date.now() - startedAt < 20_000; count++) {
        const entry = await io(() => sweepIterator.next());
        if (entry.done) { sweepIterator = undefined; break; }
        const name = typeof entry.value === 'string' ? entry.value : entry.value.name;
        if (!name.startsWith(PREFIX)) throw new Error('Job cleanup returned an unrelated blob.');
        if (entry.value.deleted) continue;
        if (entry.value.versionId && entry.value.isCurrentVersion === false) {
          // Previous versions never participate in job recovery or ETag
          // updates. Keep the current version; Azure soft-delete still applies.
          removals.push(() => backend.removeJobVersion(name, entry.value.versionId));
          if (removals.length === 8) await flushRemovals();
          continue;
        }
        const record = await backend.read(name);
        if (!record) continue;
        if (record.value.deleteAfter <= now()) {
          try { await backend.remove(name, record.etag); }
          catch (error) { if (![404, 412].includes(Number(error.statusCode))) throw error; }
        } else await current(name);
      }
      await flushRemovals();
    } catch (error) {
      sweepIterator = undefined;
      throw error;
    } finally {
      sweeping = false;
    }
  }

  return {
    router, submit, sweep,
    setExecutor(value) { executor = value; },
    async close() {
      closed = true;
      const running = [...active.values()];
      running.forEach(value => value.controller.abort(jobError('ai_job_interrupted', 'The AI service restarted. This job was not automatically resubmitted.')));
      await Promise.all(running.map(value => value.work));
    },
  };
}

module.exports = { createAIJobs, createMemoryJobBackend, JOB_TIMEOUT_MS, RESULT_TTL_MS };
