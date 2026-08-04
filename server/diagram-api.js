// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

/**
 * Authenticated cloud persistence for diagram documents.
 *
 * Storage layout (Azure Blob Storage, keyless via DefaultAzureCredential):
 *
 *   owners/<ownerKey>/documents/<docId>/current.json      current document
 *   owners/<ownerKey>/documents/<docId>/versions/<id>.json immutable snapshots
 *   shares/<tokenHash>.json                                 share-token index
 *
 * `ownerKey` is SHA-256(principal.id) so one caller cannot enumerate another
 * caller's blobs by guessing a document UUID, and the share index is keyed by
 * SHA-256(token) so raw share tokens are never persisted. Easy Auth headers are
 * authoritative for identity (see access-control.getPrincipal).
 *
 * The router is written against a small `backend` abstraction (read / create /
 * replace / put / remove / list) so it can be unit-tested with an in-memory
 * fake, while the Azure-backed implementation lives in createAzureBlobBackend.
 * `remove(name, etag)` performs an optional conditional delete.
 */

const crypto = require('crypto');
const express = require('express');
const { asyncHandler } = require('./async-handler');

const MAX_DOCUMENTS_LISTED = 200;
const MAX_VERSIONS_LISTED = 200;
const MAX_COMMENTS = 200;
const MAX_SHARES = 100;
const MAX_NAME_LEN = 200;
const MAX_COMMENT_LEN = 2000;
const MAX_NOTES_LEN = 500;
const MAX_PAYLOAD_BYTES = 10 * 1024 * 1024;
const MAX_NODES = 20_000;
const MAX_EDGES = 40_000;
const SHARE_TOKEN_BYTES = 32; // 256-bit
// crypto.randomBytes(32).toString('base64url') is always 43 base64url chars.
const SHARE_TOKEN_RE = /^[A-Za-z0-9_-]{43}$/;
const IDEMPOTENCY_KEY_RE = /^[A-Za-z0-9._~-]{16,128}$/;
const VALID_ROLES = new Set(['viewer', 'editor']);
const DELETION_MARKER_KIND = 'diagram-deletion-marker-v1';

// ── Identity / hashing helpers ─────────────────────────────────────────────

function deriveOwnerKey(principalId) {
  return crypto.createHash('sha256').update(String(principalId)).digest('hex');
}

function hashToken(token) {
  return crypto.createHash('sha256').update(String(token)).digest('hex');
}

function hashCreateRequest(diagramName, payload) {
  return crypto
    .createHash('sha256')
    .update(JSON.stringify({ diagramName, payload }))
    .digest('hex');
}

function documentIdForIdempotencyKey(key) {
  const hex = crypto.createHash('sha256').update(String(key)).digest('hex');
  return [
    hex.slice(0, 8),
    hex.slice(8, 12),
    `5${hex.slice(13, 16)}`,
    `a${hex.slice(17, 20)}`,
    hex.slice(20, 32),
  ].join('-');
}

// ── Blob path helpers ──────────────────────────────────────────────────────

function documentPrefix(ownerKey, documentId) {
  return `owners/${ownerKey}/documents/${documentId}/`;
}

function currentPath(ownerKey, documentId) {
  return `${documentPrefix(ownerKey, documentId)}current.json`;
}

function versionsPrefix(ownerKey, documentId) {
  return `${documentPrefix(ownerKey, documentId)}versions/`;
}

function versionPath(ownerKey, documentId, versionId) {
  return `${versionsPrefix(ownerKey, documentId)}${versionId}.json`;
}

function sharePath(tokenHash) {
  return `shares/${tokenHash}.json`;
}

function documentLockPath(ownerKey, documentId) {
  return `owners/${ownerKey}/document-locks/${documentId}.lock`;
}

function isDeletionMarker(value) {
  return value?.kind === DELETION_MARKER_KIND;
}

function buildDeletionMarker(doc, sourceEtag) {
  return {
    kind: DELETION_MARKER_KIND,
    documentId: doc.id,
    sourceEtag,
    shareTokenHashes: (doc.shares || [])
      .map((share) => share.tokenHash)
      .filter(Boolean),
    deletedAt: new Date().toISOString(),
  };
}

// ── Validation helpers ─────────────────────────────────────────────────────

function validateDiagramName(value) {
  if (typeof value !== 'string') return { error: 'diagramName is required' };
  const name = value.trim();
  if (name.length === 0) return { error: 'diagramName must not be empty' };
  if (name.length > MAX_NAME_LEN) return { error: `diagramName must be at most ${MAX_NAME_LEN} characters` };
  return { value: name };
}

function validatePayload(payload) {
  if (!payload || typeof payload !== 'object' || Array.isArray(payload)) {
    return { error: 'payload must be an object' };
  }
  if (!Array.isArray(payload.nodes) || !Array.isArray(payload.edges)) {
    return { error: 'payload must include nodes and edges arrays' };
  }
  if (payload.nodes.length > MAX_NODES) return { error: `payload.nodes exceeds ${MAX_NODES} entries` };
  if (payload.edges.length > MAX_EDGES) return { error: `payload.edges exceeds ${MAX_EDGES} entries` };
  let serialized;
  try {
    serialized = JSON.stringify(payload);
  } catch {
    return { error: 'payload is not serializable' };
  }
  if (Buffer.byteLength(serialized) > MAX_PAYLOAD_BYTES) {
    return { error: 'payload exceeds the maximum size' };
  }
  return { value: payload };
}

// ── Response shaping ───────────────────────────────────────────────────────
//
// Wire shapes mirror the frontend cloud client (src/services/cloudDiagramService.ts):
// documents are returned under a `document` key, collections under
// `documents` / `versions` / `shares`, and field names use the client's
// vocabulary (serviceCount/connectionCount, commentId/message, shareId,
// versionId/sourceRevision, createdByEmail).

function sanitizeComment(comment) {
  return {
    commentId: comment.id,
    message: comment.text,
    authorId: comment.authorId,
    authorEmail: comment.authorEmail,
    createdAt: comment.createdAt,
  };
}

// Never leak the stored token hash to any client.
function sanitizeShare(share) {
  return {
    shareId: share.id,
    role: share.role,
    createdAt: share.createdAt,
    createdByEmail: share.createdBy,
  };
}

function sanitizeDocument(doc, access, role) {
  const owner = access === 'owner'
    ? { id: doc.owner?.id, email: doc.owner?.email }
    : { email: doc.owner?.email };
  // Only the owner sees the collaborator/share list; shared viewers and editors
  // must not learn who else the document is shared with.
  const shares = access === 'owner' && Array.isArray(doc.shares)
    ? doc.shares.map(sanitizeShare)
    : [];
  return {
    id: doc.id,
    diagramName: doc.diagramName,
    owner,
    payload: doc.payload,
    comments: Array.isArray(doc.comments) ? doc.comments.map(sanitizeComment) : [],
    shares,
    createdAt: doc.createdAt,
    updatedAt: doc.updatedAt,
    revision: doc.revision,
    access,
    role,
  };
}

function documentSummary(doc, etag, access, role) {
  const services = Array.isArray(doc.payload?.nodes) ? doc.payload.nodes.length : 0;
  const connections = Array.isArray(doc.payload?.edges) ? doc.payload.edges.length : 0;
  return {
    id: doc.id,
    diagramName: doc.diagramName,
    createdAt: doc.createdAt,
    updatedAt: doc.updatedAt,
    revision: doc.revision,
    serviceCount: services,
    connectionCount: connections,
    commentCount: Array.isArray(doc.comments) ? doc.comments.length : 0,
    shareCount: Array.isArray(doc.shares) ? doc.shares.length : 0,
    access,
    role,
    etag,
  };
}

function versionSummary(version) {
  return {
    versionId: version.versionId,
    diagramId: version.diagramId,
    diagramName: version.diagramName,
    notes: version.notes,
    createdAt: version.createdAt,
    createdByEmail: version.createdByEmail,
    sourceRevision: version.sourceRevision,
  };
}

// ── Storage error mapping ──────────────────────────────────────────────────
//
// The Azure SDK surfaces conditional-write and lookup failures as RestError
// with a numeric statusCode. Map only the ones the API contract cares about;
// everything else is treated as a transient storage outage (503) and logged so
// the failure is never silently swallowed.
function storageErrorStatus(err) {
  const code = Number(err?.statusCode ?? err?.status ?? err?.response?.status);
  if (code === 404) return 404;
  if (code === 409) return 409;
  if (code === 412) return 412;
  return 503;
}

function respondStorageError(res, logger, context, err) {
  if (res.headersSent || res.writableEnded) {
    logger.error(
      `[diagrams] ${context} storage error after the response was sent:`,
      err?.message || err,
    );
    return undefined;
  }
  const status = storageErrorStatus(err);
  if (status === 412) {
    return res.status(412).json({ error: 'The document was modified by another request. Reload and retry.' });
  }
  if (status === 409) {
    return res.status(409).json({ error: 'The document already exists or was modified concurrently.' });
  }
  if (status === 404) {
    return res.status(404).json({ error: 'Not found' });
  }
  logger.error(`[diagrams] ${context} storage error:`, err?.message || err);
  return res.status(503).json({ error: 'Diagram storage is temporarily unavailable' });
}

// ── Router factory ─────────────────────────────────────────────────────────

function createDiagramsRouter(options = {}) {
  const {
    backend = null,
    getPrincipal,
    publicUrl = '',
    logger = console,
  } = options;

  if (typeof getPrincipal !== 'function') {
    throw new TypeError('getPrincipal must be a function');
  }

  const router = express.Router();
  const shareBaseUrl = String(publicUrl || '').replace(/\/+$/, '');

  function shareUrlFor(token) {
    return `${shareBaseUrl}/#share-${token}`;
  }

  // Every diagram route requires a valid Easy Auth principal and configured
  // storage. Identity is checked first so an unauthenticated caller never
  // learns whether storage happens to be configured.
  router.use((req, res, next) => {
    const principal = getPrincipal(req);
    if (!principal) {
      return res.status(401).json({ error: 'Authentication is required.' });
    }
    if (!backend) {
      return res.status(503).json({ error: 'Diagram storage is not configured.' });
    }
    req.diagramPrincipal = principal;
    req.diagramOwnerKey = deriveOwnerKey(principal.id);
    next();
  });

  // ── Internal store operations (all owner-scoped) ─────────────────────────

  async function loadCurrent(ownerKey, documentId) {
    const record = await backend.read(currentPath(ownerKey, documentId));
    return record && !isDeletionMarker(record.value) ? record : null;
  }

  function withDocumentLock(ownerKey, documentId, operation) {
    if (typeof backend.withLock !== 'function') return operation();
    return backend.withLock(documentLockPath(ownerKey, documentId), operation);
  }

  function buildNewDocument(principal, name, payload, createRequest) {
    const now = new Date().toISOString();
    return {
      id: createRequest?.documentId || crypto.randomUUID(),
      diagramName: name,
      owner: { id: principal.id, email: principal.email },
      payload,
      comments: [],
      shares: [],
      createdAt: now,
      updatedAt: now,
      revision: 1,
      ...(createRequest ? { createRequest } : {}),
    };
  }

  function buildComment(principal, text) {
    return {
      id: crypto.randomUUID(),
      authorId: principal.id,
      authorEmail: principal.email,
      text,
      createdAt: new Date().toISOString(),
    };
  }

  function buildVersion(doc, principal, notes) {
    const now = new Date().toISOString();
    return {
      versionId: crypto.randomUUID(),
      diagramId: doc.id,
      diagramName: doc.diagramName,
      payload: doc.payload,
      comments: Array.isArray(doc.comments) ? doc.comments.map(sanitizeComment) : [],
      notes,
      createdAt: now,
      createdByEmail: principal.email,
      sourceRevision: doc.revision,
    };
  }

  // Apply an updated payload/name while preserving all server-owned fields.
  function applyUpdate(doc, name, payload) {
    return {
      ...doc,
      diagramName: name,
      payload,
      updatedAt: new Date().toISOString(),
      revision: (Number(doc.revision) || 0) + 1,
    };
  }

  function setEtag(res, etag) {
    if (etag) res.set('ETag', etag);
  }

  // Resolve a share token to its owning document, rejecting orphaned indexes.
  async function resolveShare(token) {
    const tokenHash = hashToken(token);
    const index = await backend.read(sharePath(tokenHash));
    if (!index) return { status: 404 };
    const current = await loadCurrent(index.value.ownerKey, index.value.documentId);
    if (!current) return { status: 404 };
    const share = (current.value.shares || []).find(
      (entry) => entry.tokenHash === tokenHash && entry.id === index.value.shareId,
    );
    if (!share) return { status: 404 }; // index without a live share entry
    return {
      ownerKey: index.value.ownerKey,
      documentId: index.value.documentId,
      doc: current.value,
      etag: current.etag,
      share,
    };
  }

  // ── Owner: list / create ─────────────────────────────────────────────────

  router.get('/', asyncHandler(async (req, res) => {
    const ownerKey = req.diagramOwnerKey;
    try {
      const prefix = `owners/${ownerKey}/documents/`;
      const summaries = [];
      for await (const name of backend.list(prefix)) {
        if (!name.endsWith('/current.json')) continue;
        const record = await backend.read(name);
        if (record && !isDeletionMarker(record.value)) {
          summaries.push(documentSummary(record.value, record.etag, 'owner', 'owner'));
        }
        if (summaries.length >= MAX_DOCUMENTS_LISTED * 2) break;
      }
      summaries.sort((a, b) => String(b.updatedAt).localeCompare(String(a.updatedAt)));
      res.json({ documents: summaries.slice(0, MAX_DOCUMENTS_LISTED) });
    } catch (err) {
      return respondStorageError(res, logger, 'list', err);
    }
  }));

  router.post('/', asyncHandler(async (req, res) => {
    const body = req.body || {};
    const name = validateDiagramName(body.diagramName);
    if (name.error) return res.status(400).json({ error: name.error });
    const payload = validatePayload(body.payload);
    if (payload.error) return res.status(400).json({ error: payload.error });

    const idempotencyKey = req.get('idempotency-key');
    if (idempotencyKey && !IDEMPOTENCY_KEY_RE.test(idempotencyKey)) {
      return res.status(400).json({ error: 'Idempotency-Key must be 16 to 128 URL-safe characters.' });
    }
    const createRequest = idempotencyKey
      ? {
          keyHash: hashToken(idempotencyKey),
          requestHash: hashCreateRequest(name.value, payload.value),
          documentId: documentIdForIdempotencyKey(idempotencyKey),
        }
      : null;
    const doc = buildNewDocument(req.diagramPrincipal, name.value, payload.value, createRequest);
    const path = currentPath(req.diagramOwnerKey, doc.id);
    try {
      const { etag } = await backend.create(path, doc);
      setEtag(res, etag);
      res.status(201).json({ document: sanitizeDocument(doc, 'owner', 'owner') });
    } catch (err) {
      if (createRequest && storageErrorStatus(err) === 409) {
        try {
          const existing = await backend.read(path);
          if (
            existing
            && !isDeletionMarker(existing.value)
            && existing.value?.createRequest?.keyHash === createRequest.keyHash
            && existing.value?.createRequest?.requestHash === createRequest.requestHash
          ) {
            setEtag(res, existing.etag);
            res.set('Idempotency-Replayed', 'true');
            return res.status(200).json({
              document: sanitizeDocument(existing.value, 'owner', 'owner'),
            });
          }
          return res.status(409).json({
            error: 'The Idempotency-Key was already used for a different create request.',
          });
        } catch (readError) {
          return respondStorageError(res, logger, 'create replay', readError);
        }
      }
      return respondStorageError(res, logger, 'create', err);
    }
  }));

  // ── Owner: get / update / delete a document ──────────────────────────────

  router.get('/:id', asyncHandler(async (req, res) => {
    try {
      const record = await loadCurrent(req.diagramOwnerKey, req.params.id);
      if (!record) return res.status(404).json({ error: 'Not found' });
      setEtag(res, record.etag);
      res.json({ document: sanitizeDocument(record.value, 'owner', 'owner') });
    } catch (err) {
      return respondStorageError(res, logger, 'get', err);
    }
  }));

  router.put('/:id', asyncHandler(async (req, res) => {
    const ifMatch = req.get('if-match');
    if (!ifMatch) return res.status(428).json({ error: 'If-Match header is required' });

    const body = req.body || {};
    const name = validateDiagramName(body.diagramName);
    if (name.error) return res.status(400).json({ error: name.error });
    const payload = validatePayload(body.payload);
    if (payload.error) return res.status(400).json({ error: payload.error });

    try {
      const record = await loadCurrent(req.diagramOwnerKey, req.params.id);
      if (!record) return res.status(404).json({ error: 'Not found' });
      const updated = applyUpdate(record.value, name.value, payload.value);
      const { etag } = await backend.replace(
        currentPath(req.diagramOwnerKey, req.params.id),
        updated,
        ifMatch,
      );
      setEtag(res, etag);
      res.json({ document: sanitizeDocument(updated, 'owner', 'owner') });
    } catch (err) {
      return respondStorageError(res, logger, 'update', err);
    }
  }));

  router.delete('/:id', asyncHandler(async (req, res) => {
    const ifMatch = req.get('if-match');
    if (!ifMatch) return res.status(428).json({ error: 'If-Match header is required' });

    try {
      const outcome = await withDocumentLock(req.diagramOwnerKey, req.params.id, async () => {
        const currentName = currentPath(req.diagramOwnerKey, req.params.id);
        const record = await backend.read(currentName);
        if (!record) return { status: 404, error: 'Not found' };
        if (isDeletionMarker(record.value)) {
          if (record.value.sourceEtag !== ifMatch && record.etag !== ifMatch) {
            return {
              status: 412,
              error: 'The document was modified by another request. Reload and retry.',
            };
          }
        } else if (record.etag !== ifMatch) {
          return {
            status: 412,
            error: 'The document was modified by another request. Reload and retry.',
          };
        }

        const marker = isDeletionMarker(record.value)
          ? record.value
          : buildDeletionMarker(record.value, ifMatch);
        const markerEtag = isDeletionMarker(record.value)
          ? record.etag
          : (await backend.replace(currentName, marker, ifMatch)).etag;

        // A marker makes the document immediately inaccessible while preserving
        // enough state to retry cleanup after any transient storage failure.
        for (const tokenHash of marker.shareTokenHashes || []) {
          await backend.remove(sharePath(tokenHash));
        }
        const prefix = documentPrefix(req.diagramOwnerKey, req.params.id);
        const names = [];
        for await (const name of backend.list(prefix)) {
          if (name !== currentName) names.push(name);
        }
        for (const name of names) await backend.remove(name);
        await backend.remove(currentName, markerEtag);
        return { status: 204 };
      });
      if (outcome.error) return res.status(outcome.status).json({ error: outcome.error });
      return res.status(outcome.status).end();
    } catch (err) {
      return respondStorageError(res, logger, 'delete', err);
    }
  }));

  // ── Owner: versions ──────────────────────────────────────────────────────

  router.post('/:id/versions', asyncHandler(async (req, res) => {
    const notesRaw = (req.body || {}).notes;
    if (notesRaw != null && typeof notesRaw !== 'string') {
      return res.status(400).json({ error: 'notes must be a string' });
    }
    if (typeof notesRaw === 'string' && notesRaw.length > MAX_NOTES_LEN) {
      return res.status(400).json({ error: `notes must be at most ${MAX_NOTES_LEN} characters` });
    }
    const notes = typeof notesRaw === 'string' ? notesRaw : '';

    try {
      const outcome = await withDocumentLock(req.diagramOwnerKey, req.params.id, async () => {
        const record = await loadCurrent(req.diagramOwnerKey, req.params.id);
        if (!record) return { status: 404, error: 'Not found' };
        const version = buildVersion(record.value, req.diagramPrincipal, notes);
        await backend.put(versionPath(req.diagramOwnerKey, req.params.id, version.versionId), version);
        return { status: 201, version };
      });
      if (outcome.error) return res.status(outcome.status).json({ error: outcome.error });
      return res.status(outcome.status).json({ version: outcome.version });
    } catch (err) {
      return respondStorageError(res, logger, 'version-create', err);
    }
  }));

  router.get('/:id/versions', asyncHandler(async (req, res) => {
    try {
      const record = await loadCurrent(req.diagramOwnerKey, req.params.id);
      if (!record) return res.status(404).json({ error: 'Not found' });
      const versions = await listVersionSummaries(req.diagramOwnerKey, req.params.id);
      res.json({ versions });
    } catch (err) {
      return respondStorageError(res, logger, 'version-list', err);
    }
  }));

  router.get('/:id/versions/:versionId', asyncHandler(async (req, res) => {
    try {
      const current = await loadCurrent(req.diagramOwnerKey, req.params.id);
      if (!current) return res.status(404).json({ error: 'Not found' });
      const record = await backend.read(
        versionPath(req.diagramOwnerKey, req.params.id, req.params.versionId),
      );
      if (!record) return res.status(404).json({ error: 'Not found' });
      res.json({ version: record.value });
    } catch (err) {
      return respondStorageError(res, logger, 'version-get', err);
    }
  }));

  async function listVersionSummaries(ownerKey, documentId) {
    const summaries = [];
    for await (const name of backend.list(versionsPrefix(ownerKey, documentId))) {
      const record = await backend.read(name);
      if (record) summaries.push(versionSummary(record.value));
      if (summaries.length >= MAX_VERSIONS_LISTED * 2) break;
    }
    summaries.sort((a, b) => String(b.createdAt).localeCompare(String(a.createdAt)));
    return summaries.slice(0, MAX_VERSIONS_LISTED);
  }

  // ── Owner: comments ──────────────────────────────────────────────────────

  function validateCommentText(value) {
    if (typeof value !== 'string' || value.trim().length === 0) {
      return { error: 'message is required' };
    }
    if (value.length > MAX_COMMENT_LEN) {
      return { error: `message must be at most ${MAX_COMMENT_LEN} characters` };
    }
    return { value };
  }

  // The frontend client posts `message`; accept `text` too for resilience.
  function commentTextFrom(body) {
    const source = body || {};
    return typeof source.message === 'string' ? source.message : source.text;
  }

  // Append a comment to a freshly-read document, then persist with optimistic
  // concurrency on the read etag. Shared by owner and shared-token callers.
  async function appendCommentAndSave(res, ownerKey, documentId, principal, text, access, role) {
    const record = await loadCurrent(ownerKey, documentId);
    if (!record) return res.status(404).json({ error: 'Not found' });
    const doc = record.value;
    doc.comments = Array.isArray(doc.comments) ? doc.comments : [];
    if (doc.comments.length >= MAX_COMMENTS) {
      return res.status(409).json({ error: 'This document has reached its comment limit.' });
    }
    doc.comments.push(buildComment(principal, text));
    doc.updatedAt = new Date().toISOString();
    doc.revision = (Number(doc.revision) || 0) + 1;
    const { etag } = await backend.replace(currentPath(ownerKey, documentId), doc, record.etag);
    setEtag(res, etag);
    return res.status(201).json({ document: sanitizeDocument(doc, access, role) });
  }

  router.post('/:id/comments', asyncHandler(async (req, res) => {
    const text = validateCommentText(commentTextFrom(req.body));
    if (text.error) return res.status(400).json({ error: text.error });
    try {
      return await appendCommentAndSave(
        res, req.diagramOwnerKey, req.params.id, req.diagramPrincipal, text.value, 'owner', 'owner',
      );
    } catch (err) {
      return respondStorageError(res, logger, 'comment', err);
    }
  }));

  // ── Owner: shares ────────────────────────────────────────────────────────

  router.post('/:id/shares', asyncHandler(async (req, res) => {
    const role = (req.body || {}).role;
    if (!VALID_ROLES.has(role)) {
      return res.status(400).json({ error: "role must be 'viewer' or 'editor'" });
    }
    try {
      const record = await loadCurrent(req.diagramOwnerKey, req.params.id);
      if (!record) return res.status(404).json({ error: 'Not found' });
      const doc = record.value;
      doc.shares = Array.isArray(doc.shares) ? doc.shares : [];
      if (doc.shares.length >= MAX_SHARES) {
        return res.status(409).json({ error: 'This document has reached its share limit.' });
      }

      const token = crypto.randomBytes(SHARE_TOKEN_BYTES).toString('base64url');
      const tokenHash = hashToken(token);
      const shareId = crypto.randomUUID();
      const createdAt = new Date().toISOString();
      const shareEntry = {
        id: shareId,
        role,
        tokenHash,
        createdAt,
        createdBy: req.diagramPrincipal.email,
      };

      // Write the hash-only index first, then commit the document. If the
      // document write fails we roll the index back so no orphan resolves.
      await backend.put(sharePath(tokenHash), {
        ownerKey: req.diagramOwnerKey,
        documentId: doc.id,
        shareId,
        role,
        tokenHash,
        createdAt,
      });
      doc.shares.push(shareEntry);
      doc.updatedAt = createdAt;
      doc.revision = (Number(doc.revision) || 0) + 1;
      try {
        const { etag } = await backend.replace(
          currentPath(req.diagramOwnerKey, req.params.id), doc, record.etag,
        );
        setEtag(res, etag);
      } catch (err) {
        await backend.remove(sharePath(tokenHash));
        throw err;
      }

      res.status(201).json({
        share: sanitizeShare(shareEntry),
        token,
        url: shareUrlFor(token),
      });
    } catch (err) {
      return respondStorageError(res, logger, 'share-create', err);
    }
  }));

  router.get('/:id/shares', asyncHandler(async (req, res) => {
    try {
      const record = await loadCurrent(req.diagramOwnerKey, req.params.id);
      if (!record) return res.status(404).json({ error: 'Not found' });
      const shares = (record.value.shares || []).map(sanitizeShare);
      res.json({ shares });
    } catch (err) {
      return respondStorageError(res, logger, 'share-list', err);
    }
  }));

  router.delete('/:id/shares/:shareId', asyncHandler(async (req, res) => {
    try {
      const record = await loadCurrent(req.diagramOwnerKey, req.params.id);
      if (!record) return res.status(404).json({ error: 'Not found' });
      const doc = record.value;
      const share = (doc.shares || []).find((entry) => entry.id === req.params.shareId);
      if (!share) return res.status(404).json({ error: 'Not found' });

      // Security-first: drop the index before rewriting the document so the
      // token stops resolving even if the document write is delayed.
      if (share.tokenHash) await backend.remove(sharePath(share.tokenHash));
      doc.shares = doc.shares.filter((entry) => entry.id !== req.params.shareId);
      doc.updatedAt = new Date().toISOString();
      doc.revision = (Number(doc.revision) || 0) + 1;
      await backend.replace(currentPath(req.diagramOwnerKey, req.params.id), doc, record.etag);
      res.status(204).end();
    } catch (err) {
      return respondStorageError(res, logger, 'share-revoke', err);
    }
  }));

  // ── Shared access (via token) ────────────────────────────────────────────

  function requireToken(req, res) {
    const token = req.params.token;
    if (!SHARE_TOKEN_RE.test(String(token || ''))) {
      res.status(404).json({ error: 'Not found' });
      return null;
    }
    return token;
  }

  router.get('/shared/:token', asyncHandler(async (req, res) => {
    const token = requireToken(req, res);
    if (!token) return undefined;
    try {
      const resolved = await resolveShare(token);
      if (resolved.status) return res.status(resolved.status).json({ error: 'Not found' });
      setEtag(res, resolved.etag);
      const isOwner = resolved.doc.owner?.id === req.diagramPrincipal.id;
      res.json({
        document: sanitizeDocument(
          resolved.doc,
          isOwner ? 'owner' : 'shared',
          isOwner ? 'owner' : resolved.share.role,
        ),
      });
    } catch (err) {
      return respondStorageError(res, logger, 'shared-get', err);
    }
  }));

  router.put('/shared/:token', asyncHandler(async (req, res) => {
    const token = requireToken(req, res);
    if (!token) return undefined;
    const ifMatch = req.get('if-match');
    if (!ifMatch) return res.status(428).json({ error: 'If-Match header is required' });

    const body = req.body || {};
    const name = validateDiagramName(body.diagramName);
    if (name.error) return res.status(400).json({ error: name.error });
    const payload = validatePayload(body.payload);
    if (payload.error) return res.status(400).json({ error: payload.error });

    try {
      const resolved = await resolveShare(token);
      if (resolved.status) return res.status(resolved.status).json({ error: 'Not found' });
      if (resolved.share.role !== 'editor') {
        return res.status(403).json({ error: 'Editor access is required to modify this document.' });
      }
      const updated = applyUpdate(resolved.doc, name.value, payload.value);
      const { etag } = await backend.replace(
        currentPath(resolved.ownerKey, resolved.documentId),
        updated,
        ifMatch,
      );
      setEtag(res, etag);
      res.json({ document: sanitizeDocument(updated, 'shared', 'editor') });
    } catch (err) {
      return respondStorageError(res, logger, 'shared-update', err);
    }
  }));

  router.post('/shared/:token/comments', asyncHandler(async (req, res) => {
    const token = requireToken(req, res);
    if (!token) return undefined;
    const text = validateCommentText(commentTextFrom(req.body));
    if (text.error) return res.status(400).json({ error: text.error });
    try {
      const resolved = await resolveShare(token);
      if (resolved.status) return res.status(resolved.status).json({ error: 'Not found' });
      return await appendCommentAndSave(
        res, resolved.ownerKey, resolved.documentId, req.diagramPrincipal, text.value,
        'shared', resolved.share.role,
      );
    } catch (err) {
      return respondStorageError(res, logger, 'shared-comment', err);
    }
  }));

  router.post('/shared/:token/versions', asyncHandler(async (req, res) => {
    const token = requireToken(req, res);
    if (!token) return undefined;
    const notesRaw = (req.body || {}).notes;
    if (notesRaw != null && typeof notesRaw !== 'string') {
      return res.status(400).json({ error: 'notes must be a string' });
    }
    if (typeof notesRaw === 'string' && notesRaw.length > MAX_NOTES_LEN) {
      return res.status(400).json({ error: `notes must be at most ${MAX_NOTES_LEN} characters` });
    }
    const notes = typeof notesRaw === 'string' ? notesRaw : '';
    try {
      const resolved = await resolveShare(token);
      if (resolved.status) return res.status(resolved.status).json({ error: 'Not found' });
      if (resolved.share.role !== 'editor') {
        return res.status(403).json({ error: 'Editor access is required to snapshot this document.' });
      }
      const outcome = await withDocumentLock(resolved.ownerKey, resolved.documentId, async () => {
        const current = await resolveShare(token);
        if (current.status) return { status: current.status, error: 'Not found' };
        if (current.share.role !== 'editor') {
          return {
            status: 403,
            error: 'Editor access is required to snapshot this document.',
          };
        }
        const version = buildVersion(current.doc, req.diagramPrincipal, notes);
        await backend.put(versionPath(current.ownerKey, current.documentId, version.versionId), version);
        return { status: 201, version };
      });
      if (outcome.error) return res.status(outcome.status).json({ error: outcome.error });
      return res.status(outcome.status).json({ version: outcome.version });
    } catch (err) {
      return respondStorageError(res, logger, 'shared-version-create', err);
    }
  }));

  router.get('/shared/:token/versions', asyncHandler(async (req, res) => {
    const token = requireToken(req, res);
    if (!token) return undefined;
    try {
      const resolved = await resolveShare(token);
      if (resolved.status) return res.status(resolved.status).json({ error: 'Not found' });
      const versions = await listVersionSummaries(resolved.ownerKey, resolved.documentId);
      res.json({ versions });
    } catch (err) {
      return respondStorageError(res, logger, 'shared-version-list', err);
    }
  }));

  router.get('/shared/:token/versions/:versionId', asyncHandler(async (req, res) => {
    const token = requireToken(req, res);
    if (!token) return undefined;
    try {
      const resolved = await resolveShare(token);
      if (resolved.status) return res.status(resolved.status).json({ error: 'Not found' });
      const record = await backend.read(
        versionPath(resolved.ownerKey, resolved.documentId, req.params.versionId),
      );
      if (!record) return res.status(404).json({ error: 'Not found' });
      res.json({ version: record.value });
    } catch (err) {
      return respondStorageError(res, logger, 'shared-version-get', err);
    }
  }));

  return router;
}

// ── Azure Blob Storage backend ─────────────────────────────────────────────

function streamToString(readable) {
  return new Promise((resolve, reject) => {
    const chunks = [];
    readable.on('data', (chunk) => chunks.push(Buffer.isBuffer(chunk) ? chunk : Buffer.from(chunk)));
    readable.on('end', () => resolve(Buffer.concat(chunks).toString('utf8')));
    readable.on('error', reject);
  });
}

/**
 * Wrap an Azure Blob container as the small backend the router expects. Uses
 * DefaultAzureCredential (managed identity in ACA, `az login` in dev) — no
 * account keys or SAS. Conditional writes rely on blob ETags for optimistic
 * concurrency (ifMatch) and create-only semantics (ifNoneMatch '*').
 */
function createAzureBlobBackend(options = {}) {
  const {
    endpoint,
    containerName = 'diagrams',
    credential,
    containerClient,
    leaseRenewalIntervalMs = 30_000,
    logger = console,
  } = options;
  let container = containerClient;
  if (!container) {
    if (!endpoint) throw new Error('AZURE_BLOB_ENDPOINT is required');
    if (!credential) throw new Error('credential is required');
    // Lazy require so unit tests that inject a fake backend do not need the SDK.
    const { ContainerClient } = require('@azure/storage-blob');
    const base = String(endpoint).replace(/\/+$/, '');
    container = new ContainerClient(`${base}/${encodeURIComponent(containerName)}`, credential);
  }

  let ensured = null;
  async function ensureContainer() {
    if (!ensured) {
      ensured = container.createIfNotExists().catch((error) => {
        ensured = null;
        throw error;
      });
    }
    await ensured;
  }

  function blob(name) {
    return container.getBlockBlobClient(name);
  }

  async function upload(name, value, conditions) {
    await ensureContainer();
    const body = JSON.stringify(value);
    const resp = await blob(name).upload(body, Buffer.byteLength(body), {
      conditions,
      blobHTTPHeaders: { blobContentType: 'application/json' },
    });
    return { etag: resp.etag };
  }

  async function ensureLockBlob(name) {
    try {
      await upload(name, { kind: 'diagram-document-lock-v1' }, { ifNoneMatch: '*' });
    } catch (error) {
      if (storageErrorStatus(error) !== 409) throw error;
    }
  }

  async function withLock(name, operation) {
    await ensureLockBlob(name);
    const leaseClient = blob(name).getBlobLeaseClient();
    const deadline = Date.now() + 15_000;
    while (true) {
      try {
        await leaseClient.acquireLease(60);
        break;
      } catch (error) {
        if (storageErrorStatus(error) !== 409 || Date.now() >= deadline) {
          if (storageErrorStatus(error) === 409) {
            const timeoutError = new Error('Timed out waiting for the document lock');
            timeoutError.statusCode = 503;
            throw timeoutError;
          }
          throw error;
        }
        await new Promise((resolve) => setTimeout(resolve, 100 + Math.floor(Math.random() * 150)));
      }
    }

    let renewalError = null;
    let renewalPromise = null;
    const renewalTimer = setInterval(() => {
      if (renewalPromise) return;
      renewalPromise = leaseClient.renewLease()
        .then(() => {
          // A later successful renewal confirms that a transient failure did
          // not cause the lease to be lost.
          renewalError = null;
        })
        .catch((error) => {
          renewalError = error;
        })
        .finally(() => {
          renewalPromise = null;
        });
    }, leaseRenewalIntervalMs);

    let result;
    let operationError = null;
    try {
      result = await operation();
    } catch (error) {
      operationError = error;
    } finally {
      clearInterval(renewalTimer);
    }

    if (renewalPromise) await renewalPromise;

    let releaseError = null;
    try {
      await leaseClient.releaseLease();
    } catch (error) {
      releaseError = error;
    }

    if (operationError) {
      if (releaseError) {
        logger.warn(
          '[diagrams] Document lock release failed while handling an operation error:',
          releaseError?.message || releaseError,
        );
      }
      throw operationError;
    }

    if (renewalError && releaseError) {
      logger.error(
        '[diagrams] Document lock ownership could not be confirmed after the operation:',
        renewalError?.message || renewalError,
      );
      const lockError = new Error('Document lock ownership could not be confirmed');
      lockError.statusCode = 503;
      lockError.cause = renewalError;
      throw lockError;
    }

    if (renewalError) {
      logger.warn(
        '[diagrams] Document lock renewal failed, but the lease was released successfully:',
        renewalError?.message || renewalError,
      );
    }
    if (releaseError) {
      logger.warn(
        '[diagrams] Document lock release failed after a successful operation:',
        releaseError?.message || releaseError,
      );
    }
    return result;
  }

  return {
    async read(name) {
      try {
        const resp = await blob(name).download();
        const text = await streamToString(resp.readableStreamBody);
        return { value: JSON.parse(text), etag: resp.etag };
      } catch (err) {
        if (Number(err?.statusCode) === 404) return null;
        throw err;
      }
    },
    create(name, value) {
      return upload(name, value, { ifNoneMatch: '*' });
    },
    replace(name, value, etag) {
      return upload(name, value, { ifMatch: etag });
    },
    put(name, value) {
      return upload(name, value, {});
    },
    async remove(name, etag) {
      const resp = await blob(name).deleteIfExists(
        etag ? { conditions: { ifMatch: etag } } : {},
      );
      return Boolean(resp.succeeded);
    },
    async *list(prefix) {
      for await (const item of container.listBlobsFlat({ prefix })) {
        yield item.name;
      }
    },
    withLock,
  };
}

module.exports = {
  createDiagramsRouter,
  createAzureBlobBackend,
  deriveOwnerKey,
  hashToken,
};
