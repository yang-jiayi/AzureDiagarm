// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

const crypto = require('crypto');
const express = require('express');
const { asyncHandler } = require('./async-handler');

const ACCESS_PARTITION_KEY = 'allowed';
const EMAIL_PATTERN = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;

function normalizeEmail(value) {
  if (typeof value !== 'string') return '';
  const email = value.trim().toLowerCase();
  if (email.length === 0 || email.length > 254 || !EMAIL_PATTERN.test(email)) {
    return '';
  }
  return email;
}

function normalizePrincipalEmail(value) {
  if (typeof value !== 'string') return '';
  const email = value.trim().toLowerCase();
  const markerIndex = email.indexOf('#ext#@');
  if (markerIndex > 0) {
    const encodedEmail = email.slice(0, markerIndex);
    const separatorIndex = encodedEmail.lastIndexOf('_');
    if (separatorIndex > 0) {
      const guestEmail = normalizeEmail(
        `${encodedEmail.slice(0, separatorIndex)}@${encodedEmail.slice(separatorIndex + 1)}`,
      );
      if (guestEmail) return guestEmail;
    }
  }
  return normalizeEmail(email);
}

function normalizeOrigin(value) {
  if (typeof value !== 'string' || value.trim().length === 0) return '';
  try {
    return new URL(value).origin;
  } catch {
    return '';
  }
}

function rowKeyForEmail(email) {
  return crypto.createHash('sha256').update(email).digest('hex');
}

function decodeClientPrincipal(value) {
  if (typeof value !== 'string' || value.length === 0 || value.length > 32_768) {
    return null;
  }
  try {
    const decoded = JSON.parse(Buffer.from(value, 'base64').toString('utf8'));
    if (!decoded || !Array.isArray(decoded.claims)) return null;
    const claims = new Map();
    for (const claim of decoded.claims) {
      if (typeof claim?.typ !== 'string' || typeof claim?.val !== 'string') continue;
      claims.set(claim.typ.toLowerCase(), claim.val);
    }
    return { decoded, claims };
  } catch {
    return null;
  }
}

function firstClaim(claims, types) {
  for (const type of types) {
    const value = claims.get(type.toLowerCase());
    if (value) return value;
  }
  return '';
}

function getPrincipal(req) {
  const clientPrincipal = decodeClientPrincipal(req.get('x-ms-client-principal'));
  const claims = clientPrincipal?.claims || new Map();
  const email = normalizePrincipalEmail(req.get('x-ms-client-principal-name'))
    || normalizePrincipalEmail(firstClaim(claims, [
      'email',
      'emails',
      'preferred_username',
      'upn',
      'http://schemas.xmlsoap.org/ws/2005/05/identity/claims/emailaddress',
      'http://schemas.xmlsoap.org/ws/2005/05/identity/claims/upn',
      'name',
    ]));
  const decodedUserId = typeof clientPrincipal?.decoded?.userId === 'string'
    ? clientPrincipal.decoded.userId
    : '';
  const id = (
    req.get('x-ms-client-principal-id')
    || firstClaim(claims, [
      'oid',
      'http://schemas.microsoft.com/identity/claims/objectidentifier',
      'http://schemas.xmlsoap.org/ws/2005/05/identity/claims/nameidentifier',
      'sub',
    ])
    || decodedUserId
    || ''
  ).trim();
  if (!email || id.length === 0 || id.length > 128) return null;
  return { email, id };
}

function createAccessControlRouter(options = {}) {
  const router = express.Router();
  const enabled = options.enabled === true;
  const adminEmail = normalizeEmail(options.adminEmail);
  const publicOrigin = normalizeOrigin(options.publicAppUrl);
  const table = options.table || null;
  const logger = options.logger || console;
  const cacheTtlMs = Math.max(1_000, Number(options.cacheTtlMs) || 60_000);
  const maxUsers = Math.max(1, Number(options.maxUsers) || 500);
  const configured = !enabled || Boolean(adminEmail && publicOrigin && table);

  let cache = {
    expiresAt: 0,
    users: [],
    emails: new Set(),
  };
  let refreshPromise = null;

  function invalidateCache() {
    cache = { expiresAt: 0, users: [], emails: new Set() };
  }

  async function readUsers(force = false) {
    const now = Date.now();
    if (!force && cache.expiresAt > now) return cache;
    if (!force && refreshPromise) return refreshPromise;

    const refresh = (async () => {
      const users = [];
      const entities = table.listEntities({
        queryOptions: {
          filter: `PartitionKey eq '${ACCESS_PARTITION_KEY}'`,
          select: ['email', 'addedAt', 'addedBy'],
        },
      });

      for await (const entity of entities) {
        const email = normalizeEmail(entity.email);
        if (!email || email === adminEmail) continue;
        users.push({
          email,
          addedAt: typeof entity.addedAt === 'string' ? entity.addedAt : '',
          addedBy: normalizeEmail(entity.addedBy),
          isAdmin: false,
          immutable: false,
        });
      }

      users.sort((a, b) => a.email.localeCompare(b.email));
      const next = {
        expiresAt: Date.now() + cacheTtlMs,
        users,
        emails: new Set(users.map((user) => user.email)),
      };
      cache = next;
      return next;
    })();

    refreshPromise = refresh;
    try {
      return await refresh;
    } finally {
      if (refreshPromise === refresh) refreshPromise = null;
    }
  }

  async function accessAllowed(email) {
    if (email === adminEmail) return true;
    const current = await readUsers();
    return current.emails.has(email);
  }

  function configurationUnavailable(res) {
    return res.status(503).json({ error: 'Access control is not fully configured.' });
  }

  function requireSameOrigin(req, res, next) {
    const requestOrigin = normalizeOrigin(req.get('origin'));
    if (!requestOrigin || requestOrigin !== publicOrigin) {
      return res.status(403).json({ error: 'The request origin is not allowed.' });
    }
    next();
  }

  async function requireAdmin(req, res, next) {
    if (!enabled) return res.status(404).json({ error: 'Not found.' });
    if (!configured) return configurationUnavailable(res);
    const principal = getPrincipal(req);
    if (!principal) return res.status(401).json({ error: 'Authentication is required.' });
    if (principal.email !== adminEmail) {
      return res.status(403).json({ error: 'Administrator access is required.' });
    }
    req.accessPrincipal = principal;
    next();
  }

  router.get('/check', asyncHandler(async (req, res) => {
    if (!enabled) return res.status(204).end();
    if (!configured) return configurationUnavailable(res);

    const principal = getPrincipal(req);
    if (!principal) return res.status(401).end();

    try {
      if (await accessAllowed(principal.email)) return res.status(204).end();
      return res.status(403).end();
    } catch (error) {
      logger.error('[access/check] error:', error.message);
      return res.status(503).end();
    }
  }));

  router.get('/me', asyncHandler(async (req, res) => {
    if (!enabled) {
      return res.json({
        enabled: false,
        authenticated: false,
        email: null,
        isAdmin: false,
        allowed: true,
      });
    }
    if (!configured) return configurationUnavailable(res);

    const principal = getPrincipal(req);
    if (!principal) {
      return res.status(401).json({ error: 'Authentication is required.' });
    }

    try {
      const isAdmin = principal.email === adminEmail;
      const allowed = isAdmin || await accessAllowed(principal.email);
      if (!allowed) return res.status(403).json({ error: 'This account is not on the access list.' });
      return res.json({
        enabled: true,
        authenticated: true,
        email: principal.email,
        isAdmin,
        allowed: true,
      });
    } catch (error) {
      logger.error('[access/me] error:', error.message);
      return res.status(503).json({ error: 'Access control is temporarily unavailable.' });
    }
  }));

  router.get('/users', asyncHandler(requireAdmin), asyncHandler(async (_req, res) => {
    try {
      const current = await readUsers(true);
      return res.json({
        users: [
          {
            email: adminEmail,
            addedAt: '',
            addedBy: '',
            isAdmin: true,
            immutable: true,
          },
          ...current.users,
        ],
      });
    } catch (error) {
      logger.error('[access/users] list error:', error.message);
      return res.status(503).json({ error: 'Failed to read the access list.' });
    }
  }));

  router.post('/users', asyncHandler(requireAdmin), requireSameOrigin, asyncHandler(async (req, res) => {
    const email = normalizeEmail(req.body?.email);
    if (!email) return res.status(400).json({ error: 'A valid email address is required.' });
    if (email === adminEmail) {
      return res.status(409).json({ error: 'The administrator is always allowed.' });
    }

    try {
      const current = await readUsers(true);
      if (current.emails.has(email)) {
        return res.status(409).json({ error: 'This email address is already allowed.' });
      }
      if (current.users.length >= maxUsers) {
        return res.status(409).json({ error: 'The access list has reached its configured limit.' });
      }

      const now = new Date().toISOString();
      await table.createEntity({
        partitionKey: ACCESS_PARTITION_KEY,
        rowKey: rowKeyForEmail(email),
        email,
        addedAt: now,
        addedBy: req.accessPrincipal.email,
      });
      invalidateCache();
      logger.info(`[access/users] ${JSON.stringify(email)} added by ${JSON.stringify(req.accessPrincipal.email)}`);
      return res.status(201).json({
        user: {
          email,
          addedAt: now,
          addedBy: req.accessPrincipal.email,
          isAdmin: false,
          immutable: false,
        },
      });
    } catch (error) {
      if (error.statusCode === 409) {
        return res.status(409).json({ error: 'This email address is already allowed.' });
      }
      logger.error('[access/users] add error:', error.message);
      return res.status(503).json({ error: 'Failed to update the access list.' });
    }
  }));

  router.delete('/users', asyncHandler(requireAdmin), requireSameOrigin, asyncHandler(async (req, res) => {
    const email = normalizeEmail(req.body?.email);
    if (!email) return res.status(400).json({ error: 'A valid email address is required.' });
    if (email === adminEmail) {
      return res.status(400).json({ error: 'The administrator cannot be removed.' });
    }

    try {
      await table.deleteEntity(ACCESS_PARTITION_KEY, rowKeyForEmail(email));
      invalidateCache();
      logger.info(`[access/users] ${JSON.stringify(email)} removed by ${JSON.stringify(req.accessPrincipal.email)}`);
      return res.status(204).end();
    } catch (error) {
      if (error.statusCode === 404) {
        return res.status(404).json({ error: 'This email address is not on the access list.' });
      }
      logger.error('[access/users] delete error:', error.message);
      return res.status(503).json({ error: 'Failed to update the access list.' });
    }
  }));

  return router;
}

module.exports = {
  createAccessControlRouter,
  decodeClientPrincipal,
  normalizeEmail,
  normalizePrincipalEmail,
  rowKeyForEmail,
};
