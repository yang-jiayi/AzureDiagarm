// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.
const crypto = require('crypto');
const express = require('express');
const { asyncHandler } = require('./async-handler');
const { getPrincipal, normalizeEmail } = require('./access-control');
const { createArchivedFeedbackContact } = require('./feedback-configuration');

const errorStatus = error => Number(error?.statusCode || error?.status || error?.code);
const ownerKey = principal => principal
  ? crypto.createHash('sha256').update(`feedback:${principal.id}`).digest('hex') : null;

function safeContext(body) {
  if (body.includeMetadata !== true) return {};
  const ctx = body.context || {};
  const context = {};
  if (Number.isSafeInteger(ctx.serviceCount) && ctx.serviceCount >= 0) context.serviceCount = Math.min(ctx.serviceCount, 10_000);
  if (typeof ctx.model === 'string') context.model = ctx.model.replace(/[^a-zA-Z0-9._-]/g, '').slice(0, 64);
  try {
    const url = new URL(ctx.url);
    if (['http:', 'https:'].includes(url.protocol)) context.url = url.origin;
  } catch { /* Invalid URLs are omitted, never retained verbatim. */ }
  // Diagram names, user agents, arbitrary identifiers and URL paths are not collected.
  return context;
}

function createFeedbackService({
  table, container, emailClient, emailSender, emailRecipient, contactEnabled = false,
  retentionDays = 30, legacyRetentionEnabled = false, logger = console,
  archiveContact = item => createArchivedFeedbackContact(item.contact),
}) {
  const ttl = retentionDays * 86400;
  let containerValidation;
  let tableValidation;
  async function validateContainer() {
    if (!container) return;
    containerValidation ||= container.read().then(({ resource }) => {
      if (resource?.partitionKey?.paths?.join() !== '/id' || !resource.defaultTtl) {
        throw new Error('Feedback Cosmos container requires partition /id and enabled TTL.');
      }
    }).catch(error => { containerValidation = null; throw error; });
    return containerValidation;
  }
  async function validateTable() {
    if (!table?.createTable) return;
    tableValidation ||= table.createTable().catch(error => {
      if (errorStatus(error) === 409) return;
      tableValidation = null;
      throw error;
    });
    return tableValidation;
  }
  async function validate() {
    if (table) {
      try { await validateTable(); return; }
      catch (error) { if (!container) throw error; }
    }
    await validateContainer();
  }
  async function remove(item) {
    if (table && item.rowKey) {
      try { await table.deleteEntity('feedback', item.rowKey, { etag: item.etag }); }
      catch (error) { if (errorStatus(error) !== 404) throw error; }
    } else if (container) {
      try { await container.item(item.id, item.id).delete(); }
      catch (error) { if (errorStatus(error) !== 404) throw error; }
    }
  }
  async function find(id) {
    let tableError;
    if (table) {
      try {
        for await (const item of table.listEntities({
          queryOptions: { filter: `PartitionKey eq 'feedback' and id eq '${id}'` },
        })) return item;
      } catch (error) { tableError = error; }
    }
    if (container) {
      try {
        const { resource } = await container.item(id, id).read();
        if (resource?.type === 'feedback') return resource;
      } catch (error) { if (errorStatus(error) !== 404) throw error; }
    }
    if (tableError) throw tableError;
    return null;
  }
  const expires = item => Math.min(
    Date.parse(item.expiresAt || '') || Infinity,
    legacyRetentionEnabled ? (Date.parse(item.createdAt || '') + ttl * 1000 || Infinity) : Infinity,
  );
  async function sweep(now = Date.now()) {
    if (table) {
      const filter = legacyRetentionEnabled
        ? "PartitionKey eq 'feedback'"
        : `PartitionKey eq 'feedback' and expiresAt le '${new Date(now).toISOString()}'`;
      for await (const item of table.listEntities({ queryOptions: { filter } })) {
        if (expires(item) <= now || (legacyRetentionEnabled && !Number.isFinite(expires(item)))) await remove(item);
      }
    }
    if (container && legacyRetentionEnabled) {
      // Cosmos TTL handles new records. Retiring pre-TTL records requires an
      // explicit rollout decision after reviewing the existing archive.
      const iterator = container.items.query({
        query: 'SELECT * FROM c WHERE c.type = @type AND (c.createdAt < @cutoff OR NOT IS_DEFINED(c.createdAt))',
        parameters: [
          { name: '@type', value: 'feedback' },
          { name: '@cutoff', value: new Date(now - ttl * 1000).toISOString() },
        ],
      });
      for await (const page of iterator.getAsyncIterator()) {
        for (const item of page.resources) await remove(item);
      }
    }
  }
  async function saveArchive(item) {
    const archived = { ...item, contact: archiveContact(item) };
    let tableError;
    if (table) {
      try {
        await validateTable();
        await table.createEntity({
          partitionKey: 'feedback',
          rowKey: `${String(253402300799999 - Date.now()).padStart(15, '0')}-${item.id}`,
          id: item.id, rating: item.rating, category: item.category, comment: item.comment,
          contactJson: JSON.stringify(archived.contact), contextJson: JSON.stringify(item.context),
          createdAt: item.createdAt, expiresAt: item.expiresAt, ownerHash: item.ownerHash || '',
        });
        return;
      } catch (error) {
        tableError = error;
        if (!container) throw error;
        logger.warn('[feedback] Table archive unavailable; trying Cosmos DB.');
      }
    }
    if (container) {
      try {
        await validateContainer();
        await container.items.create(archived);
      } catch (error) {
        if (tableError) throw new AggregateError([tableError, error], 'Feedback archives are unavailable.');
        throw error;
      }
    }
  }
  async function sendEmail(item) {
    const poller = await emailClient.beginSend({
      senderAddress: emailSender,
      content: {
        subject: `AzureDiagarm feedback: ${item.rating}/5 - ${item.category.replace(/[\r\n]+/g, ' ')}`,
        plainText: [
          `Rating: ${item.rating}/5`, `Category: ${item.category}`, `Submitted: ${item.createdAt}`,
          ...(item.contact?.consent ? [
            `Follow-up contact: ${item.contact.email}`,
            `Contact consent expires: ${item.contact.expiresAt}`,
          ] : []),
          '', 'Comment:', item.comment || '(none)', '', 'Optional context:', JSON.stringify(item.context),
          '', `Feedback ID: ${item.id}`,
          'Contact addresses are delivered by email only; the application archive retains consent metadata, not the address.',
          'Archive retention/deletion does not delete this email. Mailbox policies apply independently.',
        ].join('\n'),
      },
      recipients: { to: [{ address: emailRecipient }] },
    });
    const result = await poller.pollUntilDone();
    if (result.status !== 'Succeeded') throw new Error('Email delivery did not succeed.');
  }
  return {
    configured: Boolean(table || container || emailClient),
    archive: Boolean(table || container),
    emailEnabled: Boolean(emailClient),
    contactEnabled,
    retentionDays, legacyRetentionEnabled,
    validate, sweep, find, remove,
    async save(item) {
      item.expiresAt = new Date(Date.parse(item.createdAt) + ttl * 1000).toISOString();
      item.ttl = ttl;
      let emailDelivered = false;
      if (item.contact?.consent) {
        if (!contactEnabled || !emailClient) throw new Error('Follow-up contact delivery is unavailable.');
        await sendEmail(item);
        emailDelivered = true;
      }
      // Archive failures remain failures even if contact email was delivered.
      await saveArchive(item);
      if (emailClient && !emailDelivered) {
        try {
          await sendEmail(item);
          emailDelivered = true;
        } catch (error) {
          if (!table && !container) throw error;
          logger.warn('[feedback] Saved in archive; notification email failed.');
        }
      }
      return { emailDelivered };
    },
    async list(limit) {
      let tableError;
      if (table) {
        try {
          const items = [];
          for await (const entity of table.listEntities({ queryOptions: { filter: "PartitionKey eq 'feedback'" } })) {
            if (expires(entity) <= Date.now()) continue;
            items.push({
              id: entity.id, rating: entity.rating, category: entity.category, comment: entity.comment,
              contact: createArchivedFeedbackContact(JSON.parse(entity.contactJson || '{"consent":false}')),
              context: JSON.parse(entity.contextJson || '{}'), createdAt: entity.createdAt, expiresAt: entity.expiresAt,
            });
            if (items.length >= limit) break;
          }
          return items;
        } catch (error) {
          tableError = error;
          if (!container) throw error;
          logger.warn('[feedback] Table archive read unavailable; trying Cosmos DB.');
        }
      }
      try {
        const { resources } = await container.items.query({
          query: 'SELECT TOP @limit c.id, c.rating, c.category, c.comment, c.contact, c.context, c.createdAt, c.expiresAt FROM c WHERE c.type = @type AND (NOT IS_DEFINED(c.expiresAt) OR c.expiresAt > @now) ORDER BY c.createdAt DESC',
          parameters: [
            { name: '@limit', value: limit }, { name: '@type', value: 'feedback' },
            { name: '@now', value: new Date().toISOString() },
          ],
        }).fetchAll();
        return resources.filter(item => expires(item) > Date.now())
          .map(item => ({ ...item, contact: createArchivedFeedbackContact(item.contact) }));
      } catch (error) {
        if (tableError) throw new AggregateError([tableError, error], 'Feedback archives are unavailable.');
        throw error;
      }
    },
  };
}

function createFeedbackRouter({ service, adminEmail, mode, consumeRateLimit = () => 0, consumeAdminRateLimit = () => 0, logger = console, localAdminToken = '' }) {
  const router = express.Router();
  const isAdmin = req => {
    if (mode === 'public') return getPrincipal(req)?.email === normalizeEmail(adminEmail);
    const presented = Buffer.from(req.get('x-admin-token') || String(req.get('authorization') || '').replace(/^Bearer /, ''));
    const expected = Buffer.from(localAdminToken);
    return expected.length > 0 && presented.length === expected.length && crypto.timingSafeEqual(presented, expected);
  };
  const adminRateLimit = asyncHandler(async (req, res, next) => {
    const retryAfter = await consumeAdminRateLimit(req);
    if (retryAfter) return res.set('Retry-After', String(retryAfter)).status(429).json({ error: 'Request limit exceeded. Try again later.' });
    next();
  });
  router.get('/policy', (_req, res) => res.json({
    archiveEnabled: service.archive, retentionDays: service.retentionDays,
    emailEnabled: service.emailEnabled, emailDeletionSupported: false,
    contactEnabled: service.contactEnabled && service.emailEnabled,
    contactRetentionDays: 180,
    legacyRetentionEnabled: service.legacyRetentionEnabled,
  }));
  router.post('/', asyncHandler(async (req, res) => {
    if (!service.configured) return res.status(503).json({ error: 'Feedback storage is not configured.' });
    const retryAfter = await consumeRateLimit(req);
    if (retryAfter) return res.set('Retry-After', String(retryAfter)).status(429).json({ error: 'Too many submissions. Try again later.' });
    const body = req.body || {};
    if (!Number.isInteger(body.rating) || body.rating < 1 || body.rating > 5) {
      return res.status(400).json({ error: 'rating must be an integer between 1 and 5' });
    }
    const contactConsent = body.contact?.consent === true;
    if (contactConsent && !service.contactEnabled) {
      return res.status(400).json({ error: 'Follow-up contact is not enabled.' });
    }
    if (contactConsent && !service.emailEnabled) {
      return res.status(503).json({ error: 'Follow-up contact delivery is not configured.' });
    }
    const contactEmail = contactConsent ? normalizeEmail(body.contact.email) : '';
    if (contactConsent && !contactEmail) {
      return res.status(400).json({ error: 'A valid email address is required when contact consent is enabled.' });
    }
    const createdAt = new Date().toISOString();
    const item = {
      id: crypto.randomUUID(), type: 'feedback', rating: body.rating,
      category: typeof body.category === 'string' ? body.category.slice(0, 100) : 'General',
      comment: typeof body.comment === 'string' ? body.comment.trim().slice(0, 1000) : '',
      context: safeContext(body), createdAt,
      contact: contactConsent ? {
        consent: true, email: contactEmail, consentAt: createdAt,
        expiresAt: new Date(Date.parse(createdAt) + 180 * 86400_000).toISOString(),
        followUpStatus: 'new',
      } : { consent: false },
      ownerHash: ownerKey(mode === 'public' ? getPrincipal(req) : null),
    };
    try {
      const result = await service.save(item);
      return res.status(201).json({
        ok: true, id: item.id, archiveSaved: service.archive,
        expiresAt: service.archive ? item.expiresAt : null, emailDelivered: result.emailDelivered,
        canDelete: service.archive && Boolean(item.ownerHash),
      });
    } catch (error) {
      logger.error('[feedback] save unavailable:', error.name);
      return res.status(503).json({ error: 'Feedback could not be saved.' });
    }
  }));
  router.get('/list', adminRateLimit, asyncHandler(async (req, res) => {
    if (!isAdmin(req)) return res.status(403).json({ error: 'Administrator access is required.' });
    if (!service.archive) return res.status(503).json({ error: 'Feedback archive is not configured.' });
    try {
      const items = await service.list(Math.min(Math.max(parseInt(req.query.limit, 10) || 200, 1), 1000));
      return res.json({ count: items.length, items });
    } catch (error) {
      logger.error('[feedback] list unavailable:', error.name);
      return res.status(503).json({ error: 'Feedback archive is unavailable.' });
    }
  }));
  router.delete('/:id', adminRateLimit, asyncHandler(async (req, res) => {
    if (!/^[0-9a-f-]{36}$/i.test(req.params.id)) return res.status(400).json({ error: 'Invalid feedback ID.' });
    if (!service.archive) return res.status(409).json({ error: 'Email-only feedback cannot be deleted here. Contact the mailbox administrator.' });
    try {
      const item = await service.find(req.params.id);
      const owner = mode === 'public' ? ownerKey(getPrincipal(req)) : null;
      if (!item || (!isAdmin(req) && (!owner || owner !== item.ownerHash))) {
        return res.status(404).json({ error: 'Feedback was not found or is not owned by this account.' });
      }
      await service.remove(item);
      return res.json({ archiveDeleted: true, emailDeleted: false });
    } catch (error) {
      logger.error('[feedback] deletion unavailable:', error.name);
      return res.status(503).json({ error: 'Feedback could not be deleted. Try again later.' });
    }
  }));
  return router;
}
module.exports = { safeContext, ownerKey, createFeedbackService, createFeedbackRouter };
