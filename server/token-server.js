// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

/**
 * Minimal token server for Speech Service keyless auth.
 * Uses DefaultAzureCredential (az login in dev, managed identity in ACA)
 * to acquire a short-lived Speech STS token and returns it to the browser client.
 *
 * Runs on 127.0.0.1:3001 (not exposed externally — nginx proxies /api/).
 */

const express = require('express');
const { DefaultAzureCredential } = require('@azure/identity');
const { CosmosClient } = require('@azure/cosmos');
const { TableClient } = require('@azure/data-tables');
const { EmailClient, KnownEmailSendStatus } = require('@azure/communication-email');
const { createAccessControlRouter } = require('./access-control');
const { ArmKeyVaultAccessStore } = require('./arm-key-vault-access-store');
const { createOpenAIProxyRouter } = require('./openai-proxy');
const { asyncHandler, createErrorHandler } = require('./async-handler');
const crypto = require('crypto');

const app = express();
app.use((_req, res, next) => {
  res.set('X-Robots-Tag', 'noindex, nofollow, noarchive, nosnippet, noimageindex, noai, noimageai');
  res.set('Cache-Control', 'no-store');
  next();
});
// The Azure OpenAI proxy forwards vision requests that embed base64 images, so
// it needs a larger body limit. This route-scoped parser runs before the small
// global parser below; the global parser then skips bodies already parsed here.
app.use('/api/openai', express.json({ limit: '12mb' }));
app.use(express.json({ limit: '16kb' }));
const credential = new DefaultAzureCredential();

const REGION = process.env.AZURE_SPEECH_REGION;
const RESOURCE_ID = process.env.AZURE_SPEECH_RESOURCE_ID;

// ── Microsoft Entra ID email whitelist ─────────────────────────────────────
// Azure Container Apps Easy Auth injects trusted X-MS-CLIENT-PRINCIPAL-*
// headers. Nginx asks this router to authorize every protected request.
const ACCESS_CONTROL_ENABLED = process.env.ACCESS_CONTROL_ENABLED === 'true';
const ACCESS_ADMIN_EMAIL = process.env.ACCESS_ADMIN_EMAIL;
const ACCESS_TABLES_ENDPOINT = process.env.AZURE_TABLES_ACCESS_ENDPOINT;
const ACCESS_TABLE_NAME = process.env.AZURE_TABLES_ACCESS_TABLE || 'accesswhitelist';
const ACCESS_KEY_VAULT_RESOURCE_ID = process.env.AZURE_ACCESS_KEY_VAULT_RESOURCE_ID;
const PUBLIC_APP_URL = process.env.PUBLIC_URL;

let accessTable = null;
if (ACCESS_KEY_VAULT_RESOURCE_ID) {
  accessTable = new ArmKeyVaultAccessStore(ACCESS_KEY_VAULT_RESOURCE_ID, credential);
} else if (ACCESS_TABLES_ENDPOINT) {
  accessTable = new TableClient(ACCESS_TABLES_ENDPOINT, ACCESS_TABLE_NAME, credential);
}

if (ACCESS_CONTROL_ENABLED && (!ACCESS_ADMIN_EMAIL || !accessTable || !PUBLIC_APP_URL)) {
  console.error('[access] Access control is enabled but its administrator, store, or public URL is missing.');
}

app.use('/api/access', createAccessControlRouter({
  enabled: ACCESS_CONTROL_ENABLED,
  adminEmail: ACCESS_ADMIN_EMAIL,
  publicAppUrl: PUBLIC_APP_URL,
  table: accessTable,
}));

// ── Azure OpenAI proxy ─────────────────────────────────────────────────────
// Keeps Azure OpenAI credentials server-side so they are never shipped to the
// browser. Prefers keyless auth via DefaultAzureCredential (managed identity in
// ACA, `az login` in dev); falls back to AZURE_OPENAI_API_KEY when set.
const OPENAI_ENDPOINT = process.env.AZURE_OPENAI_ENDPOINT;
const OPENAI_API_KEY = process.env.AZURE_OPENAI_API_KEY; // optional fallback
const OPENAI_API_VERSION = process.env.AZURE_OPENAI_API_VERSION || '2024-05-01-preview';
const OPENAI_ALLOWED_DEPLOYMENTS = new Set(
  (process.env.AZURE_OPENAI_ALLOWED_DEPLOYMENTS || '')
    .split(',')
    .map((value) => value.trim())
    .filter(Boolean),
);

// Derive a rate-limit bucket that a caller cannot rotate at will.
//
// nginx forwards X-Forwarded-For as `$proxy_add_x_forwarded_for`, so the
// *first* entry is whatever the caller sent — using it would let anyone reset
// their own quota (and spoof another caller's) by varying the header. Azure
// Front Door owns X-Azure-ClientIP / X-Azure-SocketIP and overwrites any
// client-supplied value, and nginx rejects requests that do not carry this
// deployment's X-Azure-FDID, so those headers are trustworthy here. When they
// are absent the last X-Forwarded-For entry is used, because that one is
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

const consumeOpenAiRateLimit = createFixedWindowRateLimiter(
  60 * 60 * 1000,
  Math.max(1, Number(process.env.OPENAI_RATE_LIMIT_PER_HOUR) || 120),
);
const consumeUtilityApiRateLimit = createFixedWindowRateLimiter(60 * 60 * 1000, 120);
// Deliberately tight: the only client is an operator reading the feedback
// archive, so a low ceiling keeps the shared admin token from being brute-forced.
const consumeAdminApiRateLimit = createFixedWindowRateLimiter(60 * 60 * 1000, 30);

if (!OPENAI_ENDPOINT) {
  console.warn('[openai-proxy] AZURE_OPENAI_ENDPOINT is not set. /api/openai will return 503.');
}
if (OPENAI_ALLOWED_DEPLOYMENTS.size === 0) {
  console.warn('[openai-proxy] AZURE_OPENAI_ALLOWED_DEPLOYMENTS is empty. Any valid deployment name is accepted.');
}

// ── Durable feedback storage ───────────────────────────────────────────────
// Direct email delivery is preferred for low-cost deployments. Azure Table
// Storage and Cosmos DB remain supported for deployments that need an archive.
const FEEDBACK_EMAIL_ENDPOINT = process.env.FEEDBACK_EMAIL_ENDPOINT;
const FEEDBACK_EMAIL_SENDER = process.env.FEEDBACK_EMAIL_SENDER;
const FEEDBACK_EMAIL_RECIPIENT = process.env.FEEDBACK_EMAIL_RECIPIENT;
const TABLES_ENDPOINT = process.env.AZURE_TABLES_ENDPOINT;
const TABLES_FEEDBACK_TABLE = process.env.AZURE_TABLES_FEEDBACK_TABLE || 'feedback';
const COSMOS_ENDPOINT = process.env.AZURE_COSMOS_ENDPOINT;
const COSMOS_DATABASE_ID = process.env.COSMOS_DATABASE_ID || 'diagrams';
const COSMOS_FEEDBACK_CONTAINER_ID = process.env.COSMOS_FEEDBACK_CONTAINER_ID || 'feedback';

let feedbackEmailClient = null;
function getFeedbackEmailClient() {
  if (!FEEDBACK_EMAIL_ENDPOINT || !FEEDBACK_EMAIL_SENDER || !FEEDBACK_EMAIL_RECIPIENT) {
    return null;
  }
  if (!feedbackEmailClient) {
    feedbackEmailClient = new EmailClient(FEEDBACK_EMAIL_ENDPOINT, credential);
  }
  return feedbackEmailClient;
}

let feedbackTable = null;
function getFeedbackTable() {
  if (!TABLES_ENDPOINT) return null;
  if (!feedbackTable) {
    feedbackTable = new TableClient(TABLES_ENDPOINT, TABLES_FEEDBACK_TABLE, credential);
  }
  return feedbackTable;
}

// Lazily created singleton — reuse one CosmosClient for the process lifetime
// (Cosmos best practice; avoids per-request connection/auth overhead).
let feedbackContainer = null;
function getFeedbackContainer() {
  if (!COSMOS_ENDPOINT) return null;
  if (!feedbackContainer) {
    const client = new CosmosClient({ endpoint: COSMOS_ENDPOINT, aadCredentials: credential });
    feedbackContainer = client
      .database(COSMOS_DATABASE_ID)
      .container(COSMOS_FEEDBACK_CONTAINER_ID);
  }
  return feedbackContainer;
}

async function persistFeedback(item) {
  const emailClient = getFeedbackEmailClient();
  let emailError = null;
  if (emailClient) {
    try {
      const safeCategory = item.category.replace(/[\r\n]+/g, ' ').slice(0, 100);
      const message = {
        senderAddress: FEEDBACK_EMAIL_SENDER,
        content: {
          subject: `AzureDiagarm feedback: ${item.rating}/5 - ${safeCategory}`,
          plainText: [
            `Rating: ${item.rating}/5`,
            `Category: ${item.category}`,
            `Submitted: ${item.createdAt}`,
            '',
            'Comment:',
            item.comment || '(none)',
            '',
            'Context:',
            JSON.stringify(item.context, null, 2),
            '',
            `Feedback ID: ${item.id}`,
          ].join('\n'),
        },
        recipients: {
          to: [{ address: FEEDBACK_EMAIL_RECIPIENT }],
        },
      };
      const poller = await emailClient.beginSend(message);
      const result = await poller.pollUntilDone();
      if (result.status !== KnownEmailSendStatus.Succeeded) {
        throw new Error(`Feedback email delivery failed with status ${result.status}`);
      }
    } catch (error) {
      emailError = error;
    }
  }

  const table = getFeedbackTable();
  if (table) {
    try {
      const reverseTimestamp = String(253402300799999 - Date.now()).padStart(15, '0');
      await table.createEntity({
        partitionKey: 'feedback',
        rowKey: `${reverseTimestamp}-${item.id}`,
        id: item.id,
        rating: item.rating,
        category: item.category,
        comment: item.comment,
        contextJson: JSON.stringify(item.context),
        createdAt: item.createdAt,
      });
      return;
    } catch (error) {
      if (emailError) {
        throw new AggregateError([emailError, error], 'Feedback email and Table Storage delivery failed');
      }
      if (!emailClient) throw error;
      console.error('[feedback] archive error:', error.message);
      return;
    }
  }

  const container = getFeedbackContainer();
  if (container) {
    try {
      await container.items.create(item);
      return;
    } catch (error) {
      if (emailError) {
        throw new AggregateError([emailError, error], 'Feedback email and Cosmos DB delivery failed');
      }
      if (!emailClient) throw error;
      console.error('[feedback] archive error:', error.message);
      return;
    }
  }

  if (emailClient && !emailError) return;
  if (emailError) throw emailError;
  throw new Error('Feedback storage is not configured');
}

const FEEDBACK_RATE_LIMIT_WINDOW_MS = 60 * 60 * 1000;
const FEEDBACK_RATE_LIMIT_MAX = 10;
const consumeFeedbackRateLimit = createFixedWindowRateLimiter(
  FEEDBACK_RATE_LIMIT_WINDOW_MS,
  FEEDBACK_RATE_LIMIT_MAX,
);

async function readFeedback(limit) {
  const table = getFeedbackTable();
  if (table) {
    const items = [];
    const entities = table.listEntities({
      queryOptions: {
        filter: "PartitionKey eq 'feedback'",
        select: ['id', 'rating', 'category', 'comment', 'contextJson', 'createdAt'],
      },
    });

    for await (const entity of entities) {
      items.push({
        id: entity.id,
        rating: entity.rating,
        category: entity.category,
        comment: entity.comment,
        context: JSON.parse(entity.contextJson || '{}'),
        createdAt: entity.createdAt,
      });
      if (items.length >= limit) break;
    }
    return items;
  }

  const container = getFeedbackContainer();
  if (!container) return null;

  const { resources } = await container.items
    .query({
      query: 'SELECT TOP @limit c.id, c.rating, c.category, c.comment, c.context, c.createdAt FROM c WHERE c.type = @type ORDER BY c.createdAt DESC',
      parameters: [
        { name: '@limit', value: limit },
        { name: '@type', value: 'feedback' },
      ],
    })
    .fetchAll();
  return resources;
}

if (!REGION) {
  console.warn('[speech-token] AZURE_SPEECH_REGION is not set. Requests will fail.');
}
if (!RESOURCE_ID) {
  console.warn('[speech-token] AZURE_SPEECH_RESOURCE_ID is not set. Requests will fail.');
}

app.get('/api/speech-token', asyncHandler(async (req, res) => {
  const retryAfter = consumeUtilityApiRateLimit(req);
  if (retryAfter > 0) {
    res.set('Retry-After', String(retryAfter));
    return res.status(429).json({ error: 'Request limit exceeded. Please try again later.' });
  }
  if (!REGION || !RESOURCE_ID) {
    return res.status(503).json({ error: 'AZURE_SPEECH_REGION and AZURE_SPEECH_RESOURCE_ID must be configured' });
  }
  try {
    const { token: aadToken } = await credential.getToken(
      'https://cognitiveservices.azure.com/.default',
    );
    // JS Speech SDK requires the aad#{resourceId}#{aadToken} format for Entra ID auth
    res.json({ token: `aad#${RESOURCE_ID}#${aadToken}`, region: REGION });
  } catch (err) {
    console.error('[speech-token] error:', err.message);
    res.status(500).json({ error: 'Failed to acquire speech token' });
  }
}));

app.use('/api/openai', createOpenAIProxyRouter({
  endpoint: OPENAI_ENDPOINT,
  credential,
  apiKey: OPENAI_API_KEY,
  apiVersion: OPENAI_API_VERSION,
  allowedDeployments: OPENAI_ALLOWED_DEPLOYMENTS,
  consumeRateLimit: consumeOpenAiRateLimit,
}));

// ── Microsoft Learn docs grounding ─────────────────────────────────────────
// Server-side search of official Microsoft Learn docs via the public Learn MCP
// endpoint. Used to ground deployment-guide generation in current, citable
// documentation. Best-effort: failures return empty results so generation can
// proceed ungrounded.
const LEARN_MCP_URL = process.env.LEARN_MCP_URL || 'https://learn.microsoft.com/api/mcp';

async function searchLearnDocs(query, top) {
  const upstream = await fetch(LEARN_MCP_URL, {
    method: 'POST',
    headers: {
      'Content-Type': 'application/json',
      Accept: 'application/json, text/event-stream',
    },
    body: JSON.stringify({
      jsonrpc: '2.0',
      id: 1,
      method: 'tools/call',
      params: { name: 'microsoft_docs_search', arguments: { query } },
    }),
    signal: AbortSignal.timeout(15_000),
  });

  if (!upstream.ok) {
    throw new Error(`Learn MCP returned ${upstream.status}`);
  }

  // The endpoint replies with Server-Sent Events; find the data: line that
  // carries the tool result and unwrap result.content[].text (a JSON string).
  const body = await upstream.text();
  let payload = null;
  for (const line of body.split('\n')) {
    const trimmed = line.trim();
    if (!trimmed.startsWith('data:')) continue;
    const json = trimmed.slice(5).trim();
    if (!json || json === '[DONE]') continue;
    try {
      const obj = JSON.parse(json);
      if (obj.result && Array.isArray(obj.result.content)) {
        payload = obj;
        break;
      }
    } catch {
      /* ignore non-JSON / partial frames */
    }
  }
  if (!payload) return [];

  const textNode = payload.result.content.find((c) => c.type === 'text');
  if (!textNode) return [];

  const inner = JSON.parse(textNode.text);
  const results = Array.isArray(inner.results) ? inner.results : [];
  return results.slice(0, top).map((r) => ({
    title: String(r.title || '').slice(0, 200),
    url: String(r.contentUrl || ''),
    excerpt: typeof r.content === 'string' ? r.content.slice(0, 600) : '',
  }));
}

app.post('/api/docs-search', asyncHandler(async (req, res) => {
  const retryAfter = consumeUtilityApiRateLimit(req);
  if (retryAfter > 0) {
    res.set('Retry-After', String(retryAfter));
    return res.status(429).json({ error: 'Request limit exceeded. Please try again later.' });
  }
  const { query, top } = req.body || {};
  if (typeof query !== 'string' || query.trim().length === 0) {
    return res.status(400).json({ error: 'query is required' });
  }
  const limit = Math.min(Math.max(Number(top) || 6, 1), 10);
  try {
    const results = await searchLearnDocs(query.trim().slice(0, 400), limit);
    res.json({ results });
  } catch (err) {
    console.error('[docs-search] error:', err.message);
    // Soft-fail: grounding is best-effort.
    res.json({ results: [], error: 'docs search failed' });
  }
}));

app.get('/api/ice-token', asyncHandler(async (req, res) => {
  const retryAfter = consumeUtilityApiRateLimit(req);
  if (retryAfter > 0) {
    res.set('Retry-After', String(retryAfter));
    return res.status(429).json({ error: 'Request limit exceeded. Please try again later.' });
  }
  if (!REGION || !RESOURCE_ID) {
    return res.status(503).json({ error: 'AZURE_SPEECH_REGION and AZURE_SPEECH_RESOURCE_ID must be configured' });
  }
  try {
    const { token: aadToken } = await credential.getToken(
      'https://cognitiveservices.azure.com/.default',
    );
    // ICE relay endpoint also requires aad#resourceId#token format
    const authToken = `aad#${RESOURCE_ID}#${aadToken}`;
    const iceUrl = `https://${REGION}.tts.speech.microsoft.com/cognitiveservices/avatar/relay/token/v1`;
    const iceRes = await fetch(iceUrl, {
      signal: AbortSignal.timeout(15_000),
      headers: { Authorization: `Bearer ${authToken}` },
    });
    if (!iceRes.ok) {
      const body = await iceRes.text().catch(() => '');
      console.error(`[ice-token] error ${iceRes.status}: ${body}`);
      return res.status(502).json({ error: `ICE relay returned ${iceRes.status}` });
    }
    const data = await iceRes.json();
    res.json(data);
  } catch (err) {
    console.error('[ice-token] error:', err.message);
    res.status(500).json({ error: 'Failed to acquire ICE token' });
  }
}));

// ── Feedback (Cosmos DB) ──────────────────────────────────────────────────────
// ── Azure resource import (Resource Graph) ────────────────────────────────
// Lets an operator reverse-engineer a live Resource Group into a diagram by
// querying Azure Resource Graph server-side (via DefaultAzureCredential) and
// letting the client map the result deterministically. Resource Graph is
// Reader-sufficient and returns only real top-level resources.
//
// SECURITY: these routes let the *server identity* enumerate and export
// resources, so they are DISABLED by default and only enabled when
// AZURE_IMPORT_ENABLED=true. Leave unset on any shared/public deployment —
// the app's managed identity must never be exposed through /api/. Intended for
// local dev (`az login`) and single-tenant self-host.
const AZURE_IMPORT_ENABLED = String(process.env.AZURE_IMPORT_ENABLED || '').toLowerCase() === 'true';
const ARM_BASE = 'https://management.azure.com';
const GUID_RE = /^[0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{12}$/;
// Azure RG names: letters, digits, '.', '_', '-', '(', ')'; 1-90 chars; no trailing period.
const RG_NAME_RE = /^[A-Za-z0-9._()-]{1,90}$/;

async function armToken() {
  const { token } = await credential.getToken('https://management.azure.com/.default');
  return token;
}

// Guard applied to every /api/azure/* route.
function requireAzureImport(_req, res, next) {
  if (!AZURE_IMPORT_ENABLED) {
    return res.status(503).json({ error: 'Azure import is disabled. Set AZURE_IMPORT_ENABLED=true to enable (local / self-host only).' });
  }
  next();
}

app.get('/api/azure/subscriptions', requireAzureImport, asyncHandler(async (_req, res) => {
  try {
    const token = await armToken();
    const r = await fetch(`${ARM_BASE}/subscriptions?api-version=2022-12-01`, {
      headers: { Authorization: `Bearer ${token}` },
    });
    if (!r.ok) {
      const body = await r.text().catch(() => '');
      console.error(`[azure-import] subscriptions ${r.status}: ${body.slice(0, 300)}`);
      return res.status(502).json({ error: `Failed to list subscriptions (${r.status})` });
    }
    const data = await r.json();
    const subs = (data.value || []).map((s) => ({ subscriptionId: s.subscriptionId, displayName: s.displayName }));
    res.json({ subscriptions: subs });
  } catch (err) {
    console.error('[azure-import] subscriptions error:', err.message);
    res.status(500).json({ error: 'Failed to list subscriptions' });
  }
}));

app.get('/api/azure/resource-groups', requireAzureImport, asyncHandler(async (req, res) => {
  const subscriptionId = String(req.query.subscriptionId || '');
  if (!GUID_RE.test(subscriptionId)) {
    return res.status(400).json({ error: 'invalid subscriptionId' });
  }
  try {
    const token = await armToken();
    const r = await fetch(`${ARM_BASE}/subscriptions/${subscriptionId}/resourcegroups?api-version=2021-04-01`, {
      headers: { Authorization: `Bearer ${token}` },
    });
    if (!r.ok) {
      const body = await r.text().catch(() => '');
      console.error(`[azure-import] resource-groups ${r.status}: ${body.slice(0, 300)}`);
      return res.status(502).json({ error: `Failed to list resource groups (${r.status})` });
    }
    const data = await r.json();
    const groups = (data.value || []).map((g) => ({ name: g.name, location: g.location })).sort((a, b) => a.name.localeCompare(b.name));
    res.json({ resourceGroups: groups });
  } catch (err) {
    console.error('[azure-import] resource-groups error:', err.message);
    res.status(500).json({ error: 'Failed to list resource groups' });
  }
}));

app.post('/api/azure/resource-graph', requireAzureImport, asyncHandler(async (req, res) => {
  const { subscriptionId, resourceGroup } = req.body || {};
  if (!GUID_RE.test(String(subscriptionId || ''))) {
    return res.status(400).json({ error: 'invalid subscriptionId' });
  }
  if (!RG_NAME_RE.test(String(resourceGroup || '')) || String(resourceGroup).endsWith('.')) {
    return res.status(400).json({ error: 'invalid resourceGroup' });
  }
  try {
    const token = await armToken();
    // Reader-sufficient: returns top-level resources only (no ARM export noise).
    // resourceGroup is validated above; strip single quotes defensively so it
    // cannot break out of the KQL string literal.
    const rg = String(resourceGroup).replace(/'/g, '');
    const query = `Resources | where resourceGroup =~ '${rg}' | project id, name, type, kind, location, properties | limit 1000`;
    const r = await fetch(`${ARM_BASE}/providers/Microsoft.ResourceGraph/resources?api-version=2021-03-01`, {
      method: 'POST',
      headers: { Authorization: `Bearer ${token}`, 'Content-Type': 'application/json' },
      body: JSON.stringify({ subscriptions: [subscriptionId], query, options: { resultFormat: 'objectArray' } }),
    });
    if (!r.ok) {
      const body = await r.text().catch(() => '');
      console.error(`[azure-import] resource-graph ${r.status}: ${body.slice(0, 300)}`);
      return res.status(502).json({ error: `Resource Graph query failed (${r.status})` });
    }
    const data = await r.json();
    res.json({ resources: Array.isArray(data.data) ? data.data : [] });
  } catch (err) {
    console.error('[azure-import] resource-graph error:', err.message);
    res.status(500).json({ error: 'Resource Graph query failed' });
  }
}));

app.post('/api/feedback', asyncHandler(async (req, res) => {
  if (!getFeedbackEmailClient() && !getFeedbackTable() && !getFeedbackContainer()) {
    return res.status(503).json({ error: 'Feedback storage is not configured' });
  }

  const retryAfter = consumeFeedbackRateLimit(req);
  if (retryAfter > 0) {
    res.set('Retry-After', String(retryAfter));
    return res.status(429).json({ error: 'Too many feedback submissions. Please try again later.' });
  }

  const body = req.body || {};
  const rating = Number(body.rating);
  if (!Number.isInteger(rating) || rating < 1 || rating > 5) {
    return res.status(400).json({ error: 'rating must be an integer between 1 and 5' });
  }

  const category = typeof body.category === 'string' ? body.category.slice(0, 100) : 'General';
  const comment = typeof body.comment === 'string' ? body.comment.slice(0, 1000) : '';
  const ctx = body.context && typeof body.context === 'object' ? body.context : {};

  const item = {
    id: crypto.randomUUID(),
    type: 'feedback',
    rating,
    category,
    comment,
    context: {
      diagramName: typeof ctx.diagramName === 'string' ? ctx.diagramName.slice(0, 200) : '',
      serviceCount: Number.isFinite(Number(ctx.serviceCount)) ? Number(ctx.serviceCount) : 0,
      model: typeof ctx.model === 'string' ? ctx.model.slice(0, 100) : '',
      url: typeof ctx.url === 'string' ? ctx.url.slice(0, 500) : '',
      userAgent: typeof ctx.userAgent === 'string' ? ctx.userAgent.slice(0, 500) : '',
    },
    createdAt: new Date().toISOString(),
  };

  try {
    await persistFeedback(item);
    res.status(201).json({ ok: true, id: item.id });
  } catch (err) {
    console.error('[feedback] error:', err.message);
    res.status(500).json({ error: 'Failed to store feedback' });
  }
}));

// ── Admin: read persisted feedback (protected) ──────────────────────────────
// Lets an operator read verbatim comments from the configured archive. The
// route is disabled unless a dedicated token is configured and remains behind
// the application whitelist enforced by nginx.
const FEEDBACK_ADMIN_TOKEN = process.env.FEEDBACK_ADMIN_TOKEN || '';
const FEEDBACK_ADMIN_TOKEN_BYTES = Buffer.from(FEEDBACK_ADMIN_TOKEN, 'utf8');

// Compare raw bytes, not string lengths: two strings of equal character length
// can encode to buffers of different byte lengths (any non-ASCII input), and
// crypto.timingSafeEqual throws a RangeError in that case. Inside an async
// Express 4 handler that RangeError would surface as an unhandled rejection and
// take the whole container down, so the guard has to be byte-accurate.
function adminTokenMatches(presented) {
  if (typeof presented !== 'string' || presented.length === 0) return false;
  const presentedBytes = Buffer.from(presented, 'utf8');
  if (presentedBytes.length !== FEEDBACK_ADMIN_TOKEN_BYTES.length) return false;
  return crypto.timingSafeEqual(presentedBytes, FEEDBACK_ADMIN_TOKEN_BYTES);
}

app.get('/api/feedback/list', asyncHandler(async (req, res) => {
  if (!FEEDBACK_ADMIN_TOKEN) {
    return res.status(503).json({ error: 'Feedback admin endpoint is not configured' });
  }
  const retryAfter = consumeAdminApiRateLimit(req);
  if (retryAfter > 0) {
    res.set('Retry-After', String(retryAfter));
    return res.status(429).json({ error: 'Request limit exceeded. Please try again later.' });
  }
  const auth = req.get('authorization') || '';
  const presented = auth.startsWith('Bearer ') ? auth.slice(7) : (req.get('x-admin-token') || '');
  if (!adminTokenMatches(presented)) {
    return res.status(401).json({ error: 'unauthorized' });
  }

  if (!getFeedbackTable() && !getFeedbackContainer()) {
    return res.status(503).json({ error: 'Feedback archive is not configured' });
  }

  const limit = Math.min(Math.max(parseInt(req.query.limit, 10) || 200, 1), 1000);
  try {
    const items = await readFeedback(limit);
    res.json({ count: items.length, items });
  } catch (err) {
    console.error('[feedback/list] error:', err.message);
    res.status(500).json({ error: 'Failed to read feedback' });
  }
}));

// Final safety net: any error forwarded by asyncHandler is logged and answered
// with a generic 500 instead of escaping to the process and killing the
// container (start.sh stops the container when this server exits).
app.use(createErrorHandler(console));

const PORT = parseInt(process.env.TOKEN_SERVER_PORT || '3001', 10);
app.listen(PORT, '127.0.0.1', () => {
  console.log(`[speech-token] Listening on 127.0.0.1:${PORT}`);
});
