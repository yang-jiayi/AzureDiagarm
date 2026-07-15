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

// Deployment names are user-selected on the client; constrain them so they
// cannot be used to inject a different upstream path (SSRF / path traversal).
const DEPLOYMENT_NAME_RE = /^[A-Za-z0-9._-]{1,64}$/;

function getClientKey(req) {
  const forwarded = req.get('x-forwarded-for') || '';
  return (
    req.get('x-azure-clientip')
    || forwarded.split(',')[0]
    || req.ip
    || 'unknown'
  ).trim().slice(0, 128);
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

function buildOpenAIUrl(deployment, apiFormat) {
  const base = OPENAI_ENDPOINT.endsWith('/') ? OPENAI_ENDPOINT : `${OPENAI_ENDPOINT}/`;
  if (apiFormat === 'chat-completions') {
    return `${base}openai/deployments/${encodeURIComponent(deployment)}/chat/completions?api-version=${OPENAI_API_VERSION}`;
  }
  return `${base}openai/v1/responses`;
}

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
const feedbackRateLimits = new Map();

function getFeedbackClientKey(req) {
  const forwarded = req.get('x-forwarded-for') || '';
  return (
    req.get('x-azure-clientip')
    || forwarded.split(',')[0]
    || req.ip
    || 'unknown'
  ).trim().slice(0, 128);
}

function consumeFeedbackRateLimit(req) {
  const now = Date.now();
  for (const [key, value] of feedbackRateLimits) {
    if (value.resetAt <= now) feedbackRateLimits.delete(key);
  }

  const key = getFeedbackClientKey(req);
  const current = feedbackRateLimits.get(key);
  if (!current || current.resetAt <= now) {
    feedbackRateLimits.set(key, { count: 1, resetAt: now + FEEDBACK_RATE_LIMIT_WINDOW_MS });
    return 0;
  }
  if (current.count >= FEEDBACK_RATE_LIMIT_MAX) {
    return Math.max(1, Math.ceil((current.resetAt - now) / 1000));
  }
  current.count += 1;
  return 0;
}

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

app.get('/api/speech-token', async (req, res) => {
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
});

app.post('/api/openai', async (req, res) => {
  if (!OPENAI_ENDPOINT) {
    return res.status(503).json({ error: 'AZURE_OPENAI_ENDPOINT is not configured on the server' });
  }

  const { apiFormat, deployment, body } = req.body || {};
  if (apiFormat !== 'responses' && apiFormat !== 'chat-completions') {
    return res.status(400).json({ error: "apiFormat must be 'responses' or 'chat-completions'" });
  }
  if (typeof deployment !== 'string' || !DEPLOYMENT_NAME_RE.test(deployment)) {
    return res.status(400).json({ error: 'invalid deployment name' });
  }
  if (OPENAI_ALLOWED_DEPLOYMENTS.size > 0 && !OPENAI_ALLOWED_DEPLOYMENTS.has(deployment)) {
    return res.status(403).json({ error: 'deployment is not allowed' });
  }
  if (!body || typeof body !== 'object') {
    return res.status(400).json({ error: 'missing request body' });
  }
  const retryAfter = consumeOpenAiRateLimit(req);
  if (retryAfter > 0) {
    res.set('Retry-After', String(retryAfter));
    return res.status(429).json({ error: 'OpenAI request limit exceeded. Please try again later.' });
  }

  const upstreamBody = { ...body };
  if (apiFormat === 'responses') {
    upstreamBody.model = deployment;
    upstreamBody.store = false;
    upstreamBody.max_output_tokens = Math.min(
      Math.max(Number(upstreamBody.max_output_tokens) || 1, 1),
      32768,
    );
  } else {
    upstreamBody.max_tokens = Math.min(
      Math.max(Number(upstreamBody.max_tokens) || 1, 1),
      32768,
    );
  }

  try {
    const headers = { 'Content-Type': 'application/json' };
    if (OPENAI_API_KEY) {
      headers['api-key'] = OPENAI_API_KEY;
    } else {
      // Keyless: data-plane scope for Azure OpenAI / Cognitive Services.
      const { token } = await credential.getToken('https://cognitiveservices.azure.com/.default');
      headers['Authorization'] = `Bearer ${token}`;
    }

    const upstream = await fetch(buildOpenAIUrl(deployment, apiFormat), {
      method: 'POST',
      headers,
      body: JSON.stringify(upstreamBody),
      signal: AbortSignal.timeout(295_000),
    });

    const text = await upstream.text();
    res.status(upstream.status);
    res.set('Content-Type', upstream.headers.get('content-type') || 'application/json');
    res.send(text);
  } catch (err) {
    console.error('[openai-proxy] error:', err.message);
    res.status(502).json({ error: 'Azure OpenAI proxy request failed' });
  }
});

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

app.post('/api/docs-search', async (req, res) => {
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
});

app.get('/api/ice-token', async (req, res) => {
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
});

// ── Feedback (Cosmos DB) ──────────────────────────────────────────────────────
app.post('/api/feedback', async (req, res) => {
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
});

// ── Admin: read persisted feedback (protected) ──────────────────────────────
// Lets an operator read verbatim comments from Cosmos server-side — the app
// reaches Cosmos via the private endpoint, so this works even though the
// account has public network access disabled. Protected by a shared admin
// token (FEEDBACK_ADMIN_TOKEN); the endpoint is disabled (503) when unset.
const FEEDBACK_ADMIN_TOKEN = process.env.FEEDBACK_ADMIN_TOKEN || '';

function adminTokenMatches(presented) {
  if (!presented || presented.length !== FEEDBACK_ADMIN_TOKEN.length) return false;
  return crypto.timingSafeEqual(Buffer.from(presented), Buffer.from(FEEDBACK_ADMIN_TOKEN));
}

app.get('/api/feedback/list', async (req, res) => {
  if (!FEEDBACK_ADMIN_TOKEN) {
    return res.status(503).json({ error: 'Feedback admin endpoint is not configured' });
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
});

const PORT = parseInt(process.env.TOKEN_SERVER_PORT || '3001', 10);
app.listen(PORT, '127.0.0.1', () => {
  console.log(`[speech-token] Listening on 127.0.0.1:${PORT}`);
});
