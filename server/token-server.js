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
const {
  createAccessControlRouter,
  getAccessControlConfiguration,
  getPrincipal,
  normalizeEmail,
} = require('./access-control');
const { ArmKeyVaultAccessStore } = require('./arm-key-vault-access-store');
const { createOpenAIProxyRouter, logFoundryConfiguration } = require('./openai-proxy');
const { createFixedWindowRateLimiter, createTableRateLimiter } = require('./rate-limiter');
const { createDiagramsRouter, createAzureBlobBackend } = require('./diagram-api');
const { asyncHandler, createErrorHandler } = require('./async-handler');
const {
  createArchivedFeedbackContact,
  hasFeedbackArchiveConfiguration,
  hasFeedbackContactConfiguration,
  hasFeedbackDeliveryConfiguration,
} = require('./feedback-configuration');
const { createGracefulShutdown } = require('./graceful-shutdown');
const { createReadinessHandler } = require('./readiness');
const crypto = require('crypto');

const app = express();
app.disable('x-powered-by');
app.use((_req, res, next) => {
  res.set('X-Robots-Tag', 'noindex, nofollow, noarchive, nosnippet, noimageindex, noai, noimageai');
  res.set('Cache-Control', 'no-store');
  next();
});
// The Azure OpenAI proxy forwards vision requests that embed base64 images, so
// it needs a larger body limit. This route-scoped parser runs before the small
// global parser below; the global parser then skips bodies already parsed here.
app.use('/api/openai', express.json({ limit: '12mb' }));
// Diagram documents embed entire node/edge graphs (up to ~10MB), so the
// persistence API also needs a larger route-scoped parser ahead of the global
// small parser.
app.use('/api/diagrams', express.json({ limit: '12mb' }));
app.use(express.json({ limit: '16kb' }));
const credential = new DefaultAzureCredential();
let shuttingDown = false;

// Nginx, Azure Front Door, Docker, and Container Apps all probe this route.
// Keeping it on the Node process ensures a wedged or unavailable API process
// is not masked by nginx continuing to serve a static "ok" response.
app.get('/healthz', (_req, res) => {
  res.type('text/plain').send('ok\n');
});

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

const accessControlConfiguration = getAccessControlConfiguration({
  enabled: ACCESS_CONTROL_ENABLED,
  adminEmail: ACCESS_ADMIN_EMAIL,
  publicAppUrl: PUBLIC_APP_URL,
  table: accessTable,
});
if (!accessControlConfiguration.configured) {
  console.error(
    `[access] Readiness blocked by missing access-control configuration: ${accessControlConfiguration.missing.join(', ')}.`,
  );
}

app.get('/readyz', createReadinessHandler({
  isShuttingDown: () => shuttingDown,
  isConfigured: () => accessControlConfiguration.configured,
}));

app.use('/api/access', createAccessControlRouter({
  enabled: ACCESS_CONTROL_ENABLED,
  adminEmail: ACCESS_ADMIN_EMAIL,
  publicAppUrl: PUBLIC_APP_URL,
  table: accessTable,
}));

// ── Authenticated diagram persistence ───────────────────────────────────────
// Stores diagram documents, immutable versions, comments and share tokens in
// Azure Blob Storage using DefaultAzureCredential (no account keys / SAS). When
// AZURE_BLOB_ENDPOINT is unset the router mounts but returns 503 so the feature
// degrades cleanly rather than crashing the container.
const DIAGRAMS_BLOB_ENDPOINT = process.env.AZURE_BLOB_ENDPOINT;
const DIAGRAMS_CONTAINER = process.env.AZURE_BLOB_DIAGRAMS_CONTAINER || 'diagrams';

let diagramsBackend = null;
if (DIAGRAMS_BLOB_ENDPOINT) {
  diagramsBackend = createAzureBlobBackend({
    endpoint: DIAGRAMS_BLOB_ENDPOINT,
    containerName: DIAGRAMS_CONTAINER,
    credential,
  });
} else {
  console.warn('[diagrams] AZURE_BLOB_ENDPOINT is not set. /api/diagrams will return 503.');
}

app.use('/api/diagrams', createDiagramsRouter({
  backend: diagramsBackend,
  getPrincipal,
  publicUrl: PUBLIC_APP_URL,
  logger: console,
}));

// ── Azure OpenAI proxy ─────────────────────────────────────────────────────
// Keeps Azure OpenAI credentials server-side so they are never shipped to the
// browser. Prefers keyless auth via DefaultAzureCredential (managed identity in
// ACA, `az login` in dev); falls back to AZURE_OPENAI_API_KEY when set.
const OPENAI_ENDPOINT = process.env.AZURE_OPENAI_ENDPOINT;
const OPENAI_API_KEY = process.env.AZURE_OPENAI_API_KEY; // optional fallback
const OPENAI_API_VERSION = process.env.AZURE_OPENAI_API_VERSION || '2024-05-01-preview';
const FOUNDRY_ENDPOINT = process.env.AZURE_FOUNDRY_ENDPOINT;
const FOUNDRY_API_KEY = process.env.AZURE_FOUNDRY_API_KEY; // optional fallback
const OPENAI_ALLOWED_DEPLOYMENTS = new Set(
  (process.env.AZURE_OPENAI_ALLOWED_DEPLOYMENTS || '')
    .split(',')
    .map((value) => value.trim())
    .filter(Boolean),
);
const FOUNDRY_ALLOWED_DEPLOYMENTS = new Set(
  (process.env.AZURE_FOUNDRY_ALLOWED_DEPLOYMENTS || '')
    .split(',')
    .map((value) => value.trim())
    .filter(Boolean),
);

// ── Per-client rate limiting ───────────────────────────────────────────────
// The table-backed limiter shares counters across all Container Apps replicas
// via optimistic-concurrency writes to Azure Table Storage, giving an honest
// global limit.  When AZURE_TABLES_ENDPOINT is not configured the in-process
// limiter is used instead; rate limiting is then per-replica only, and the
// effective limit is maxRequests × <replica count>.
const OPENAI_RATE_LIMIT_PER_HOUR = Math.max(
  1, Number(process.env.OPENAI_RATE_LIMIT_PER_HOUR) || 120,
);
const RATE_LIMIT_TABLE_NAME = process.env.AZURE_TABLES_RATE_LIMIT_TABLE || 'ratelimit';
// AZURE_TABLES_ENDPOINT is read early here so we can decide which limiter to
// create; the same constant is re-declared (same value) later for the feedback
// table client.
const _RATE_LIMIT_TABLES_ENDPOINT = process.env.AZURE_TABLES_ENDPOINT;

let consumeOpenAiRateLimit;
if (_RATE_LIMIT_TABLES_ENDPOINT) {
  const rateLimitTableClient = new TableClient(
    _RATE_LIMIT_TABLES_ENDPOINT,
    RATE_LIMIT_TABLE_NAME,
    credential,
  );
  let rateLimitTableReady = null;
  const ensureRateLimitTable = () => {
    if (!rateLimitTableReady) {
      rateLimitTableReady = rateLimitTableClient.createTable()
        .catch((error) => {
          if (error.statusCode === 409) return;
          rateLimitTableReady = null;
          throw error;
        });
    }
    return rateLimitTableReady;
  };
  const consumeSharedOpenAiRateLimit = createTableRateLimiter(
    rateLimitTableClient,
    60 * 60 * 1000,
    OPENAI_RATE_LIMIT_PER_HOUR,
    {
      storageErrorRetryAfterSeconds: 5,
      onStorageError: (operation, error) => {
        console.error(
          `[openai-proxy] Shared rate-limit storage ${operation} failed; request rejected:`,
          error.message,
        );
      },
    },
  );
  consumeOpenAiRateLimit = async (req) => {
    try {
      await ensureRateLimitTable();
    } catch (error) {
      console.error('[openai-proxy] Unable to ensure the shared rate-limit table:', error.message);
      return 5;
    }
    return consumeSharedOpenAiRateLimit(req);
  };
  // Probe at startup, but reset the cached promise after a failure so a
  // transient outage can recover on a later request.
  ensureRateLimitTable().catch((error) => {
    console.error('[openai-proxy] Shared rate-limit table startup probe failed:', error.message);
  });
  console.info(
    `[openai-proxy] Using shared Table Storage rate limiter (table: ${RATE_LIMIT_TABLE_NAME}). `
    + 'Rate limit is globally enforced across all replicas.',
  );
} else {
  consumeOpenAiRateLimit = createFixedWindowRateLimiter(
    60 * 60 * 1000,
    OPENAI_RATE_LIMIT_PER_HOUR,
  );
  if (OPENAI_ENDPOINT) {
    console.warn(
      '[openai-proxy] AZURE_TABLES_ENDPOINT is not set. '
      + 'The OpenAI rate limiter is in-process only — with multiple Container Apps replicas '
      + `the effective per-IP limit is ${OPENAI_RATE_LIMIT_PER_HOUR} × <replica count>. `
      + 'Set AZURE_TABLES_ENDPOINT (and optionally AZURE_TABLES_RATE_LIMIT_TABLE) to enforce '
      + 'a global limit.',
    );
  }
}
const consumeUtilityApiRateLimit = createFixedWindowRateLimiter(60 * 60 * 1000, 120);
// Deliberately tight: the only client is an operator reading the feedback
// archive, so a low ceiling keeps the shared admin token from being brute-forced.
const consumeAdminApiRateLimit = createFixedWindowRateLimiter(60 * 60 * 1000, 30);

if (!OPENAI_ENDPOINT) {
  console.warn('[openai-proxy] AZURE_OPENAI_ENDPOINT is not set. /api/openai will return 503.');
}
if (OPENAI_ALLOWED_DEPLOYMENTS.size === 0) {
  console.warn('[openai-proxy] AZURE_OPENAI_ALLOWED_DEPLOYMENTS is empty. All Azure OpenAI requests will be rejected (503) until the allowlist is configured.');
}
logFoundryConfiguration(FOUNDRY_ENDPOINT, FOUNDRY_ALLOWED_DEPLOYMENTS, console);

// ── Durable feedback storage ───────────────────────────────────────────────
// Direct email delivery is preferred for low-cost deployments. Azure Table
// Storage and Cosmos DB remain supported for deployments that need an archive.
const FEEDBACK_EMAIL_ENDPOINT = process.env.FEEDBACK_EMAIL_ENDPOINT;
const FEEDBACK_EMAIL_SENDER = process.env.FEEDBACK_EMAIL_SENDER;
const FEEDBACK_EMAIL_RECIPIENT = process.env.FEEDBACK_EMAIL_RECIPIENT;
const FEEDBACK_CONTACT_ENABLED = process.env.FEEDBACK_CONTACT_ENABLED === 'true';
const TABLES_ENDPOINT = process.env.AZURE_TABLES_ENDPOINT;
const TABLES_FEEDBACK_TABLE = process.env.AZURE_TABLES_FEEDBACK_TABLE || 'feedback';
const COSMOS_ENDPOINT = process.env.AZURE_COSMOS_ENDPOINT;
const COSMOS_DATABASE_ID = process.env.COSMOS_DATABASE_ID || 'diagrams';
const COSMOS_FEEDBACK_CONTAINER_ID = process.env.COSMOS_FEEDBACK_CONTAINER_ID || 'feedback';
const feedbackConfiguration = {
  emailEndpoint: FEEDBACK_EMAIL_ENDPOINT,
  emailSender: FEEDBACK_EMAIL_SENDER,
  emailRecipient: FEEDBACK_EMAIL_RECIPIENT,
  contactEnabled: FEEDBACK_CONTACT_ENABLED,
  tablesEndpoint: TABLES_ENDPOINT,
  cosmosEndpoint: COSMOS_ENDPOINT,
};

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
let feedbackTableReady = null;
async function getFeedbackTable() {
  if (!TABLES_ENDPOINT) return null;
  if (!feedbackTable) {
    feedbackTable = new TableClient(TABLES_ENDPOINT, TABLES_FEEDBACK_TABLE, credential);
  }
  if (!feedbackTableReady) {
    feedbackTableReady = feedbackTable.createTable().catch((error) => {
      if (error.statusCode === 409) return;
      feedbackTableReady = null;
      throw error;
    });
  }
  await feedbackTableReady;
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
  const deliveryErrors = [];
  let emailDelivered = false;
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
            ...(item.contact?.consent ? [
              `Follow-up contact: ${item.contact.email}`,
              `Contact consent expires: ${item.contact.expiresAt}`,
            ] : []),
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
      emailDelivered = true;
    } catch (error) {
      deliveryErrors.push(error);
    }
  }

  // Contact addresses are delivered only through the configured email channel.
  // Archives retain consent metadata but never the address itself.
  if (item.contact?.consent && !emailDelivered) {
    throw deliveryErrors[0] || new Error('Follow-up contact delivery is unavailable');
  }

  const archiveItem = {
    ...item,
    contact: createArchivedFeedbackContact(item.contact),
  };

  if (TABLES_ENDPOINT) {
    try {
      const table = await getFeedbackTable();
      const reverseTimestamp = String(253402300799999 - Date.now()).padStart(15, '0');
      await table.createEntity({
        partitionKey: 'feedback',
        rowKey: `${reverseTimestamp}-${item.id}`,
        id: archiveItem.id,
        rating: archiveItem.rating,
        category: archiveItem.category,
        comment: archiveItem.comment,
        contactJson: JSON.stringify(archiveItem.contact),
        contextJson: JSON.stringify(archiveItem.context),
        createdAt: archiveItem.createdAt,
      });
      return;
    } catch (error) {
      deliveryErrors.push(error);
      if (emailDelivered) {
        console.error('[feedback] Table Storage archive error:', error.message);
        return;
      }
    }
  }

  const container = getFeedbackContainer();
  if (container) {
    try {
      await container.items.create(archiveItem);
      return;
    } catch (error) {
      deliveryErrors.push(error);
      if (emailDelivered) {
        console.error('[feedback] Cosmos DB archive error:', error.message);
        return;
      }
    }
  }

  if (emailDelivered) return;
  if (deliveryErrors.length === 1) throw deliveryErrors[0];
  if (deliveryErrors.length > 1) {
    throw new AggregateError(deliveryErrors, 'Feedback delivery failed in all configured channels');
  }
  throw new Error('Feedback storage is not configured');
}

const FEEDBACK_RATE_LIMIT_WINDOW_MS = 60 * 60 * 1000;
const FEEDBACK_RATE_LIMIT_MAX = 10;
const consumeFeedbackRateLimit = createFixedWindowRateLimiter(
  FEEDBACK_RATE_LIMIT_WINDOW_MS,
  FEEDBACK_RATE_LIMIT_MAX,
);

async function readFeedback(limit) {
  let tableError = null;
  if (TABLES_ENDPOINT) {
    try {
      const table = await getFeedbackTable();
      const items = [];
      const entities = table.listEntities({
        queryOptions: {
          filter: "PartitionKey eq 'feedback'",
          select: ['id', 'rating', 'category', 'comment', 'contactJson', 'contextJson', 'createdAt'],
        },
      });

      for await (const entity of entities) {
        items.push({
          id: entity.id,
          rating: entity.rating,
          category: entity.category,
          comment: entity.comment,
          contact: JSON.parse(entity.contactJson || '{"consent":false}'),
          context: JSON.parse(entity.contextJson || '{}'),
          createdAt: entity.createdAt,
        });
        if (items.length >= limit) break;
      }
      return items;
    } catch (error) {
      tableError = error;
      console.error('[feedback] Table Storage read failed; trying Cosmos DB fallback:', error.message);
    }
  }

  const container = getFeedbackContainer();
  if (!container) {
    if (tableError) throw tableError;
    return null;
  }

  try {
    const { resources } = await container.items
      .query({
        query: 'SELECT TOP @limit c.id, c.rating, c.category, c.comment, c.contact, c.context, c.createdAt FROM c WHERE c.type = @type ORDER BY c.createdAt DESC',
        parameters: [
          { name: '@limit', value: limit },
          { name: '@type', value: 'feedback' },
        ],
      })
      .fetchAll();
    return resources;
  } catch (error) {
    if (tableError) {
      throw new AggregateError([tableError, error], 'Feedback reads failed in Table Storage and Cosmos DB');
    }
    throw error;
  }
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
  foundryEndpoint: FOUNDRY_ENDPOINT,
  credential,
  apiKey: OPENAI_API_KEY,
  foundryApiKey: FOUNDRY_API_KEY,
  apiVersion: OPENAI_API_VERSION,
  allowedDeployments: OPENAI_ALLOWED_DEPLOYMENTS,
  allowedFoundryDeployments: FOUNDRY_ALLOWED_DEPLOYMENTS,
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
  if (!hasFeedbackDeliveryConfiguration(feedbackConfiguration)) {
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
  const contactConsent = body.contact?.consent === true;
  if (contactConsent && !FEEDBACK_CONTACT_ENABLED) {
    return res.status(400).json({ error: 'follow-up contact is not enabled' });
  }
  if (contactConsent && !hasFeedbackContactConfiguration(feedbackConfiguration)) {
    return res.status(503).json({ error: 'follow-up contact delivery is not configured' });
  }
  const contactEmail = contactConsent ? normalizeEmail(body.contact?.email) : '';
  if (contactConsent && !contactEmail) {
    return res.status(400).json({ error: 'a valid email address is required when contact consent is enabled' });
  }
  const ctx = body.context && typeof body.context === 'object' ? body.context : {};
  const createdAt = new Date();
  const contactExpiresAt = new Date(createdAt);
  contactExpiresAt.setUTCDate(contactExpiresAt.getUTCDate() + 180);

  const item = {
    id: crypto.randomUUID(),
    type: 'feedback',
    rating,
    category,
    comment,
    contact: contactConsent ? {
      consent: true,
      email: contactEmail,
      consentAt: createdAt.toISOString(),
      expiresAt: contactExpiresAt.toISOString(),
      followUpStatus: 'new',
    } : {
      consent: false,
    },
    context: {
      diagramName: typeof ctx.diagramName === 'string' ? ctx.diagramName.slice(0, 200) : '',
      serviceCount: Number.isFinite(Number(ctx.serviceCount)) ? Number(ctx.serviceCount) : 0,
      model: typeof ctx.model === 'string' ? ctx.model.slice(0, 100) : '',
      url: typeof ctx.url === 'string' ? ctx.url.slice(0, 500) : '',
      userAgent: typeof ctx.userAgent === 'string' ? ctx.userAgent.slice(0, 500) : '',
    },
    createdAt: createdAt.toISOString(),
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

  if (!hasFeedbackArchiveConfiguration(feedbackConfiguration)) {
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
const server = app.listen(PORT, '127.0.0.1', () => {
  console.log(`[speech-token] Listening on 127.0.0.1:${PORT}`);
});
const shutdown = createGracefulShutdown(server, { logger: console, timeoutMs: 25_000 });
for (const signal of ['SIGTERM', 'SIGINT']) {
  process.once(signal, () => {
    shuttingDown = true;
    shutdown(signal);
  });
}
