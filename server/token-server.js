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
const { EmailClient } = require('@azure/communication-email');
const {
  createAccessControlRouter,
  getAccessControlConfiguration,
  getPrincipal,
} = require('./access-control');
const { ArmKeyVaultAccessStore } = require('./arm-key-vault-access-store');
const { createOpenAIProxyRouter } = require('./openai-proxy');
const { runtimeAstraConfiguration } = require('./astra-policy');
const { createFixedWindowRateLimiter, createTableRateLimiter } = require('./rate-limiter');
const { createDiagramsRouter, createAzureBlobBackend } = require('./diagram-api');
const { asyncHandler, createErrorHandler } = require('./async-handler');
const { deploymentConfig, createOriginGuard } = require('./deployment-security');
const { MemoryBudgetStore, CosmosBudgetStore, TableBudgetStore, createBudgetManager, budgetIdentity } = require('./ai-budget');
const { createFeedbackService, createFeedbackRouter } = require('./feedback');
const { createArchivedFeedbackContact } = require('./feedback-configuration');
const { createGracefulShutdown } = require('./graceful-shutdown');
const { createReadinessHandler } = require('./readiness');

const deployment = deploymentConfig();

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

const accessRouter = createAccessControlRouter({
  enabled: ACCESS_CONTROL_ENABLED,
  adminEmail: ACCESS_ADMIN_EMAIL,
  publicAppUrl: PUBLIC_APP_URL,
  table: accessTable,
});
app.use('/api/access', accessRouter);
// Defense in depth: protected APIs enforce the same cached allowlist as nginx.
app.use('/api', accessRouter.requireAllowed, createOriginGuard(deployment));

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
const OPENAI_ENDPOINT = deployment.astra.endpoint;
const OPENAI_API_KEY = process.env.AZURE_OPENAI_API_KEY; // optional fallback
const ASTRA_DEPLOYMENT = deployment.astra.deployment;
const OPENAI_ALLOWED_DEPLOYMENTS = new Set(ASTRA_DEPLOYMENT ? [ASTRA_DEPLOYMENT] : []);

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
const consumeAdminApiRateLimit = createFixedWindowRateLimiter(60 * 60 * 1000, 30);

if (OPENAI_ALLOWED_DEPLOYMENTS.size === 0) {
  console.warn('[openai-proxy] GPT-6 Astra is unconfigured. Managed AI requests return 503.');
}
console.info(`[openai-proxy] Managed AI: GPT-6 Astra Responses only. BYO connections: ${deployment.allowByoAIEndpoints ? 'enabled' : 'disabled'}.`);

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

const FEEDBACK_RATE_LIMIT_WINDOW_MS = 60 * 60 * 1000;
const FEEDBACK_RATE_LIMIT_MAX = 10;
const consumeFeedbackRateLimit = createFixedWindowRateLimiter(
  FEEDBACK_RATE_LIMIT_WINDOW_MS,
  FEEDBACK_RATE_LIMIT_MAX,
);

const feedbackService = createFeedbackService({
  table: getFeedbackTable(), container: getFeedbackContainer(),
  emailClient: getFeedbackEmailClient(), emailSender: FEEDBACK_EMAIL_SENDER, emailRecipient: FEEDBACK_EMAIL_RECIPIENT,
  contactEnabled: FEEDBACK_CONTACT_ENABLED,
  archiveContact: item => createArchivedFeedbackContact(item.contact),
  retentionDays: deployment.retentionDays,
  legacyRetentionEnabled: deployment.legacyRetentionEnabled,
});
app.use('/api/feedback', createFeedbackRouter({
  service: feedbackService, mode: deployment.mode, adminEmail: ACCESS_ADMIN_EMAIL,
  consumeRateLimit: consumeFeedbackRateLimit, localAdminToken: process.env.FEEDBACK_ADMIN_TOKEN || '',
  consumeAdminRateLimit: consumeAdminApiRateLimit,
}));

let budgetStore;
if (deployment.store === 'cosmos') {
  const client = new CosmosClient({ endpoint: COSMOS_ENDPOINT, aadCredentials: credential });
  budgetStore = new CosmosBudgetStore(client.database(COSMOS_DATABASE_ID)
    .container(process.env.COSMOS_BUDGET_CONTAINER_ID || COSMOS_FEEDBACK_CONTAINER_ID), client);
} else if (deployment.store === 'table') {
  budgetStore = new TableBudgetStore(new TableClient(
    process.env.AZURE_TABLES_BUDGET_ENDPOINT || TABLES_ENDPOINT,
    process.env.AZURE_TABLES_BUDGET_TABLE || 'aibudgets', credential,
  ));
} else {
  budgetStore = new MemoryBudgetStore();
}
const aiBudget = createBudgetManager({
  store: budgetStore, dailyTokens: deployment.dailyTokens, concurrency: deployment.concurrency,
});
app.get('/api/ai/budget', asyncHandler(async (req, res) => {
  try {
    return res.json({ ...await aiBudget.status(budgetIdentity(req, deployment.mode)), mode: deployment.mode });
  } catch (error) {
    res.set('Retry-After', String(error.retryAfter || 5));
    return res.status(error.status || 503).json({ available: false, error: 'AI budget is unavailable.' });
  }
}));

if (!REGION) {
  console.warn('[speech-token] AZURE_SPEECH_REGION is not set. Requests will fail.');
}
if (!RESOURCE_ID) {
  console.warn('[speech-token] AZURE_SPEECH_RESOURCE_ID is not set. Requests will fail.');
}

// SECURITY: the browser must NEVER receive the container's managed-identity token.
// That token is issued for https://cognitiveservices.azure.com/.default, i.e. the whole
// Cognitive Services data plane, and is therefore interchangeable with the credential the
// Azure OpenAI proxy uses — replaying it would bypass the deployment allowlist, output-token
// clamping, `store: false` and rate limiting that the proxy exists to enforce.
// Instead we exchange it server-side for a Speech-only STS token (~10 min lifetime) via the
// resource's custom domain, and hand the client only that.
const SPEECH_STS_ENDPOINT = (() => {
  const explicit = process.env.AZURE_SPEECH_STS_ENDPOINT;
  if (explicit) return explicit.replace(/\/+$/, '');
  const account = typeof RESOURCE_ID === 'string' ? RESOURCE_ID.split('/').pop() : '';
  if (!account || !/^[A-Za-z0-9][A-Za-z0-9-]{1,62}$/.test(account)) return '';
  return `https://${account}.cognitiveservices.azure.com`;
})();

// STS tokens are valid for 10 minutes; refresh a minute early and share across callers.
const SPEECH_STS_TTL_MS = 9 * 60 * 1000;
let speechStsCache = null;
let speechStsInflight = null;

async function issueSpeechStsToken() {
  if (speechStsCache && speechStsCache.expiresAt > Date.now()) return speechStsCache.token;
  if (speechStsInflight) return speechStsInflight;

  speechStsInflight = (async () => {
    const { token: aadToken } = await credential.getToken('https://cognitiveservices.azure.com/.default');
    const response = await fetch(`${SPEECH_STS_ENDPOINT}/sts/v1.0/issueToken`, {
      method: 'POST',
      signal: AbortSignal.timeout(15_000),
      headers: { Authorization: `Bearer ${aadToken}`, 'Content-Length': '0' },
    });
    if (!response.ok) {
      const detail = await response.text().catch(() => '');
      throw new Error(`Speech STS exchange returned ${response.status}: ${detail.slice(0, 200)}`);
    }
    const token = (await response.text()).trim();
    if (!token) throw new Error('Speech STS exchange returned an empty token');
    speechStsCache = { token, expiresAt: Date.now() + SPEECH_STS_TTL_MS };
    return token;
  })().finally(() => {
    speechStsInflight = null;
  });

  return speechStsInflight;
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
  if (!SPEECH_STS_ENDPOINT) {
    console.error('[speech-token] cannot derive the Speech STS endpoint from AZURE_SPEECH_RESOURCE_ID');
    return res.status(503).json({ error: 'Speech token exchange is not configured' });
  }
  try {
    // Never fall back to returning the managed-identity token: failing closed is required.
    const token = await issueSpeechStsToken();
    res.json({ token, region: REGION });
  } catch (err) {
    console.error('[speech-token] error:', err.message);
    res.status(502).json({ error: 'Failed to acquire speech token' });
  }
}));

app.get('/api/runtime-config', (req, res) => {
  res.set('Cache-Control', 'no-store');
  res.json(runtimeAstraConfiguration(deployment.astra, deployment.allowByoAIEndpoints));
});

app.use('/api/openai', createOpenAIProxyRouter({
  endpoint: OPENAI_ENDPOINT,
  astraDeployment: ASTRA_DEPLOYMENT,
  credential,
  apiKey: OPENAI_API_KEY,
  allowedDeployments: OPENAI_ALLOWED_DEPLOYMENTS,
  allowByoAIEndpoints: deployment.allowByoAIEndpoints,
  consumeRateLimit: consumeOpenAiRateLimit,
  budget: aiBudget,
  mode: deployment.mode,
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

// Final safety net: any error forwarded by asyncHandler is logged and answered
// with a generic 500 instead of escaping to the process and killing the
// container (start.sh stops the container when this server exits).
app.use(createErrorHandler(console));

const PORT = parseInt(process.env.TOKEN_SERVER_PORT || '3001', 10);
async function start() {
  // Storage/TTL misconfiguration must fail before the public listener is ready.
  await budgetStore.validate?.();
  await feedbackService.validate();
  async function retain() {
    await feedbackService.sweep();
    await budgetStore.sweep?.();
  }
  try { await retain(); }
  catch (error) { console.error('[retention] Initial cleanup failed; will retry next interval:', error.name); }
  let retaining = false;
  const retentionTimer = setInterval(async () => {
    if (retaining) return;
    retaining = true;
    try { await retain(); }
    catch (error) { console.error('[retention] Cleanup failed; will retry next interval:', error.name); }
    finally { retaining = false; }
  }, 15 * 60 * 1000).unref();
  const server = app.listen(PORT, '127.0.0.1', () => {
    console.log(`[server] Listening on 127.0.0.1:${server.address().port}; deployment mode: ${deployment.mode}`);
  });
  const shutdown = createGracefulShutdown(server, { logger: console, timeoutMs: 25_000 });
  for (const signal of ['SIGTERM', 'SIGINT']) {
    process.once(signal, () => {
      shuttingDown = true;
      clearInterval(retentionTimer);
      shutdown(signal);
    });
  }
}
start().catch(error => {
  console.error('[server] Startup validation failed:', error.message);
  process.exitCode = 1;
});
