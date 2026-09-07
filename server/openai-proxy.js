// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

const crypto = require('crypto');
const express = require('express');
const { asyncHandler } = require('./async-handler');
const { budgetIdentity, reservationTokens, actualUsage, hasUnmeteredInput } = require('./ai-budget');
const { normalizeAzureOpenAIEndpoint: normalizeManagedEndpoint, normalizeHttpsOrigin, DEPLOYMENT_NAME_RE } = require('./astra-policy');

const DEFAULT_TIMEOUT_MS = 210_000;
const BYO_MODEL_NAME_RE = /^(?=.{1,128}$)(?=.*[A-Za-z0-9])[A-Za-z0-9._:-]+$/;
const API_KEY_RE = /^[\x21-\x7e]{8,512}$/;

function buildOpenAIUrl(endpoint) {
  const base = endpoint.endsWith('/') ? endpoint : `${endpoint}/`;
  return `${base}openai/v1/responses`;
}

function byoValidationError(code, message, status = 400) {
  return Object.assign(new Error(message), { code, status });
}

function resolveByoRequestConfig(rawConfig, allowByoAIEndpoints) {
  if (allowByoAIEndpoints !== true) {
    throw byoValidationError('byo_not_enabled', 'Bring-your-own AI connections are not enabled on this server.', 403);
  }
  if (!rawConfig || typeof rawConfig !== 'object' || Array.isArray(rawConfig)
    || Object.keys(rawConfig).some(key => !['provider', 'endpoint', 'apiKey'].includes(key))) {
    throw byoValidationError('invalid_byo_configuration', 'The custom AI configuration is invalid.');
  }
  const { provider, endpoint, apiKey } = rawConfig;
  if (provider !== 'azure-openai' && provider !== 'openai') {
    throw byoValidationError('invalid_byo_provider', "The custom AI provider must be 'azure-openai' or 'openai'.");
  }
  let origin;
  try {
    origin = provider === 'azure-openai'
      ? normalizeManagedEndpoint(endpoint)
      : normalizeHttpsOrigin(endpoint, hostname => hostname === 'api.openai.com');
  } catch {
    throw byoValidationError('invalid_byo_endpoint', 'Use a trusted Azure OpenAI resource origin or the official OpenAI HTTPS origin, without ports or API paths.');
  }
  if (typeof apiKey !== 'string' || !API_KEY_RE.test(apiKey)) {
    throw byoValidationError('invalid_byo_api_key', 'The custom AI API key is missing or invalid.');
  }
  return { provider, endpoint: origin, apiKey };
}

function buildByoAIUrl(config, apiFormat) {
  const path = apiFormat === 'responses' ? 'responses' : 'chat/completions';
  return `${config.endpoint}${config.provider === 'azure-openai' ? 'openai/' : ''}v1/${path}`;
}

function classifyByoUpstreamError(classified) {
  const replacements = {
    azure_openai_authentication_failed: ['byo_authentication_failed', 'The custom AI service rejected the supplied API key.'],
    azure_openai_rate_limited: ['byo_rate_limited', 'The custom AI service rate-limited the request.'],
    azure_openai_timeout: ['byo_timeout', 'The custom AI service timed out while processing the request.'],
    azure_openai_unavailable: ['byo_unavailable', 'The custom AI service is temporarily unavailable.'],
    azure_openai_non_json_error: ['byo_request_failed', 'The custom AI service returned an unexpected error format.'],
    azure_openai_request_failed: ['byo_request_failed', 'The custom AI service rejected the request.'],
  };
  const replacement = replacements[classified.code];
  return replacement ? { code: replacement[0], message: replacement[1] } : classified;
}

function isJsonMediaType(contentType) {
  const mediaType = String(contentType || '').split(';', 1)[0].trim().toLowerCase();
  return /^application\/(?:[a-z0-9!#$&^_.+-]+\+)?json$/.test(mediaType);
}

function parseUpstreamError(text) {
  if (!text) return { code: null, message: null };
  try {
    const payload = JSON.parse(text);
    const error = payload && typeof payload.error === 'object' ? payload.error : payload;
    const code = typeof error?.code === 'string'
      ? error.code
      : (typeof error?.type === 'string' ? error.type : null);
    const message = typeof error?.message === 'string' ? error.message : null;
    return { code, message };
  } catch {
    return { code: null, message: null };
  }
}

function classifyUpstreamError(status, contentType, upstreamCode, upstreamMessage) {
  const normalizedCode = String(upstreamCode || '').toLowerCase();
  const normalizedMessage = String(upstreamMessage || '').toLowerCase();
  if (status === 401 || status === 403) {
    return {
      code: 'azure_openai_authentication_failed',
      message: 'The upstream AI model service rejected the server credential.',
    };
  }
  if (status === 404 || normalizedCode.includes('deploymentnotfound')) {
    return {
      code: 'deployment_not_found',
      message: 'The selected model deployment was not found.',
    };
  }
  if (status === 429) {
    return {
      code: 'azure_openai_rate_limited',
      message: 'The upstream AI model service rate-limited the request.',
    };
  }
  if (status === 408 || status === 504) {
    return {
      code: 'azure_openai_timeout',
      message: 'The upstream AI model service timed out while processing the request.',
    };
  }
  if (status === 500 || status === 502 || status === 503) {
    return {
      code: 'azure_openai_unavailable',
      message: 'The upstream AI model service is temporarily unavailable.',
    };
  }
  if (status === 413) {
    return {
      code: 'request_too_large',
      message: 'The AI model request is too large.',
    };
  }
  if (
    normalizedCode.includes('contentfilter')
    || normalizedCode.includes('content_filter')
    || normalizedMessage.includes('content filter')
  ) {
    return {
      code: 'content_filtered',
      message: 'The upstream content filter rejected the request.',
    };
  }
  if (
    status === 400
    && (normalizedMessage.includes('image') || normalizedMessage.includes('vision'))
  ) {
    return {
      code: 'image_not_supported',
      message: 'The selected model deployment rejected the image input.',
    };
  }
  if (status === 400 || status === 422) {
    return {
      code: 'invalid_upstream_request',
      message: 'The upstream AI model service rejected the request format.',
    };
  }
  if (String(contentType || '').toLowerCase().includes('text/html')) {
    return {
      code: 'azure_openai_non_json_error',
      message: 'The upstream AI model service returned an unexpected non-JSON error.',
    };
  }
  return {
    code: 'azure_openai_request_failed',
    message: 'The upstream AI model service rejected the request.',
  };
}

function getHeader(headers, names, credentials = []) {
  for (const name of names) {
    const value = headers.get(name);
    if (typeof value === 'string'
      && /^(?:[a-f0-9]{32}|[a-f0-9]{8}-[a-f0-9]{4}-[a-f0-9]{4}-[a-f0-9]{4}-[a-f0-9]{12}|req_[a-z0-9]{8,96})$/i.test(value)
      && !credentials.some(secret => typeof secret === 'string' && secret && value.includes(secret))) return value;
  }
  return null;
}

function sendError(res, status, requestId, error) {
  return res.status(status).json({
    error: {
      source: error.source,
      code: error.code,
      message: error.message,
      requestId,
      ...(error.upstreamStatus ? { upstreamStatus: error.upstreamStatus } : {}),
      ...(error.upstreamRequestId ? { upstreamRequestId: error.upstreamRequestId } : {}),
    },
  });
}

function logEvent(logger, level, event) {
  const method = typeof logger?.[level] === 'function' ? logger[level] : logger?.log;
  if (typeof method === 'function') {
    method.call(logger, `[openai-proxy] ${JSON.stringify(event)}`);
  }
}

function createOpenAIProxyRouter(options) {
  const {
    endpoint,
    astraDeployment,
    credential,
    apiKey,
    allowedDeployments = new Set(),
    allowByoAIEndpoints = false,
    fetchImpl = globalThis.fetch,
    timeoutMs = DEFAULT_TIMEOUT_MS,
    consumeRateLimit = () => 0,
    budget,
    mode = 'local',
    logger = console,
  } = options;

  if (typeof fetchImpl !== 'function') {
    throw new TypeError('fetchImpl must be a function');
  }

  const router = express.Router();
  router.post('/', asyncHandler(async (req, res) => {
    const requestId = crypto.randomUUID();
    const startedAt = Date.now();
    res.set('X-AzureDiagarm-Request-Id', requestId);

    const envelope = req.body && typeof req.body === 'object' && !Array.isArray(req.body) ? req.body : {};
    const { apiFormat, deployment, body } = envelope;
    const hasByo = Object.hasOwn(envelope, 'byo');
    let byoConfig = null;
    if (hasByo) {
      try { byoConfig = resolveByoRequestConfig(envelope.byo, allowByoAIEndpoints); }
      catch (error) {
        return sendError(res, error.status, requestId, { source: 'proxy', code: error.code, message: error.message });
      }
    }
    if (apiFormat !== 'responses' && !(byoConfig && apiFormat === 'chat-completions')) {
      return sendError(res, 400, requestId, {
        source: 'proxy',
        code: 'invalid_api_format',
        message: byoConfig ? "apiFormat must be 'responses' or 'chat-completions'." : 'Managed GPT-6 Astra requires the Responses API.',
      });
    }

    if (['endpoint', 'apiKey', 'baseUrl', 'base_url', 'provider']
      .some(key => Object.hasOwn(envelope, key))) {
      return sendError(res, 403, requestId, {
        source: 'proxy', code: 'byo_not_enabled',
        message: 'Custom AI credentials and endpoints require a valid, explicitly enabled BYO configuration.',
      });
    }

    const managedConfigured = Boolean(endpoint || astraDeployment || allowedDeployments.size || apiKey);
    if ((!byoConfig || managedConfigured) && (!DEPLOYMENT_NAME_RE.test(astraDeployment || '')
      || allowedDeployments.size !== 1 || !allowedDeployments.has(astraDeployment))) {
      return sendError(res, 503, requestId, {
        source: 'proxy', code: 'astra_not_configured',
        message: 'Configure the explicit GPT-6 Astra deployment and its identical singleton allowlist.',
      });
    }
    if (byoConfig && (typeof deployment !== 'string' || !BYO_MODEL_NAME_RE.test(deployment))) {
      return sendError(res, 400, requestId, {
        source: 'proxy', code: 'invalid_deployment_name', message: 'A valid, explicit custom model or deployment ID is required.',
      });
    }
    if (!byoConfig && deployment !== astraDeployment) {
      return sendError(res, 403, requestId, {
        source: 'proxy', code: 'deployment_not_allowed',
        message: 'Only the configured GPT-6 Astra deployment is allowed.',
      });
    }
    if (!body || typeof body !== 'object' || Array.isArray(body)) {
      return sendError(res, 400, requestId, {
        source: 'proxy',
        code: 'missing_request_body',
        message: 'The Azure OpenAI request body is missing.',
      });
    }
    if (!Object.hasOwn(body, 'model') || body.model !== deployment
      || ['endpoint', 'apiKey', 'base_url', 'baseUrl', 'provider', 'byo'].some(key => Object.hasOwn(body, key))) {
      return sendError(res, 403, requestId, {
        source: 'proxy', code: 'deployment_not_allowed',
        message: 'deployment and body.model must explicitly match, without endpoint or credential overrides.',
      });
    }
    let managedEndpoint;
    try { if (managedConfigured) managedEndpoint = normalizeManagedEndpoint(endpoint); } catch {
      return sendError(res, 503, requestId, {
        source: 'proxy', code: 'proxy_not_configured', message: 'The managed Azure OpenAI endpoint is not configured correctly.',
      });
    }

    const upstreamSource = byoConfig ? 'byo' : 'azure_openai';
    const provider = byoConfig ? (byoConfig.provider === 'openai' ? 'byo_openai' : 'byo_azure_openai') : 'azure_openai';
    const loggedDeployment = byoConfig ? 'user-provided' : deployment;
    const wrongFormatKeys = apiFormat === 'responses'
      ? ['messages', 'max_tokens', 'max_completion_tokens', 'thinking', 'output_config', 'reasoning_effort', 'response_format']
      : ['input', 'max_output_tokens', 'reasoning', 'text', 'thinking', 'output_config'];
    if (wrongFormatKeys.some(key => Object.hasOwn(body, key))
      || (apiFormat === 'chat-completions' && Object.hasOwn(body, 'max_tokens') && Object.hasOwn(body, 'max_completion_tokens'))) {
      return sendError(res, 400, requestId, {
        source: 'proxy', code: 'invalid_api_format',
        message: 'The request body must match its explicit API format.',
      });
    }

    const upstreamBody = { ...body };
    // Complete responses only: stored inputs, remote tools and multiple
    // generations have unbounded or unobservable usage.
    if (body.background || body.previous_response_id || body.conversation
      || body.tools?.length || body.functions?.length || body.mcp_servers?.length || body.container
      || body.prompt || body.audio
      || (body.modalities !== undefined && (!Array.isArray(body.modalities) || body.modalities.some(value => value !== 'text')))
      || (body.n !== undefined && body.n !== 1) || body.best_of || hasUnmeteredInput(body)) {
      return sendError(res, 400, requestId, {
        source: 'proxy', code: 'unsupported_request_mode',
        message: 'Use a complete, non-streaming request with inline text/images, without stored inputs or remote tools.',
      });
    }
    upstreamBody.store = false;
    upstreamBody.stream = false;
    const outputField = apiFormat === 'responses' ? 'max_output_tokens'
      : (Object.hasOwn(body, 'max_completion_tokens') ? 'max_completion_tokens' : 'max_tokens');
    upstreamBody[outputField] = Math.floor(Math.min(
      Math.max(Number(upstreamBody[outputField]) || 1, 1), 32768,
    ));

    const retryAfter = await consumeRateLimit(req);
    if (retryAfter > 0) {
      res.set('Retry-After', String(retryAfter));
      logEvent(logger, 'warn', {
        event: 'proxy_rate_limit_exceeded', requestId, deployment: loggedDeployment,
        apiFormat, provider, status: 429, retryAfterSeconds: retryAfter,
        durationMs: Date.now() - startedAt,
      });
      return sendError(res, 429, requestId, {
        source: 'proxy',
        code: 'proxy_rate_limit_exceeded',
        message: 'The application OpenAI request limit was exceeded.',
      });
    }

    const headers = { 'Content-Type': 'application/json' };
    if (byoConfig) {
      if (byoConfig.provider === 'openai') headers.Authorization = `Bearer ${byoConfig.apiKey}`;
      else headers['api-key'] = byoConfig.apiKey;
    } else if (apiKey) {
      headers['api-key'] = apiKey;
    } else {
      try {
        const tokenResult = await credential?.getToken('https://cognitiveservices.azure.com/.default');
        if (!tokenResult?.token) throw new Error('Credential returned no token');
        headers.Authorization = `Bearer ${tokenResult.token}`;
      } catch (error) {
        logEvent(logger, 'error', {
          event: 'credential_acquisition_failed',
          requestId,
          deployment: loggedDeployment,
          apiFormat,
          provider,
        });
        return sendError(res, 502, requestId, {
          source: 'credential',
          code: 'credential_acquisition_failed',
          message: 'The server could not acquire an Azure OpenAI credential.',
        });
      }
    }

    let identity;
    let lease;
    let usage;
    let dispatched = false;
    const controller = new AbortController();
    const cancel = () => { if (!res.writableEnded) controller.abort(); };
    req.once('aborted', cancel);
    res.once('close', cancel);
    const deadline = setTimeout(() => controller.abort(new DOMException('Timed out', 'TimeoutError')), timeoutMs);
    deadline.unref?.();
    try {
      if (budget) {
        try {
          identity = budgetIdentity(req, mode);
          lease = await budget.reserve(identity, reservationTokens(upstreamBody, apiFormat));
        } catch (error) {
          const status = error.status || 503;
          const code = error.code || 'ai_budget_unavailable';
          const retryAfterSeconds = error.retryAfter || 5;
          res.set('Retry-After', String(retryAfterSeconds));
          logEvent(logger, status >= 500 ? 'error' : 'warn', {
            event: code, requestId, deployment: loggedDeployment, apiFormat, provider,
            status, retryAfterSeconds, durationMs: Date.now() - startedAt,
          });
          return sendError(res, status, requestId, {
            source: 'budget', code,
            message: error.status ? error.message : 'AI budget is unavailable. Try again shortly.',
          });
        }
      } else if (mode === 'public') {
        return sendError(res, 503, requestId, {
          source: 'budget', code: 'ai_budget_unavailable', message: 'AI budget is not configured.',
        });
      }
      if (res.destroyed) return;
      if (controller.signal.aborted) {
        return sendError(res, 504, requestId, {
          source: 'budget', code: 'ai_budget_timeout',
          message: 'The request timed out while reserving its AI budget. Try again shortly.',
        });
      }
    let upstream;
    try {
      const upstreamUrl = byoConfig ? buildByoAIUrl(byoConfig, apiFormat) : buildOpenAIUrl(managedEndpoint);
      dispatched = true;
      upstream = await fetchImpl(upstreamUrl, {
        method: 'POST',
        headers,
        body: JSON.stringify(upstreamBody),
        signal: controller.signal,
        redirect: 'error',
      });
    } catch (error) {
      const timedOut = error?.name === 'AbortError' || error?.name === 'TimeoutError';
      const status = timedOut ? 504 : 502;
      const code = byoConfig
        ? (timedOut ? 'byo_timeout' : 'byo_connection_failed')
        : (timedOut ? 'azure_openai_timeout' : 'azure_openai_connection_failed');
      logEvent(logger, 'error', {
        event: code,
        requestId,
        deployment: loggedDeployment,
        apiFormat,
        provider,
        durationMs: Date.now() - startedAt,
      });
      return sendError(res, status, requestId, {
        source: 'proxy_transport',
        code,
        message: byoConfig
          ? (timedOut ? 'The custom AI request timed out.' : 'The server could not connect to the custom AI service.')
          : (timedOut
              ? 'The Azure OpenAI request timed out.'
              : 'The server could not connect to Azure OpenAI.'),
      });
    }

    const contentType = upstream.headers.get('content-type') || '';
    // Custom services must not echo keys or endpoint details through diagnostic
    // headers. The application-generated request ID is always available.
    const upstreamRequestId = byoConfig ? null : getHeader(upstream.headers, [
      'apim-request-id',
      'x-ms-request-id',
      'x-request-id',
      'request-id',
      'trace-id',
    ], [apiKey, headers.Authorization?.slice('Bearer '.length)]);
    if (upstreamRequestId) {
      res.set('X-Upstream-Request-Id', upstreamRequestId);
    }
    let retryAfterHeader = upstream.headers.get('retry-after')?.trim();
    if (retryAfterHeader && /^\d{1,8}$/.test(retryAfterHeader)) {
      retryAfterHeader = String(Math.min(86400, Number(retryAfterHeader)));
    } else if (retryAfterHeader) {
      const date = Date.parse(retryAfterHeader);
      retryAfterHeader = Number.isFinite(date) ? String(Math.min(86400, Math.max(0, Math.ceil((date - Date.now()) / 1000)))) : null;
    }
    if (!retryAfterHeader) {
      for (const name of ['retry-after-ms', 'x-ms-retry-after-ms']) {
        const value = upstream.headers.get(name)?.trim();
        if (!value || !/^\d+(?:\.\d+)?$/.test(value)) continue;
        const milliseconds = Number(value);
        if (!Number.isFinite(milliseconds) || milliseconds > Number.MAX_SAFE_INTEGER) continue;
        retryAfterHeader = String(Math.min(86400, Math.ceil(milliseconds / 1000)));
        break;
      }
    }
    if (retryAfterHeader) {
      res.set('Retry-After', retryAfterHeader);
    }

    // Reading the upstream body can still fail after the response headers
    // arrive (connection reset mid-stream, upstream timeout). Left unhandled
    // this rejects the request promise and terminates the process.
    let text;
    try {
      text = await upstream.text();
    } catch (error) {
      logEvent(logger, 'error', {
        event: 'upstream_body_read_failed',
        requestId,
        deployment: loggedDeployment,
        apiFormat,
        provider,
        upstreamStatus: upstream.status,
        upstreamRequestId,
        durationMs: Date.now() - startedAt,
      });
      return sendError(res, 502, requestId, {
        source: 'proxy_transport',
        code: byoConfig ? 'byo_connection_failed' : 'azure_openai_connection_failed',
        message: 'The server could not read the AI response.',
        upstreamStatus: upstream.status,
        upstreamRequestId,
      });
    }
    if (!upstream.ok) {
      // A rejected request has no successful generation to charge.
      // 5xx/timeout outcomes remain reserved because usage may be unknown.
      if ([400, 401, 403, 404, 413, 422, 429].includes(upstream.status)) usage = 0;
      const { code: upstreamCode, message: upstreamMessage } = parseUpstreamError(text);
      const upstreamError = classifyUpstreamError(
        upstream.status,
        contentType,
        upstreamCode,
        upstreamMessage,
      );
      const classified = byoConfig ? classifyByoUpstreamError(upstreamError) : upstreamError;
      logEvent(logger, 'error', {
        event: classified.code,
        requestId,
        deployment: loggedDeployment,
        apiFormat,
        provider,
        upstreamStatus: upstream.status,
        upstreamRequestId,
        jsonContentType: isJsonMediaType(contentType),
        durationMs: Date.now() - startedAt,
      });
      return sendError(res, upstream.status, requestId, {
        source: upstreamSource,
        code: classified.code,
        message: classified.message,
        upstreamStatus: upstream.status,
        upstreamRequestId,
      });
    }

    if (!isJsonMediaType(contentType)) {
      logEvent(logger, 'error', {
        event: 'invalid_upstream_response',
        requestId,
        deployment: loggedDeployment,
        apiFormat,
        provider,
        upstreamStatus: upstream.status,
        upstreamRequestId,
        jsonContentType: isJsonMediaType(contentType),
        durationMs: Date.now() - startedAt,
      });
      return sendError(res, 502, requestId, {
        source: upstreamSource,
        code: 'invalid_upstream_response',
        message: 'The AI service returned an unexpected response format.',
        upstreamStatus: upstream.status,
        upstreamRequestId,
      });
    }

    try { usage = actualUsage(JSON.parse(text)); } catch { /* Keep reservation for unknown usage. */ }
    logEvent(logger, 'info', {
      event: 'request_succeeded',
      requestId,
      deployment: loggedDeployment,
      apiFormat,
      provider,
      upstreamStatus: upstream.status,
      upstreamRequestId,
      durationMs: Date.now() - startedAt,
    });
    res.status(upstream.status);
    res.set('Content-Type', 'application/json');
    return res.send(text);
    } finally {
      clearTimeout(deadline);
      req.off('aborted', cancel);
      res.off('close', cancel);
      if (lease) {
        try { await budget.settle(identity, lease, dispatched ? usage : 0); }
        catch (error) {
          // Do not refund uncertain usage. The expiring shared lease releases
          // concurrency even if storage is unavailable or this replica dies.
          logEvent(logger, 'error', { event: 'ai_budget_settlement_failed', requestId });
        }
      }
    }
  }));

  return router;
}

module.exports = {
  buildOpenAIUrl,
  buildByoAIUrl,
  resolveByoRequestConfig,
  classifyUpstreamError,
  createOpenAIProxyRouter,
  isJsonMediaType,
  parseUpstreamError,
};
