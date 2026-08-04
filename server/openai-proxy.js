// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

const crypto = require('crypto');
const express = require('express');
const { asyncHandler } = require('./async-handler');

const DEPLOYMENT_NAME_RE = /^[A-Za-z0-9._-]{1,64}$/;
const DEFAULT_API_VERSION = '2024-05-01-preview';
const DEFAULT_ANTHROPIC_VERSION = '2023-06-01';
const DEFAULT_TIMEOUT_MS = 210_000;

function logFoundryConfiguration(endpoint, allowedDeployments, logger = console) {
  const hasEndpoint = typeof endpoint === 'string' && endpoint.trim().length > 0;
  const hasDeployments = allowedDeployments instanceof Set && allowedDeployments.size > 0;
  if (!hasEndpoint && !hasDeployments) {
    logger.info(
      '[openai-proxy] Optional Microsoft Foundry provider is disabled; '
      + 'Anthropic models will not be offered.',
    );
    return;
  }
  if (!hasEndpoint) {
    logger.warn(
      '[openai-proxy] AZURE_FOUNDRY_ENDPOINT is not set. '
      + 'Anthropic requests will return 503.',
    );
  }
  if (!hasDeployments) {
    logger.warn(
      '[openai-proxy] AZURE_FOUNDRY_ALLOWED_DEPLOYMENTS is empty. '
      + 'Anthropic requests will be rejected.',
    );
  }
}

function buildOpenAIUrl(endpoint, deployment, apiFormat, apiVersion) {
  const base = endpoint.endsWith('/') ? endpoint : `${endpoint}/`;
  if (apiFormat === 'anthropic-messages') {
    return `${base}anthropic/v1/messages`;
  }
  if (apiFormat === 'chat-completions') {
    return `${base}openai/deployments/${encodeURIComponent(deployment)}/chat/completions?api-version=${apiVersion}`;
  }
  return `${base}openai/v1/responses`;
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

function getHeader(headers, names) {
  for (const name of names) {
    const value = headers.get(name);
    if (value) return value.slice(0, 256);
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
      ...(error.upstreamCode ? { upstreamCode: error.upstreamCode } : {}),
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
    foundryEndpoint,
    credential,
    apiKey,
    foundryApiKey,
    apiVersion = DEFAULT_API_VERSION,
    anthropicVersion = DEFAULT_ANTHROPIC_VERSION,
    allowedDeployments = new Set(),
    allowedFoundryDeployments = new Set(),
    fetchImpl = globalThis.fetch,
    timeoutMs = DEFAULT_TIMEOUT_MS,
    consumeRateLimit = () => 0,
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

    const { apiFormat, deployment, body } = req.body || {};
    if (
      apiFormat !== 'responses'
      && apiFormat !== 'chat-completions'
      && apiFormat !== 'anthropic-messages'
    ) {
      return sendError(res, 400, requestId, {
        source: 'proxy',
        code: 'invalid_api_format',
        message: "apiFormat must be 'responses', 'chat-completions', or 'anthropic-messages'.",
      });
    }
    const isAnthropic = apiFormat === 'anthropic-messages';
    const upstreamEndpoint = isAnthropic ? foundryEndpoint : endpoint;
    const upstreamSource = isAnthropic ? 'azure_foundry' : 'azure_openai';
    const provider = isAnthropic ? 'foundry_anthropic' : 'azure_openai';
    if (!upstreamEndpoint) {
      return sendError(res, 503, requestId, {
        source: 'proxy',
        code: 'proxy_not_configured',
        message: isAnthropic
          ? 'Microsoft Foundry is not configured on the server.'
          : 'Azure OpenAI is not configured on the server.',
      });
    }
    if (typeof deployment !== 'string' || !DEPLOYMENT_NAME_RE.test(deployment)) {
      return sendError(res, 400, requestId, {
        source: 'proxy',
        code: 'invalid_deployment_name',
        message: 'The deployment name is invalid.',
      });
    }
    const deploymentAllowlist = isAnthropic
      ? allowedFoundryDeployments
      : allowedDeployments;
    if (deploymentAllowlist.size === 0) {
      return sendError(res, 503, requestId, {
        source: 'proxy',
        code: 'deployment_allowlist_not_configured',
        message: isAnthropic
          ? 'Microsoft Foundry deployment access is not configured on the server.'
          : 'Azure OpenAI deployment access is not configured on the server.',
      });
    }
    if (!deploymentAllowlist.has(deployment)) {
      return sendError(res, 403, requestId, {
        source: 'proxy',
        code: 'deployment_not_allowed',
        message: 'The deployment is not allowed by the server configuration.',
      });
    }
    if (!body || typeof body !== 'object' || Array.isArray(body)) {
      return sendError(res, 400, requestId, {
        source: 'proxy',
        code: 'missing_request_body',
        message: 'The Azure OpenAI request body is missing.',
      });
    }

    const retryAfter = await consumeRateLimit(req);
    if (retryAfter > 0) {
      res.set('Retry-After', String(retryAfter));
      return sendError(res, 429, requestId, {
        source: 'proxy',
        code: 'proxy_rate_limit_exceeded',
        message: 'The application OpenAI request limit was exceeded.',
      });
    }

    const upstreamBody = { ...body };
    if (apiFormat === 'responses') {
      upstreamBody.model = deployment;
      upstreamBody.store = false;
      upstreamBody.max_output_tokens = Math.min(
        Math.max(Number(upstreamBody.max_output_tokens) || 1, 1),
        32768,
      );
    } else if (apiFormat === 'chat-completions') {
      upstreamBody.max_tokens = Math.min(
        Math.max(Number(upstreamBody.max_tokens) || 1, 1),
        32768,
      );
    } else {
      if (!Array.isArray(upstreamBody.messages) || upstreamBody.messages.length === 0) {
        return sendError(res, 400, requestId, {
          source: 'proxy',
          code: 'invalid_upstream_request',
          message: 'Anthropic Messages requests require a non-empty messages array.',
        });
      }
      const requestedEffort = upstreamBody.output_config?.effort;
      const effort = ['low', 'medium', 'high', 'max'].includes(requestedEffort)
        ? requestedEffort
        : 'low';
      upstreamBody.model = deployment;
      upstreamBody.max_tokens = Math.min(
        Math.max(Number(upstreamBody.max_tokens) || 1, 1),
        32768,
      );
      upstreamBody.thinking = { type: 'adaptive' };
      upstreamBody.output_config = { effort };
      upstreamBody.stream = false;
    }

    const headers = { 'Content-Type': 'application/json' };
    const selectedApiKey = isAnthropic ? foundryApiKey : apiKey;
    if (selectedApiKey) {
      headers[isAnthropic ? 'x-api-key' : 'api-key'] = selectedApiKey;
    } else {
      try {
        const tokenScope = isAnthropic
          ? 'https://ai.azure.com/.default'
          : 'https://cognitiveservices.azure.com/.default';
        const tokenResult = await credential?.getToken(tokenScope);
        if (!tokenResult?.token) throw new Error('Credential returned no token');
        headers.Authorization = `Bearer ${tokenResult.token}`;
      } catch (error) {
        logEvent(logger, 'error', {
          event: 'credential_acquisition_failed',
          requestId,
          deployment,
          apiFormat,
          provider,
          errorName: error?.name || 'Error',
          errorCode: error?.code || null,
        });
        return sendError(res, 502, requestId, {
          source: 'credential',
          code: 'credential_acquisition_failed',
          message: 'The server could not acquire an Azure OpenAI credential.',
        });
      }
    }
    if (isAnthropic) {
      headers['anthropic-version'] = anthropicVersion;
    }

    let upstream;
    try {
      upstream = await fetchImpl(buildOpenAIUrl(upstreamEndpoint, deployment, apiFormat, apiVersion), {
        method: 'POST',
        headers,
        body: JSON.stringify(upstreamBody),
        signal: AbortSignal.timeout(timeoutMs),
      });
    } catch (error) {
      const timedOut = error?.name === 'AbortError' || error?.name === 'TimeoutError';
      const status = timedOut ? 504 : 502;
      const code = timedOut ? 'azure_openai_timeout' : 'azure_openai_connection_failed';
      logEvent(logger, 'error', {
        event: code,
        requestId,
        deployment,
        apiFormat,
        provider,
        durationMs: Date.now() - startedAt,
        errorName: error?.name || 'Error',
        errorCode: error?.code || null,
      });
      return sendError(res, status, requestId, {
        source: 'proxy_transport',
        code,
        message: timedOut
          ? 'The Azure OpenAI request timed out.'
          : 'The server could not connect to Azure OpenAI.',
      });
    }

    const contentType = upstream.headers.get('content-type') || '';
    const upstreamRequestId = getHeader(upstream.headers, [
      'apim-request-id',
      'x-ms-request-id',
      'x-request-id',
      'request-id',
      'trace-id',
    ]);
    if (upstreamRequestId) {
      res.set('X-Upstream-Request-Id', upstreamRequestId);
    }
    const retryAfterHeader = upstream.headers.get('retry-after');
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
        deployment,
        apiFormat,
        provider,
        upstreamStatus: upstream.status,
        upstreamRequestId,
        durationMs: Date.now() - startedAt,
        errorName: error?.name || 'Error',
        errorCode: error?.code || null,
      });
      return sendError(res, 502, requestId, {
        source: 'proxy_transport',
        code: 'azure_openai_connection_failed',
        message: 'The server could not read the Azure OpenAI response.',
        upstreamStatus: upstream.status,
        upstreamRequestId,
      });
    }
    if (!upstream.ok) {
      const { code: upstreamCode, message: upstreamMessage } = parseUpstreamError(text);
      const classified = classifyUpstreamError(
        upstream.status,
        contentType,
        upstreamCode,
        upstreamMessage,
      );
      logEvent(logger, 'error', {
        event: classified.code,
        requestId,
        deployment,
        apiFormat,
        provider,
        upstreamStatus: upstream.status,
        upstreamCode,
        upstreamRequestId,
        contentType: contentType.slice(0, 128),
        durationMs: Date.now() - startedAt,
      });
      return sendError(res, upstream.status, requestId, {
        source: upstreamSource,
        code: classified.code,
        message: classified.message,
        upstreamStatus: upstream.status,
        upstreamCode,
        upstreamRequestId,
      });
    }

    if (!contentType.toLowerCase().includes('json')) {
      logEvent(logger, 'error', {
        event: 'invalid_upstream_response',
        requestId,
        deployment,
        apiFormat,
        provider,
        upstreamStatus: upstream.status,
        upstreamRequestId,
        contentType: contentType.slice(0, 128),
        durationMs: Date.now() - startedAt,
      });
      return sendError(res, 502, requestId, {
        source: upstreamSource,
        code: 'invalid_upstream_response',
        message: 'Azure OpenAI returned an unexpected response format.',
        upstreamStatus: upstream.status,
        upstreamRequestId,
      });
    }

    logEvent(logger, 'info', {
      event: 'request_succeeded',
      requestId,
      deployment,
      apiFormat,
      provider,
      upstreamStatus: upstream.status,
      upstreamRequestId,
      durationMs: Date.now() - startedAt,
    });
    res.status(upstream.status);
    res.set('Content-Type', contentType);
    return res.send(text);
  }));

  return router;
}

module.exports = {
  buildOpenAIUrl,
  classifyUpstreamError,
  createOpenAIProxyRouter,
  logFoundryConfiguration,
  parseUpstreamError,
};
