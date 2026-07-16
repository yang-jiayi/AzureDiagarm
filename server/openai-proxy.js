// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

const crypto = require('crypto');
const express = require('express');

const DEPLOYMENT_NAME_RE = /^[A-Za-z0-9._-]{1,64}$/;
const DEFAULT_API_VERSION = '2024-05-01-preview';
const DEFAULT_TIMEOUT_MS = 295_000;

function buildOpenAIUrl(endpoint, deployment, apiFormat, apiVersion) {
  const base = endpoint.endsWith('/') ? endpoint : `${endpoint}/`;
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
      message: 'Azure OpenAI rejected the server credential.',
    };
  }
  if (status === 404 || normalizedCode.includes('deploymentnotfound')) {
    return {
      code: 'deployment_not_found',
      message: 'The Azure OpenAI deployment was not found.',
    };
  }
  if (status === 429) {
    return {
      code: 'azure_openai_rate_limited',
      message: 'Azure OpenAI rate-limited the request.',
    };
  }
  if (status === 408 || status === 504) {
    return {
      code: 'azure_openai_timeout',
      message: 'Azure OpenAI timed out while processing the request.',
    };
  }
  if (status === 500 || status === 502 || status === 503) {
    return {
      code: 'azure_openai_unavailable',
      message: 'Azure OpenAI is temporarily unavailable.',
    };
  }
  if (status === 413) {
    return {
      code: 'request_too_large',
      message: 'The Azure OpenAI request is too large.',
    };
  }
  if (
    normalizedCode.includes('contentfilter')
    || normalizedCode.includes('content_filter')
    || normalizedMessage.includes('content filter')
  ) {
    return {
      code: 'content_filtered',
      message: 'Azure OpenAI content filtering rejected the request.',
    };
  }
  if (
    status === 400
    && (normalizedMessage.includes('image') || normalizedMessage.includes('vision'))
  ) {
    return {
      code: 'image_not_supported',
      message: 'The selected Azure OpenAI deployment rejected the image input.',
    };
  }
  if (status === 400 || status === 422) {
    return {
      code: 'invalid_upstream_request',
      message: 'Azure OpenAI rejected the request format.',
    };
  }
  if (String(contentType || '').toLowerCase().includes('text/html')) {
    return {
      code: 'azure_openai_non_json_error',
      message: 'Azure OpenAI returned an unexpected non-JSON error.',
    };
  }
  return {
    code: 'azure_openai_request_failed',
    message: 'Azure OpenAI rejected the request.',
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
    credential,
    apiKey,
    apiVersion = DEFAULT_API_VERSION,
    allowedDeployments = new Set(),
    fetchImpl = globalThis.fetch,
    timeoutMs = DEFAULT_TIMEOUT_MS,
    consumeRateLimit = () => 0,
    logger = console,
  } = options;

  if (typeof fetchImpl !== 'function') {
    throw new TypeError('fetchImpl must be a function');
  }

  const router = express.Router();
  router.post('/', async (req, res) => {
    const requestId = crypto.randomUUID();
    const startedAt = Date.now();
    res.set('X-AzureDiagarm-Request-Id', requestId);

    if (!endpoint) {
      return sendError(res, 503, requestId, {
        source: 'proxy',
        code: 'proxy_not_configured',
        message: 'Azure OpenAI is not configured on the server.',
      });
    }

    const { apiFormat, deployment, body } = req.body || {};
    if (apiFormat !== 'responses' && apiFormat !== 'chat-completions') {
      return sendError(res, 400, requestId, {
        source: 'proxy',
        code: 'invalid_api_format',
        message: "apiFormat must be 'responses' or 'chat-completions'.",
      });
    }
    if (typeof deployment !== 'string' || !DEPLOYMENT_NAME_RE.test(deployment)) {
      return sendError(res, 400, requestId, {
        source: 'proxy',
        code: 'invalid_deployment_name',
        message: 'The deployment name is invalid.',
      });
    }
    if (allowedDeployments.size > 0 && !allowedDeployments.has(deployment)) {
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

    const retryAfter = consumeRateLimit(req);
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
    } else {
      upstreamBody.max_tokens = Math.min(
        Math.max(Number(upstreamBody.max_tokens) || 1, 1),
        32768,
      );
    }

    const headers = { 'Content-Type': 'application/json' };
    if (apiKey) {
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
          deployment,
          apiFormat,
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

    let upstream;
    try {
      upstream = await fetchImpl(buildOpenAIUrl(endpoint, deployment, apiFormat, apiVersion), {
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

    const text = await upstream.text();
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
        upstreamStatus: upstream.status,
        upstreamCode,
        upstreamRequestId,
        contentType: contentType.slice(0, 128),
        durationMs: Date.now() - startedAt,
      });
      return sendError(res, upstream.status, requestId, {
        source: 'azure_openai',
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
        upstreamStatus: upstream.status,
        upstreamRequestId,
        contentType: contentType.slice(0, 128),
        durationMs: Date.now() - startedAt,
      });
      return sendError(res, 502, requestId, {
        source: 'azure_openai',
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
      upstreamStatus: upstream.status,
      upstreamRequestId,
      durationMs: Date.now() - startedAt,
    });
    res.status(upstream.status);
    res.set('Content-Type', contentType);
    return res.send(text);
  });

  return router;
}

module.exports = {
  buildOpenAIUrl,
  classifyUpstreamError,
  createOpenAIProxyRouter,
  parseUpstreamError,
};
