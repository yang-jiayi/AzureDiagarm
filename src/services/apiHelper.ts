// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

import type { ReasoningEffort } from '../stores/modelSettingsStore';

/**
 * API Format Helper
 * Abstracts the difference between Azure OpenAI Responses API and Chat Completions API.
 * OpenAI models (GPT-5.x) use the Responses API; third-party models
 * (DeepSeek, Grok) use the Chat Completions API via Azure AI model inference.
 */

export type ApiFormat = 'responses' | 'chat-completions';

/**
 * Build the correct API URL for the given format.
 * - Responses API:       {endpoint}openai/v1/responses
 * - Chat Completions:    {endpoint}openai/deployments/{deployment}/chat/completions?api-version=2024-12-01-preview
 */
export function buildApiUrl(endpoint: string, deployment: string, apiFormat: ApiFormat): string {
  if (apiFormat === 'chat-completions') {
    return `${endpoint}openai/deployments/${deployment}/chat/completions?api-version=2024-05-01-preview`;
  }
  return `${endpoint}openai/v1/responses`;
}

/**
 * Build the request body for the given API format.
 * Handles reasoning config only for Responses API models that support it.
 */
export function buildRequestBody(params: {
  deployment: string;
  messages: any[];
  maxTokens: number;
  apiFormat: ApiFormat;
  isReasoning: boolean;
  reasoningEffort: ReasoningEffort;
  jsonOutput?: boolean;
}): any {
  const { deployment, messages, maxTokens, apiFormat, isReasoning, reasoningEffort, jsonOutput = true } = params;

  if (apiFormat === 'chat-completions') {
    return {
      messages,
      max_tokens: maxTokens,
      ...(jsonOutput ? { response_format: { type: 'json_object' } } : {}),
      temperature: 0.7,
    };
  }

  // Responses API
  const body: any = {
    model: deployment,
    input: messages,
    max_output_tokens: maxTokens,
    ...(jsonOutput ? { text: { format: { type: 'json_object' } } } : {}),
    store: false,
  };

  if (isReasoning) {
    body.reasoning = { effort: reasoningEffort };
  }

  return body;
}

/**
 * Parse the API response into a uniform shape regardless of API format.
 */
export function parseApiResponse(
  data: any,
  apiFormat: ApiFormat,
): { content: string; promptTokens: number; completionTokens: number; totalTokens: number } {
  if (apiFormat === 'chat-completions') {
    const usage = data.usage || {};
    return {
      content: data.choices?.[0]?.message?.content || '',
      promptTokens: usage.prompt_tokens || 0,
      completionTokens: usage.completion_tokens || 0,
      totalTokens: usage.total_tokens || 0,
    };
  }

  // Responses API
  let content = data.output_text || '';
  if (!content && data.output) {
    for (const item of data.output) {
      if (item.type === 'message' && item.content) {
        for (const part of item.content) {
          if (part.type === 'output_text') {
            content += part.text;
          }
        }
      }
    }
  }

  const usage = data.usage || {};
  return {
    content,
    promptTokens: usage.input_tokens || 0,
    completionTokens: usage.output_tokens || 0,
    totalTokens: usage.total_tokens || 0,
  };
}

/**
 * Result of a call to the server-side Azure OpenAI proxy.
 */
export interface OpenAIProxyResult {
  ok: boolean;
  status: number;
  data: any;
  errorText?: string;
  error?: OpenAIProxyErrorDetails;
}

export interface OpenAIProxyErrorDetails {
  source: string;
  code: string;
  message?: string;
  requestId?: string;
  upstreamStatus?: number;
  upstreamCode?: string;
  upstreamRequestId?: string;
  contentType?: string;
  responseUrl?: string;
  redirected?: boolean;
}

export class OpenAIProxyError extends Error {
  readonly status: number;
  readonly code: string;
  readonly source: string;
  readonly requestId?: string;
  readonly upstreamRequestId?: string;

  constructor(message: string, result: OpenAIProxyResult) {
    super(message);
    this.name = 'OpenAIProxyError';
    this.status = result.status;
    this.code = result.error?.code || 'unknown_error';
    this.source = result.error?.source || 'unknown';
    this.requestId = result.error?.requestId;
    this.upstreamRequestId = result.error?.upstreamRequestId;
  }
}

function parseJson(text: string): any | null {
  if (!text) return null;
  try {
    return JSON.parse(text);
  } catch {
    return null;
  }
}

function getStructuredError(payload: any): Partial<OpenAIProxyErrorDetails> | null {
  if (!payload || typeof payload !== 'object') return null;
  const candidate = payload.error;
  if (!candidate || typeof candidate !== 'object') return null;
  return candidate;
}

function isAuthenticationUrl(url: string): boolean {
  return /\/\.auth\/(?:login|me|refresh)|login\.microsoftonline\.com/i.test(url);
}

function inferUnstructuredError(
  response: Response,
  responseText: string,
  contentType: string,
): OpenAIProxyErrorDetails {
  const requestId = response.headers.get('x-azurediagarm-request-id') || undefined;
  const upstreamRequestId = response.headers.get('x-upstream-request-id') || undefined;
  const common = {
    requestId,
    upstreamRequestId,
    contentType,
    responseUrl: response.url,
    redirected: response.redirected,
  };

  if (response.redirected || isAuthenticationUrl(response.url)) {
    return {
      ...common,
      source: 'application_auth',
      code: 'application_authentication_required',
      message: 'Application sign-in is required.',
    };
  }

  const looksLikeHtml = contentType.toLowerCase().includes('text/html')
    || /^\s*<!doctype html|^\s*<html/i.test(responseText);
  if (looksLikeHtml && response.status === 403) {
    return {
      ...common,
      source: 'edge',
      code: 'edge_request_blocked',
      message: 'The request was blocked before it reached the application.',
    };
  }
  if (looksLikeHtml && [502, 503, 504].includes(response.status)) {
    return {
      ...common,
      source: 'edge',
      code: 'edge_origin_unavailable',
      message: 'The application edge could not reach the origin.',
    };
  }
  if (response.status === 401) {
    return {
      ...common,
      source: 'application_auth',
      code: 'application_authentication_required',
      message: 'Application sign-in is required.',
    };
  }
  if (response.status === 403) {
    return {
      ...common,
      source: 'application_auth',
      code: 'application_request_rejected',
      message: 'The application authentication layer rejected the request.',
    };
  }

  return {
    ...common,
    source: 'unknown',
    code: `http_${response.status || 0}`,
    message: 'The Azure OpenAI request failed.',
  };
}

export function createOpenAIProxyError(
  result: OpenAIProxyResult,
  options: { vision?: boolean } = {},
): OpenAIProxyError {
  const code = result.error?.code || `http_${result.status || 0}`;
  let message: string;

  switch (code) {
    case 'application_authentication_required':
      message = 'Your application session is no longer valid. Refresh the page and sign in again.';
      break;
    case 'application_access_denied':
      message = 'Your account is not allowed to use this application.';
      break;
    case 'application_request_rejected':
      message = 'The application authentication layer rejected the request. Refresh the page and try again.';
      break;
    case 'edge_request_blocked':
      message = 'The request was blocked before it reached Azure OpenAI. Reduce the request size or contact the administrator.';
      break;
    case 'deployment_not_allowed':
      message = 'The selected model deployment is not allowed by the server configuration.';
      break;
    case 'credential_acquisition_failed':
      message = 'The server could not acquire an Azure OpenAI credential. Contact the administrator.';
      break;
    case 'azure_openai_authentication_failed':
      message = 'Azure OpenAI rejected the server credential. Check the managed identity role assignment.';
      break;
    case 'deployment_not_found':
      message = 'Deployment not found. Please check your model deployment name.';
      break;
    case 'proxy_rate_limit_exceeded':
      message = 'The application request limit was reached. Wait a moment and try again.';
      break;
    case 'azure_openai_rate_limited':
      message = 'Azure OpenAI is rate-limiting requests. Wait a moment and try again.';
      break;
    case 'azure_openai_timeout':
    case 'edge_origin_unavailable':
      message = 'Azure OpenAI is taking too long to respond. Please try again.';
      break;
    case 'azure_openai_unavailable':
    case 'azure_openai_connection_failed':
      message = 'Azure OpenAI is temporarily unavailable. Please try again.';
      break;
    case 'request_too_large':
      message = 'The request is too large. Reduce the diagram or image size and try again.';
      break;
    case 'image_not_supported':
      message = 'The selected model may not support image analysis. Try using GPT-5.6 Sol.';
      break;
    case 'content_filtered':
      message = 'Azure OpenAI content filtering rejected the request. Revise the prompt and try again.';
      break;
    case 'invalid_upstream_response':
    case 'azure_openai_non_json_error':
      message = 'Azure OpenAI returned an unexpected response. Please try again.';
      break;
    case 'network_error':
      message = 'The application could not reach the Azure OpenAI proxy. Check your connection and try again.';
      break;
    case 'invalid_upstream_request':
      message = options.vision
        ? 'The selected model may not support image analysis. Try using GPT-5.6 Sol.'
        : 'Azure OpenAI rejected the request format. Please try again or simplify the request.';
      break;
    default:
      message = `Azure OpenAI request failed (${result.status || 'network error'}). Please try again.`;
  }

  if (result.error?.requestId) {
    message = `${message} Request ID: ${result.error.requestId}`;
  }

  return new OpenAIProxyError(message, result);
}

/**
 * Call Azure OpenAI through the server-side proxy (/api/openai).
 *
 * The proxy holds the Azure OpenAI credentials (managed identity, with optional
 * key fallback) so they are never shipped to the browser. The client sends the
 * already-built request body plus the deployment name and API format; the server
 * constructs the upstream URL from its trusted endpoint and attaches auth.
 */
export async function callAzureOpenAIProxy(params: {
  apiFormat: ApiFormat;
  deployment: string;
  body: any;
  signal?: AbortSignal;
}): Promise<OpenAIProxyResult> {
  let response: Response;
  try {
    response = await fetch('/api/openai', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({
        apiFormat: params.apiFormat,
        deployment: params.deployment,
        body: params.body,
      }),
      signal: params.signal,
    });
  } catch (error) {
    if (error instanceof Error && error.name === 'AbortError') throw error;
    const message = error instanceof Error ? error.message : 'Network request failed';
    return {
      ok: false,
      status: 0,
      data: null,
      errorText: message,
      error: {
        source: 'network',
        code: 'network_error',
        message,
      },
    };
  }

  const contentType = response.headers.get('content-type') || '';
  const responseText = await response.text().catch(() => '');
  const parsed = parseJson(responseText);
  const structured = getStructuredError(parsed);
  const requestId = response.headers.get('x-azurediagarm-request-id')
    || structured?.requestId
    || undefined;
  const upstreamRequestId = response.headers.get('x-upstream-request-id')
    || structured?.upstreamRequestId
    || undefined;

  if (!response.ok || response.redirected || isAuthenticationUrl(response.url)) {
    const inferred = inferUnstructuredError(response, responseText, contentType);
    const error: OpenAIProxyErrorDetails = structured?.source && structured?.code
      ? {
          source: String(structured.source),
          code: String(structured.code),
          message: typeof structured.message === 'string' ? structured.message : undefined,
          requestId,
          upstreamStatus: typeof structured.upstreamStatus === 'number'
            ? structured.upstreamStatus
            : undefined,
          upstreamCode: typeof structured.upstreamCode === 'string'
            ? structured.upstreamCode
            : undefined,
          upstreamRequestId,
          contentType,
          responseUrl: response.url,
          redirected: response.redirected,
        }
      : inferred;
    return {
      ok: false,
      status: response.ok ? 401 : response.status,
      data: parsed,
      errorText: responseText.slice(0, 2_000),
      error,
    };
  }

  if (!parsed || !contentType.toLowerCase().includes('json')) {
    const error: OpenAIProxyErrorDetails = {
      source: 'proxy',
      code: 'invalid_upstream_response',
      message: 'The proxy returned an unexpected response format.',
      requestId,
      upstreamRequestId,
      contentType,
      responseUrl: response.url,
      redirected: response.redirected,
    };
    return {
      ok: false,
      status: 502,
      data: null,
      errorText: responseText.slice(0, 2_000),
      error,
    };
  }

  return { ok: true, status: response.status, data: parsed };
}
