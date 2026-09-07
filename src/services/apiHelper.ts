// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

import type { ReasoningEffort } from '../stores/modelSettingsStore';
import { AIModelConfigurationError, getDeploymentName, isModelAvailable, isReasoningEffort } from '../stores/modelSettingsStore';
import { getBYOAISettings, normalizeBYOAIEndpoint, validateBYOAIProfile } from '../stores/byoAISettingsStore';
import { isValidBYOAIApiKey, readBYOAIConnectionSecret } from './byoAIConnectionSession';
import { assertCapturedAIConnectionCurrent, type CapturedAIConnection } from './aiModelRuntime';
import { awaitWithAISignal, isBYOAIEnabledOnServer, runtimeConfigCancellationError } from './runtimeConfig';

/**
 * Managed Astra Responses and explicitly selected BYO OpenAI request boundary.
 */

export type ApiFormat = 'responses' | 'chat-completions';

function assertApiFormat(apiFormat: unknown): asserts apiFormat is ApiFormat {
  if (apiFormat !== 'responses' && apiFormat !== 'chat-completions') {
    throw new AIModelConfigurationError('unsupported_api_format', 'Use the Responses or Chat Completions API.');
  }
}

export interface BYOAIProxyConfig {
  provider: 'azure-openai' | 'openai';
  endpoint: string;
  apiKey: string;
}

export function getApiFormatLabel(apiFormat: ApiFormat): string {
  assertApiFormat(apiFormat);
  return apiFormat === 'responses' ? 'Responses' : 'Chat Completions';
}

function chatMessages(messages: any[]): any[] {
  return structuredClone(messages).map(message => ({
    ...message,
    ...(Array.isArray(message.content) ? {
      content: message.content.map((part: any) => {
        if (part?.type === 'input_text') return { ...part, type: 'text' };
        if (part?.type !== 'input_image') return part;
        const { type: _type, image_url, detail, ...rest } = part;
        return {
          ...rest, type: 'image_url',
          image_url: {
            ...(typeof image_url === 'string' ? { url: image_url } : image_url),
            ...(detail !== undefined ? { detail } : {}),
          },
        };
      }),
    } : {}),
  }));
}

/**
 * Build either OpenAI format without changing the caller's prompt or quality.
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

  assertApiFormat(apiFormat);
  if (!Number.isSafeInteger(maxTokens) || maxTokens < 1 || maxTokens > 32768) {
    throw new AIModelConfigurationError('invalid_output_limit', 'The AI output limit must be a whole number from 1 to 32768.');
  }
  if (!isReasoningEffort(reasoningEffort)) {
    throw new AIModelConfigurationError('unsupported_reasoning_effort', 'The requested reasoning effort is not supported.');
  }
  if (apiFormat === 'chat-completions') {
    return {
      model: deployment,
      messages: chatMessages(messages),
      ...(isReasoning
        ? { max_completion_tokens: maxTokens, reasoning_effort: reasoningEffort }
        : { max_tokens: maxTokens }),
      ...(jsonOutput ? { response_format: { type: 'json_object' } } : {}),
      store: false,
    };
  }
  const body: any = {
    model: deployment,
    input: structuredClone(messages),
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
 * Extract text and usage without trusting malformed text/token fields.
 */
export function parseApiResponse(
  data: any,
  apiFormat: ApiFormat,
): { content: string; promptTokens: number; completionTokens: number; totalTokens: number } {
  assertApiFormat(apiFormat);
  const tokens = (value: unknown) => typeof value === 'number' && Number.isSafeInteger(value) && value >= 0 ? value : 0;
  const usage = data?.usage || {};
  if (apiFormat === 'chat-completions') {
    return {
      content: typeof data?.choices?.[0]?.message?.content === 'string' ? data.choices[0].message.content : '',
      promptTokens: tokens(usage.prompt_tokens),
      completionTokens: tokens(usage.completion_tokens),
      totalTokens: tokens(usage.total_tokens),
    };
  }
  let content = typeof data?.output_text === 'string' ? data.output_text : '';
  if (!content && Array.isArray(data?.output)) {
    for (const item of data.output) {
      if (item?.type === 'message' && Array.isArray(item.content)) {
        for (const part of item.content) {
          if (part?.type === 'output_text' && typeof part.text === 'string') {
            content += part.text;
          }
        }
      }
    }
  }

  return {
    content,
    promptTokens: tokens(usage.input_tokens),
    completionTokens: tokens(usage.output_tokens),
    totalTokens: tokens(usage.total_tokens),
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
  /** Validated provider delay, including Retry-After HTTP dates. */
  retryAfterMs?: number;
}

export class OpenAIProxyError extends Error {
  readonly status: number;
  readonly code: string;
  readonly source: string;
  readonly requestId?: string;
  readonly upstreamRequestId?: string;
  readonly retryAfterMs?: number;
  readonly upstreamStatus?: number;
  readonly upstreamCode?: string;

  constructor(message: string, result: OpenAIProxyResult) {
    super(message);
    this.name = 'OpenAIProxyError';
    this.status = result.status;
    this.code = result.error?.code || 'unknown_error';
    this.source = result.error?.source || 'unknown';
    this.requestId = result.error?.requestId;
    this.upstreamRequestId = result.error?.upstreamRequestId;
    this.retryAfterMs = result.error?.retryAfterMs;
    this.upstreamStatus = result.error?.upstreamStatus;
    this.upstreamCode = result.error?.upstreamCode;
  }
}

export function isAIBudgetError(error: unknown): boolean {
  if (!error || typeof error !== 'object') return false;
  const value = error as { source?: unknown; code?: unknown };
  return value.source === 'budget' || (typeof value.code === 'string' && (
    value.code === 'ai_daily_budget_exceeded'
    || value.code === 'ai_concurrency_limit'
    || value.code.startsWith('ai_budget_')
  ));
}

export function isAIInternalServerError(error: unknown): boolean {
  if (!error || typeof error !== 'object' || isAIBudgetError(error)) return false;
  const value = error as { status?: unknown; upstreamStatus?: unknown };
  return value.status === 500 || value.upstreamStatus === 500;
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

function diagnosticToken(value: unknown, secret?: string): string | undefined {
  return typeof value === 'string' && /^[A-Za-z0-9][A-Za-z0-9_.:-]{0,127}$/.test(value)
    && !/^sk-/i.test(value) && !(secret && value.includes(secret)) ? value : undefined;
}

function readRetryAfterMs(headers: Headers): number | undefined {
  const delays: number[] = [];
  const add = (value: number) => {
    const rounded = Math.ceil(value);
    if (Number.isSafeInteger(rounded) && rounded >= 0) delays.push(rounded);
  };
  const retryAfter = headers.get('retry-after')?.trim();
  if (retryAfter && /^\d+(?:\.\d+)?$/.test(retryAfter)) {
    add(Number(retryAfter) * 1000);
  } else if (retryAfter && /GMT$/i.test(retryAfter)) {
    const deadline = Date.parse(retryAfter);
    const serverDate = Date.parse(headers.get('date') || '');
    // The response Date avoids retrying early when the client's clock is ahead.
    add(Math.max(0, deadline - (Number.isFinite(serverDate) ? serverDate : Date.now())));
  }
  for (const name of ['retry-after-ms', 'x-ms-retry-after-ms']) {
    const value = headers.get(name)?.trim();
    if (value && /^\d+(?:\.\d+)?$/.test(value)) add(Number(value));
  }
  return delays.length ? Math.max(...delays) : undefined;
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
    message: 'The AI provider request failed.',
  };
}

export function isJsonMediaType(contentType: string): boolean {
  const mediaType = contentType.split(';', 1)[0].trim().toLowerCase();
  return /^application\/(?:[a-z0-9!#$&^_.+-]+\+)?json$/.test(mediaType);
}

/**
 * Every proxy error code that maps to a stable, pre-translatable user message.
 * The i18n coverage test iterates this list to assert each producible English
 * message has an exact Japanese entry, so the two can never silently drift.
 * Codes with a purely dynamic message (the `default` fallback) are excluded.
 */
export const PROXY_ERROR_MESSAGE_CODES = [
  'application_authentication_required',
  'authentication_required',
  'application_access_denied',
  'application_request_rejected',
  'edge_request_blocked',
  'deployment_not_allowed',
  'astra_not_configured',
  'proxy_not_configured',
  'invalid_api_format',
  'byo_not_enabled',
  'byo_authentication_failed',
  'byo_rate_limited',
  'byo_timeout',
  'byo_unavailable',
  'byo_connection_failed',
  'invalid_byo_configuration',
  'invalid_byo_provider',
  'invalid_byo_endpoint',
  'invalid_byo_api_key',
  'invalid_deployment_name',
  'credential_acquisition_failed',
  'azure_openai_authentication_failed',
  'deployment_not_found',
  'proxy_rate_limit_exceeded',
  'azure_openai_rate_limited',
  'http_429',
  'ai_daily_budget_exceeded',
  'ai_concurrency_limit',
  'ai_budget_busy',
  'ai_budget_unavailable',
  'ai_budget_timeout',
  'azure_openai_timeout',
  'edge_origin_unavailable',
  'azure_openai_unavailable',
  'azure_openai_connection_failed',
  'request_too_large',
  'image_not_supported',
  'content_filtered',
  'invalid_upstream_response',
  'azure_openai_non_json_error',
  'network_error',
  'invalid_upstream_request',
] as const;

/**
 * Resolve the user-facing message for a proxy error code. Pure and exported so
 * the i18n coverage test can enumerate every producible message.
 */
export function proxyErrorMessageForCode(
  code: string,
  options: { vision?: boolean; status?: number } = {},
): string {
  switch (code) {
    case 'application_authentication_required':
    case 'authentication_required':
      return 'Your application session is no longer valid. Refresh the page and sign in again.';
    case 'application_access_denied':
      return 'Your account is not allowed to use this application.';
    case 'application_request_rejected':
      return 'The application authentication layer rejected the request. Refresh the page and try again.';
    case 'edge_request_blocked':
      return 'The request was blocked before it reached the AI provider. Reduce the request size or contact the administrator.';
    case 'deployment_not_allowed':
      return 'Only the configured GPT-6 Astra deployment can run.';
    case 'astra_not_configured':
      return 'GPT-6 Astra is not configured. Contact the application administrator to configure the managed Astra deployment.';
    case 'proxy_not_configured':
      return 'The managed Azure OpenAI endpoint is not configured correctly.';
    case 'invalid_api_format':
      return 'GPT-6 Astra requests must use the Responses API.';
    case 'byo_not_enabled':
      return 'Bring-your-own AI is disabled by the application administrator.';
    case 'byo_authentication_failed':
    case 'invalid_byo_api_key':
      return 'The AI provider rejected this profile’s API key. Re-enter the key and test the connection again.';
    case 'invalid_byo_configuration':
    case 'invalid_byo_provider':
    case 'invalid_byo_endpoint':
      return 'The AI connection settings are invalid. Check the provider, endpoint, and model, then test again.';
    case 'credential_acquisition_failed':
      return 'The server could not acquire an Azure OpenAI credential. Contact the administrator.';
    case 'azure_openai_authentication_failed':
      return 'Azure OpenAI rejected the server credential. Check the managed identity role assignment.';
    case 'deployment_not_found':
    case 'invalid_deployment_name':
      return 'Model or deployment not found. Check the configured name.';
    case 'proxy_rate_limit_exceeded':
      return 'The application request limit was reached. Wait a moment and try again.';
    case 'azure_openai_rate_limited':
    case 'byo_rate_limited':
      return 'The AI provider is rate-limiting requests. Wait a moment and try again.';
    case 'http_429':
      return 'The AI request was rate-limited, but the response did not identify which limit was reached. Wait before retrying or contact the administrator with the request ID.';
    case 'ai_daily_budget_exceeded':
      return 'The application daily AI budget cannot cover this request. Wait until midnight UTC for the budget to reset, or contact the administrator. Failed requests with unknown usage may still count toward this budget.';
    case 'ai_concurrency_limit':
      return 'The application concurrent AI request limit was reached. Wait for an active request to finish, then try again.';
    case 'ai_budget_busy':
      return 'The application AI budget is busy. Wait a few seconds and try again.';
    case 'ai_budget_unavailable':
      return 'The application AI budget could not be checked. Please try again later or contact the administrator.';
    case 'ai_budget_timeout':
      return 'The application timed out while reserving the AI budget. Wait a moment and try again.';
    case 'azure_openai_timeout':
    case 'byo_timeout':
    case 'edge_origin_unavailable':
      return 'The AI provider is taking too long to respond. Please try again.';
    case 'azure_openai_unavailable':
    case 'azure_openai_connection_failed':
    case 'byo_unavailable':
    case 'byo_connection_failed':
      return 'The AI provider is temporarily unavailable. Please try again.';
    case 'request_too_large':
      return 'The request is too large. Reduce the diagram or image size and try again.';
    case 'image_not_supported':
      return 'The configured GPT-6 Astra deployment rejected the image analysis request. Check the image and contact the application administrator if the problem persists.';
    case 'content_filtered':
      return 'The AI provider content policy rejected the request. Revise the prompt and try again.';
    case 'invalid_upstream_response':
    case 'azure_openai_non_json_error':
      return 'The AI provider returned an unexpected response. Please try again.';
    case 'network_error':
      return 'The application could not reach the Azure OpenAI proxy. Check your connection and try again.';
    case 'invalid_upstream_request':
      return options.vision
        ? proxyErrorMessageForCode('image_not_supported')
        : 'The AI provider rejected the request format. Please try again or simplify the request.';
    default:
      if (options.status === 429) return proxyErrorMessageForCode('http_429');
      return `AI provider request failed (${options.status || 'network error'}). Please try again.`;
  }
}

export function createOpenAIProxyError(
  result: OpenAIProxyResult,
  options: { vision?: boolean } = {},
): OpenAIProxyError {
  const code = result.error?.code || `http_${result.status || 0}`;
  let message = proxyErrorMessageForCode(code, {
    vision: options.vision,
    status: result.status,
  });

  const requestId = diagnosticToken(result.error?.requestId);
  if (requestId) {
    message = `${message} Request ID: ${requestId}`;
  }

  return new OpenAIProxyError(message, result);
}

/**
 * Call the protected server-side proxy (/api/openai). Never call a provider directly.
 *
 * Managed credentials stay on the server. BYO keys stay in tab memory and are
 * sent only to this proxy, which enforces opt-in, endpoint policy, authentication
 * and budgets before constructing the upstream URL.
 */
export async function callAzureOpenAIProxy(params: {
  apiFormat: ApiFormat;
  deployment: string;
  body: any;
  signal?: AbortSignal;
  byo?: BYOAIProxyConfig;
  connection?: CapturedAIConnection;
  purpose?: 'connection-test';
}): Promise<OpenAIProxyResult> {
  if (params.signal?.aborted) throw runtimeConfigCancellationError();
  assertApiFormat(params.apiFormat);
  if (Object.keys(params).some(key => !['apiFormat', 'deployment', 'body', 'signal', 'byo', 'connection', 'purpose'].includes(key))) {
    throw new AIModelConfigurationError('unsupported_ai_provider', 'Use a configured AI connection without additional routing fields.');
  }
  let byo: BYOAIProxyConfig | undefined;
  if (params.connection !== undefined) {
    assertCapturedAIConnectionCurrent(params.connection);
    if ('byo' in params || params.deployment !== params.connection.deployment || params.apiFormat !== params.connection.apiFormat) {
      throw new AIModelConfigurationError('stale_ai_configuration', 'The selected AI connection changed before this request was sent. Review the connection and submit again.');
    }
    if (params.connection.source === 'bring-your-own') {
      const profile = getBYOAISettings().profiles.find(item => item.id === params.connection!.profileId);
      const apiKey = profile && readBYOAIConnectionSecret(profile.id, params.connection.revision);
      if (!profile || !apiKey || !validateBYOAIProfile(profile).valid) {
        throw new AIModelConfigurationError('byo_invalid_profile', 'The selected AI connection is not ready. Edit the profile and test again.');
      }
      byo = { provider: profile.provider, endpoint: profile.endpoint, apiKey };
    }
  } else if ('byo' in params) {
    const value = params.byo;
    if (!value || typeof value !== 'object' || Array.isArray(value)
      || Object.keys(value).some(key => !['provider', 'endpoint', 'apiKey'].includes(key))
      || !isValidBYOAIApiKey(value.apiKey)) {
      throw new AIModelConfigurationError('invalid_byo_configuration', 'Enter valid AI connection settings and an API key.');
    }
    const endpoint = normalizeBYOAIEndpoint(value.provider, value.endpoint);
    if (!endpoint) throw new AIModelConfigurationError('invalid_byo_configuration', 'Use a trusted AI provider HTTPS origin.');
    byo = { provider: value.provider, endpoint, apiKey: value.apiKey };
  }
  if (byo) {
    if (!isBYOAIEnabledOnServer()) throw new AIModelConfigurationError('byo_not_enabled', 'Bring-your-own AI availability must be confirmed by the application server.');
  } else {
    if (params.apiFormat !== 'responses') {
      throw new AIModelConfigurationError('unsupported_api_format', 'GPT-6 Astra requests must use the Responses API.');
    }
    const deployment = getDeploymentName('gpt-6-astra');
    if (!isModelAvailable('gpt-6-astra')) {
      throw new AIModelConfigurationError('astra_not_configured', 'Azure OpenAI is not configured. Please check your environment.');
    }
    if (params.deployment !== deployment) {
      throw new AIModelConfigurationError('unsupported_ai_model', 'Only the configured GPT-6 Astra deployment can run.');
    }
  }
  const forbiddenBodyFields = ['byo', 'provider', 'deployment', 'endpoint', 'apiKey', 'api_key', 'headers', 'url', 'baseURL', 'base_url', 'forceManaged'];
  if (!params.body || typeof params.body !== 'object' || Array.isArray(params.body)
    || typeof params.deployment !== 'string' || !/^(?=.{1,128}$)(?=.*[A-Za-z0-9])[A-Za-z0-9._:-]+$/.test(params.deployment)
    || params.body.model !== params.deployment || forbiddenBodyFields.some(field => field in params.body)) {
    throw new AIModelConfigurationError('unsupported_ai_model', 'The request must use the selected connection’s exact model or deployment.');
  }
  for (const field of ['max_output_tokens', 'max_completion_tokens', 'max_tokens']) {
    const value = params.body[field];
    if (value !== undefined && (!Number.isSafeInteger(value) || value < 1 || value > 32768)) {
      throw new AIModelConfigurationError('invalid_output_limit', 'The AI output limit must be a whole number from 1 to 32768.');
    }
  }
  if (params.connection) {
    const connection = params.connection;
    const outputLimit = params.apiFormat === 'responses' ? params.body.max_output_tokens
      : connection.isReasoning ? params.body.max_completion_tokens : params.body.max_tokens;
    const effort = params.apiFormat === 'responses' ? params.body.reasoning?.effort : params.body.reasoning_effort;
    if (outputLimit !== connection.maxCompletionTokens
      || (connection.isReasoning ? effort !== connection.reasoningEffort : effort !== undefined)) {
      throw new AIModelConfigurationError('stale_ai_configuration', 'The request must preserve the selected connection’s reasoning and output settings.');
    }
    const messages = params.body.input ?? params.body.messages;
    if (!connection.supportsVision && Array.isArray(messages) && messages.some(message =>
      Array.isArray(message?.content) && message.content.some((part: any) =>
        part?.type === 'input_image' || part?.type === 'image_url'))) {
      throw new AIModelConfigurationError('byo_vision_not_supported', 'The selected AI connection does not support images. Select a vision-capable connection.');
    }
  }
  let response: Response;
  try {
    const request = fetch('/api/openai', {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json',
        ...(params.purpose === 'connection-test' ? { 'X-AzureDiagarm-Operation': 'byo-connection-test' } : {}),
      },
      body: JSON.stringify({
        apiFormat: params.apiFormat,
        deployment: params.deployment,
        body: params.body,
        ...(byo ? { byo } : {}),
      }),
      signal: params.signal,
      credentials: 'same-origin',
      redirect: 'error',
      cache: 'no-store',
    });
    response = params.signal ? await awaitWithAISignal(request, params.signal) : await request;
  } catch (error) {
    if (error instanceof Error && error.name === 'AbortError') throw error;
    const message = proxyErrorMessageForCode('network_error');
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
  let responseText = '';
  try {
    const text = response.text();
    responseText = params.signal ? await awaitWithAISignal(text, params.signal) : await text;
  } catch (error) {
    if (params.signal?.aborted) throw runtimeConfigCancellationError();
    if (error instanceof Error && error.name === 'AbortError') throw error;
  }
  const parsed = parseJson(responseText);
  const structured = getStructuredError(parsed);
  const safe = (value: unknown) => diagnosticToken(value, byo?.apiKey);
  const requestId = safe(response.headers.get('x-azurediagarm-request-id')) || safe(structured?.requestId);
  const upstreamRequestId = safe(response.headers.get('x-upstream-request-id')) || safe(structured?.upstreamRequestId);
  const retryAfterMs = readRetryAfterMs(response.headers);

  if (!response.ok || response.redirected || isAuthenticationUrl(response.url)) {
    const inferred = inferUnstructuredError(response, responseText, contentType);
    const source = safe(structured?.source) || (isAIBudgetError(structured) ? 'budget' : undefined);
    const error: OpenAIProxyErrorDetails = safe(structured?.code)
      ? {
          source: String(source || inferred.source),
          code: safe(structured?.code)!,
          message: proxyErrorMessageForCode(safe(structured?.code)!, { status: response.status }),
          requestId,
          upstreamStatus: typeof structured?.upstreamStatus === 'number' && Number.isInteger(structured.upstreamStatus)
            && structured.upstreamStatus >= 100 && structured.upstreamStatus <= 599 ? structured.upstreamStatus
            : undefined,
          upstreamCode: safe(structured?.upstreamCode) || (!source ? safe(structured?.code) : undefined),
          upstreamRequestId,
          redirected: response.redirected,
        }
      : {
          ...inferred,
          ...(source ? { source: String(source) } : {}),
          upstreamCode: safe(structured?.code),
        };
    return {
      ok: false,
      status: response.ok ? 401 : response.status,
      data: null,
      errorText: proxyErrorMessageForCode(error.code, { status: response.status }),
      error: { ...error, contentType: undefined, requestId, upstreamRequestId, retryAfterMs },
    };
  }

  if (!parsed || !isJsonMediaType(contentType)) {
    const error: OpenAIProxyErrorDetails = {
      source: 'proxy',
      code: 'invalid_upstream_response',
      message: 'The proxy returned an unexpected response format.',
      requestId,
      upstreamRequestId,
      redirected: response.redirected,
    };
    return {
      ok: false,
      status: 502,
      data: null,
      errorText: proxyErrorMessageForCode(error.code),
      error,
    };
  }

  return { ok: true, status: response.status, data: parsed };
}
