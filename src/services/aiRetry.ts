// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

/**
 * Shared retry classification for AI generation calls.
 *
 * Long editorial prompts (blueprint / reference architecture) occasionally hit
 * the proxy's 210s upstream budget or the client's 225s abort, which surfaces
 * as "The AI provider is taking too long to respond." Those failures are
 * transient and are usually cleared by a second, cheaper attempt, so callers
 * classify the error here instead of re-implementing the matching each time.
 */

import type { ReasoningEffort } from '../stores/modelSettingsStore';
import { getModelSettingsForFeature, type FeatureType } from '../stores/modelSettingsStore';
import type { RuntimeModelOverride } from './aiModelRuntime';

/** Proxy error codes that represent a transient upstream/edge condition. */
const RETRYABLE_PROXY_CODES = new Set([
  'azure_openai_timeout',
  'azure_openai_unavailable',
  'azure_openai_connection_failed',
  'byo_timeout',
  'byo_unavailable',
  'byo_connection_failed',
  'edge_origin_unavailable',
  'proxy_rate_limit_exceeded',
  'azure_openai_rate_limited',
  'byo_rate_limited',
]);

/** Message fragments emitted by the client-side abort / empty-response paths. */
const RETRYABLE_MESSAGE_PATTERNS = [
  /request timed out/i,
  /taking too long to respond/i,
  /temporarily unavailable/i,
  /timed out after/i,
  /empty response from azure openai/i,
  /may have timed out or returned empty content/i,
];

/**
 * True when the failure is worth one more (cheaper) attempt. Authentication,
 * authorization, configuration, and content-policy failures are deliberately
 * excluded — retrying those only wastes the user's time.
 */
export function isRetryableAIFailure(error: unknown): boolean {
  if (!error) return false;

  const code = (error as { code?: unknown }).code;
  if (typeof code === 'string' && RETRYABLE_PROXY_CODES.has(code)) return true;

  const name = (error as { name?: unknown }).name;
  if (name === 'AbortError' || name === 'TimeoutError') return true;

  const status = (error as { status?: unknown }).status;
  if (typeof status === 'number' && (status === 429 || status === 502 || status === 503 || status === 504)) {
    return true;
  }

  const message = error instanceof Error ? error.message : typeof error === 'string' ? error : '';
  return RETRYABLE_MESSAGE_PATTERNS.some((pattern) => pattern.test(message));
}

/**
 * Reasoning effort used for a retry. Anything above "low" is expensive in wall
 * clock time, which is exactly what caused the timeout, so the retry always
 * runs at "low" (kept, not dropped to "none", because some model families —
 * e.g. Claude — do not accept "none").
 */
export function downshiftReasoningEffort(effort: ReasoningEffort): ReasoningEffort {
  switch (effort) {
    case 'none':
    case 'minimal':
    case 'low':
      return effort;
    default:
      return 'low';
  }
}

/**
 * Build the model override for a retry attempt: same model, cheaper reasoning.
 * Returns undefined when the caller did not pass an override so the feature's
 * configured defaults keep applying.
 */
export function degradeOverrideForRetry<T extends RuntimeModelOverride>(
  override: T | undefined,
): T | undefined {
  if (!override) return undefined;
  const reasoningEffort = downshiftReasoningEffort(override.reasoningEffort);
  if (reasoningEffort === override.reasoningEffort) return override;
  return { ...override, reasoningEffort };
}

/**
 * Override to use for a retry, even when the caller passed none.
 *
 * `degradeOverrideForRetry` cannot help when the caller relies on the feature
 * defaults, because those defaults are exactly what timed out. Resolving the
 * feature's configured model here lets the retry keep the user's model while
 * dropping the reasoning effort that blew the time budget.
 */
export function buildRetryOverride(
  feature: FeatureType,
  override: RuntimeModelOverride | undefined,
): RuntimeModelOverride {
  const base = override ?? getModelSettingsForFeature(feature);
  return { ...base, reasoningEffort: downshiftReasoningEffort(base.reasoningEffort) };
}

export interface CompactRetryOptions<T> {
  /**
   * The feature whose configured model the **transport** resolves when the
   * caller passes no override. It must match the feature the underlying call
   * uses (`callAzureOpenAI` always resolves `architectureGeneration`), because
   * a retry has to keep the model that the first attempt actually ran on —
   * silently switching models on retry could land on one the caller's UI has
   * already ruled out as incompatible.
   */
  transportFeature: FeatureType;
  override?: RuntimeModelOverride;
  /** Prefix of the wrapped error, e.g. "Blueprint generation". */
  label: string;
  /**
   * Runs the generation. `compact` asks the caller to drop few-shot exemplars
   * and cap the output size so the request finishes inside the time budget.
   */
  attempt: (compact: boolean, override?: RuntimeModelOverride) => Promise<T>;
  /**
   * Extra classification for payload-level failures the transport cannot see
   * (truncated or non-JSON responses), which a compact retry usually fixes.
   */
  isRetryable?: (error: unknown) => boolean;
}

/**
 * Run a generation with exactly one automatic fallback attempt.
 *
 * The dominant failure for the long editorial prompts is exceeding the proxy's
 * upstream budget, which reaches the user as "The AI provider is taking too
 * long to respond." Retrying once with a compact prompt at low reasoning effort
 * finishes well inside the budget. Non-transient failures (auth, quota,
 * content policy, configuration) are rethrown untouched so the user is not
 * charged for a second pointless call.
 */
export async function runWithCompactRetry<T>(options: CompactRetryOptions<T>): Promise<T> {
  const { transportFeature, override, label, attempt, isRetryable } = options;
  try {
    return await attempt(false, override);
  } catch (error) {
    if (!isRetryableAIFailure(error) && !isRetryable?.(error)) throw error;
    console.warn(
      `⏱️ ${label} failed with a transient error — retrying once with a compact prompt at low reasoning effort:`,
      error,
    );
    try {
      return await attempt(true, buildRetryOverride(transportFeature, override));
    } catch (retryError) {
      // Surface the retry's reason but make clear a retry already happened, so
      // the user does not simply press the same button again.
      const detail = retryError instanceof Error ? retryError.message : String(retryError);
      throw new Error(`${label} failed after an automatic retry: ${detail}`);
    }
  }
}
