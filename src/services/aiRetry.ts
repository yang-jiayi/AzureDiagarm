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
 * Classification for a model response that could not be turned into JSON.
 * Callers use this to decide whether a compact retry is worthwhile and how to
 * phrase the failure to the user.
 */
export enum ModelJsonErrorKind {
  /** The model returned nothing usable (empty / whitespace only). */
  Empty = 'empty',
  /** The model produced a prose refusal or content decline instead of JSON. */
  Refusal = 'refusal',
  /** The JSON started but was cut off before it closed (unbalanced braces). */
  Truncated = 'truncated',
  /** The output was present but is not valid JSON for some other reason. */
  Unparseable = 'unparseable',
}

/**
 * Typed error thrown by {@link safeParseModelJson}. The raw model text and the
 * underlying parser message are kept on `detail` for logging only — they must
 * never be surfaced to the user, whose message is the localisation-friendly
 * `.message`.
 */
export class ModelJsonError extends Error {
  readonly kind: ModelJsonErrorKind;
  /** Whether a compact retry could plausibly fix this (truncated/empty yes). */
  readonly retryable: boolean;
  /** Raw model text / parser detail. Console-only — never shown to the user. */
  readonly detail?: string;

  constructor(
    kind: ModelJsonErrorKind,
    message: string,
    options: { retryable?: boolean; detail?: string } = {},
  ) {
    super(message);
    this.name = 'ModelJsonError';
    this.kind = kind;
    this.retryable = options.retryable ?? false;
    this.detail = options.detail;
  }
}

/**
 * Stable, localisation-friendly user messages for each JSON failure kind. These
 * are keyed verbatim by the Japanese dictionary in `LanguageContext`, so any
 * change here must be mirrored there (the i18n coverage test enforces this).
 */
export const MODEL_JSON_ERROR_MESSAGES: Readonly<Record<ModelJsonErrorKind, string>> = {
  [ModelJsonErrorKind.Empty]: 'The AI model returned an empty response. Please try again.',
  [ModelJsonErrorKind.Refusal]:
    'The AI model declined to complete this request. Revise the prompt and try again.',
  [ModelJsonErrorKind.Truncated]:
    'The AI response was cut off before it finished. Please try again.',
  [ModelJsonErrorKind.Unparseable]:
    'The AI model returned a response that was not valid JSON. Please try again.',
};

/**
 * True when the failure is worth one more (cheaper) attempt. Authentication,
 * authorization, configuration, and content-policy failures are deliberately
 * excluded — retrying those only wastes the user's time.
 */
export function isRetryableAIFailure(error: unknown): boolean {
  if (!error) return false;

  // A user-initiated cancellation is terminal: never retry it. It surfaces as
  // an AbortError (which for an INTERNAL timeout is retryable), so it must be
  // distinguished by an explicit flag, not by the error name.
  if ((error as { userCancelled?: unknown }).userCancelled === true) return false;

  // A truncated or empty JSON payload is exactly what a compact retry fixes;
  // a refusal or otherwise malformed payload is not worth a second charge.
  if (error instanceof ModelJsonError) return error.retryable;

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
      // Preserve the ORIGINAL typed error so downstream UI can still classify
      // and localise it (e.g. an OpenAIProxyError keeps its `.code`, a
      // ModelJsonError keeps its `.kind`). Flattening it to a generic Error
      // here would strip that classification and leak the raw detail into the
      // user-facing string. We only annotate `retried` so the UI can avoid
      // inviting the user to press the same button again.
      console.error(`${label} failed after an automatic retry:`, retryError);
      if (retryError && typeof retryError === 'object') {
        try {
          (retryError as { retried?: boolean }).retried = true;
        } catch {
          /* frozen error object — the annotation is best-effort */
        }
      }
      throw retryError;
    }
  }
}

// ── Fence-tolerant, refusal-aware model JSON parsing ────────────────────────
//
// Models — especially user-supplied BYO endpoints that ignore json_object mode
// — routinely wrap JSON in ```json fences, prepend a sentence of prose, or
// return a plain-text refusal. A raw `JSON.parse` on any of those throws a
// parser message that is unlocalised, leaks internals, and gives the user no
// idea what to do. `safeParseModelJson` normalises all of that into a small set
// of typed, localisable outcomes.

/** Options for {@link safeParseModelJson}. */
export interface SafeParseOptions {
  /** Short human context for console diagnostics, e.g. "architecture generation". */
  context?: string;
  /** Console sink override (defaults to the global console) — used by tests. */
  logger?: Pick<Console, 'error' | 'warn'>;
}

/** Refusal / content-decline phrases models emit instead of JSON. */
const REFUSAL_PATTERNS: readonly RegExp[] = [
  /\b(i'?m sorry|i am sorry|i apologi[sz]e)\b/i,
  /\bi (can'?t|cannot|can not|am unable to|won'?t be able to|am not able to)\b/i,
  /\b(unable|not able) to (help|assist|comply|process|complete|create|generate|provide|fulfil)/i,
  /\b(can'?t|cannot) (help|assist|comply|process|complete|create|generate|provide|fulfil)/i,
  /\bas an ai\b/i,
  /\b(content|usage) polic(y|ies)\b/i,
  /\bi'?m (not able|unable)\b/i,
];

/** Strip a single leading ```lang fence and its trailing ``` (if present). */
function stripCodeFences(input: string): string {
  let out = input.trim();
  if (out.startsWith('```')) {
    const firstNewline = out.indexOf('\n');
    out = firstNewline === -1 ? '' : out.slice(firstNewline + 1);
    const lastFence = out.lastIndexOf('```');
    if (lastFence !== -1) out = out.slice(0, lastFence);
  }
  return out.trim();
}

/**
 * Return the outermost balanced `{...}` or `[...]` block, or null if the text
 * has no JSON start or never closes (i.e. it was truncated). String contents
 * (and escapes) are respected so a brace inside a string does not skew depth.
 */
function extractBalancedJson(input: string): string | null {
  let start = -1;
  for (let i = 0; i < input.length; i++) {
    if (input[i] === '{' || input[i] === '[') {
      start = i;
      break;
    }
  }
  if (start === -1) return null;

  const open = input[start];
  const close = open === '{' ? '}' : ']';
  let depth = 0;
  let inString = false;
  let escaped = false;

  for (let i = start; i < input.length; i++) {
    const ch = input[i];
    if (inString) {
      if (escaped) escaped = false;
      else if (ch === '\\') escaped = true;
      else if (ch === '"') inString = false;
      continue;
    }
    if (ch === '"') {
      inString = true;
      continue;
    }
    if (ch === open) depth++;
    else if (ch === close) {
      depth--;
      if (depth === 0) return input.slice(start, i + 1);
    }
  }
  return null; // opened but never balanced → truncated
}

function hasJsonStart(input: string): boolean {
  return /[{[]/.test(input);
}

function looksLikeRefusal(input: string): boolean {
  return REFUSAL_PATTERNS.some((pattern) => pattern.test(input));
}

/**
 * A usable model result is always a JSON object or array — callers dereference
 * `.services` / `.overallScore` / `.timestamp` etc. A bare primitive (`null`,
 * `123`, `"text"`, `true`) is not usable and, if returned, only defers the
 * crash into an unlocalised `TypeError` downstream.
 */
function isJsonContainer(value: unknown): boolean {
  return value !== null && typeof value === 'object';
}

/** True only for a NON-empty object or array (`{}` / `[]` do not qualify). */
function isNonEmptyJsonContainer(value: unknown): boolean {
  if (Array.isArray(value)) return value.length > 0;
  if (value !== null && typeof value === 'object') return Object.keys(value).length > 0;
  return false;
}

/**
 * Parse a model's textual response into JSON, tolerating the ways models drift
 * from strict JSON, and classifying the failures a raw `JSON.parse` cannot.
 *
 * Behaviour:
 *  - strips leading/trailing ```json / ``` fences and surrounding prose,
 *  - falls back to the outermost balanced `{...}` / `[...]` block,
 *  - throws a distinct, NON-retryable {@link ModelJsonError} for a prose refusal,
 *  - throws a distinct, RETRYABLE {@link ModelJsonError} for empty or truncated
 *    (unbalanced) output,
 *  - never puts the raw parser message in the thrown `.message` (it is logged to
 *    the console and preserved on `.detail` instead).
 */
export function safeParseModelJson<T = unknown>(
  raw: unknown,
  options: SafeParseOptions = {},
): T {
  const log = options.logger ?? console;
  const where = options.context ? ` (${options.context})` : '';
  const text = typeof raw === 'string' ? raw : String(raw ?? '');
  const trimmed = text.trim();

  if (!trimmed) {
    log.error(`safeParseModelJson: empty model response${where}`);
    throw new ModelJsonError(ModelJsonErrorKind.Empty, MODEL_JSON_ERROR_MESSAGES.empty, {
      retryable: true,
    });
  }

  const stripped = stripCodeFences(trimmed);

  // A refusal can wrap (or sit beside) a stray bracket, so classify the prose
  // ONCE up front and consult it wherever a bracket would otherwise be trusted.
  const refusalContext = looksLikeRefusal(stripped);

  // Fast path: the whole (de-fenced) payload is valid JSON.
  try {
    const direct = JSON.parse(stripped);
    // A bare primitive is not a usable result — reject it here rather than
    // letting a caller crash on `guide.timestamp = …` / `validation.overallScore`.
    if (!isJsonContainer(direct)) {
      log.error(`safeParseModelJson: response was a bare JSON primitive${where}`);
      throw new ModelJsonError(ModelJsonErrorKind.Unparseable, MODEL_JSON_ERROR_MESSAGES.unparseable, {
        retryable: true,
        detail: stripped.slice(0, 2_000),
      });
    }
    return direct as T;
  } catch (directError) {
    // Our own classification must not be swallowed by this catch.
    if (directError instanceof ModelJsonError) throw directError;

    // Fall back to pulling a balanced block out of surrounding prose.
    const block = extractBalancedJson(stripped);
    if (block) {
      let parsedBlock: unknown;
      try {
        parsedBlock = JSON.parse(block);
      } catch (blockError) {
        // A stray, unparsable brace inside refusal prose (e.g. "…can't create
        // that. For example {foo}.") must NOT become a retryable Unparseable —
        // that fires a pointless second request that will refuse again.
        if (refusalContext) {
          log.warn(`safeParseModelJson: refusal with a stray unparsable brace${where}:`, stripped.slice(0, 500));
          throw new ModelJsonError(ModelJsonErrorKind.Refusal, MODEL_JSON_ERROR_MESSAGES.refusal, {
            retryable: false,
            detail: stripped.slice(0, 2_000),
          });
        }
        log.error(
          `safeParseModelJson: extracted JSON block failed to parse${where}:`,
          (blockError as Error)?.message,
        );
        throw new ModelJsonError(
          ModelJsonErrorKind.Unparseable,
          MODEL_JSON_ERROR_MESSAGES.unparseable,
          { retryable: true, detail: block.slice(0, 2_000) },
        );
      }

      // The block parsed. When the surrounding prose reads as a refusal, only
      // trust it if it is a NON-EMPTY container that makes up the bulk of the
      // response — a tiny "{ }" or "[1]" lifted from an apology is incidental
      // prose, not data, and returning it would both hide the localized "the
      // model declined" message and (on the architecture path) trip a further
      // retry via EmptyArchitectureError.
      const blockDominates = block.length >= stripped.length * 0.5;
      if (refusalContext && !(isNonEmptyJsonContainer(parsedBlock) && blockDominates)) {
        log.warn(`safeParseModelJson: refusal prose around an incidental brace${where}:`, stripped.slice(0, 500));
        throw new ModelJsonError(ModelJsonErrorKind.Refusal, MODEL_JSON_ERROR_MESSAGES.refusal, {
          retryable: false,
          detail: stripped.slice(0, 2_000),
        });
      }

      // A balanced block always starts with `{`/`[`, but guard the container
      // invariant explicitly so a caller never receives a primitive.
      if (!isJsonContainer(parsedBlock)) {
        log.error(`safeParseModelJson: extracted block was not an object or array${where}`);
        throw new ModelJsonError(ModelJsonErrorKind.Unparseable, MODEL_JSON_ERROR_MESSAGES.unparseable, {
          retryable: true,
          detail: block.slice(0, 2_000),
        });
      }

      return parsedBlock as T;
    }

    // No balanced block. A pure prose refusal (no JSON start) is not worth a
    // retry; anything that opened a brace but never closed it was truncated.
    if (!hasJsonStart(stripped) && refusalContext) {
      log.warn(`safeParseModelJson: model returned a refusal${where}:`, stripped.slice(0, 500));
      throw new ModelJsonError(ModelJsonErrorKind.Refusal, MODEL_JSON_ERROR_MESSAGES.refusal, {
        retryable: false,
        detail: stripped.slice(0, 2_000),
      });
    }

    if (hasJsonStart(stripped)) {
      log.error(
        `safeParseModelJson: truncated / unbalanced JSON${where}:`,
        (directError as Error)?.message,
      );
      throw new ModelJsonError(ModelJsonErrorKind.Truncated, MODEL_JSON_ERROR_MESSAGES.truncated, {
        retryable: true,
        detail: stripped.slice(0, 2_000),
      });
    }

    log.error(
      `safeParseModelJson: response was not JSON${where}:`,
      (directError as Error)?.message,
    );
    throw new ModelJsonError(
      ModelJsonErrorKind.Unparseable,
      MODEL_JSON_ERROR_MESSAGES.unparseable,
      { retryable: true, detail: stripped.slice(0, 2_000) },
    );
  }
}
