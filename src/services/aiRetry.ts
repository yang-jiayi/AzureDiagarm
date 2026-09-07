// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

/**
 * Shared retry classification for AI generation calls.
 *
 * Classified throttles use bounded, cancellable waits with identical requests.
 * Timeouts and invalid outputs remain explicit failures; automatic retries
 * must never silently lower the user's model, reasoning, or output quality.
 */

import type { FeatureType } from '../stores/modelSettingsStore';
import type { RuntimeModelOverride } from './aiModelRuntime';
import { waitForAIRetry } from './aiBudgetQueue';
import { isAIBudgetError } from './apiHelper';

const RATE_LIMIT_CODES = new Set([
  'azure_openai_rate_limited', 'proxy_rate_limit_exceeded', 'byo_rate_limited',
]);

/** Admission, authentication and daily budgets must not become provider retries. */
export function isAIRateLimitError(error: unknown): boolean {
  if (!error || typeof error !== 'object') return false;
  const value = error as { code?: unknown; userCancelled?: unknown };
  if (value.userCancelled === true || isAIBudgetError(error)) return false;
  if (typeof value.code === 'string' && RATE_LIMIT_CODES.has(value.code)) return true;
  // An unclassified 429 can be budget exhaustion; do not infer an automatic retry.
  return false;
}

export interface AIRetryWait {
  /** The upcoming attempt, counting the initial request as attempt one. */
  attempt: number;
  maxAttempts: number;
  delayMs: number;
  retryAt: number;
}

export interface AIRateLimitRetryOptions {
  signal?: AbortSignal;
  onRetryWait?: (wait: AIRetryWait | null) => void;
}

/**
 * Repeat only rejected rate-limited transport requests, with the identical body.
 * At most two retries / 120 seconds of cooldown; a longer provider delay is
 * surfaced rather than shortened. Missing/invalid headers use 30s, then 60s.
 * Network timeouts remain the responsibility of each individual attempt.
 */
export async function runWithRateLimitRetry<T>(
  attempt: () => Promise<T>,
  options: AIRateLimitRetryOptions = {},
): Promise<T> {
  const signal = options.signal ?? new AbortController().signal;
  const assertActive = () => {
    if (signal.aborted) {
      throw Object.assign(new DOMException('Generation cancelled.', 'AbortError'), { userCancelled: true });
    }
  };
  let waitedMs = 0;
  for (let index = 0; ; index += 1) {
    assertActive();
    try {
      const result = await attempt();
      assertActive();
      return result;
    } catch (error) {
      assertActive();
      if (!isAIRateLimitError(error) || index >= 2) throw error;
      const suppliedDelay = (error as { retryAfterMs?: number }).retryAfterMs;
      const delayMs = typeof suppliedDelay === 'number' && Number.isSafeInteger(suppliedDelay) && suppliedDelay >= 0
        ? Math.max(1000, suppliedDelay)
        : 30_000 * (2 ** index);
      if (waitedMs + delayMs > 120_000) throw error;
      waitedMs += delayMs;
      options.onRetryWait?.({ attempt: index + 2, maxAttempts: 3, delayMs, retryAt: Date.now() + delayMs });
      try {
        await waitForAIRetry(delayMs, signal);
      } finally {
        // Cancellation owns UI cleanup; never publish a late "resuming" state.
        if (!signal.aborted) options.onRetryWait?.(null);
      }
    }
  }
}

/** Proxy error codes that represent a transient upstream/edge condition. */
const RETRYABLE_PROXY_CODES = new Set([
  'azure_openai_timeout',
  'azure_openai_unavailable',
  'azure_openai_connection_failed',
  'edge_origin_unavailable',
  'proxy_rate_limit_exceeded',
  'azure_openai_rate_limited',
  'byo_rate_limited',
  'byo_timeout',
  'byo_unavailable',
  'byo_connection_failed',
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
 * Callers use this to phrase the failure and offer an explicit retry.
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
  /** Whether a new user-requested attempt could plausibly recover this output. */
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
 * True when the failure is transient. Rate limits require their own wait policy,
 * not a compact fallback. Authentication,
 * authorization, configuration, and content-policy failures are deliberately
 * excluded — retrying those only wastes the user's time.
 */
export function isRetryableAIFailure(error: unknown): boolean {
  if (!error) return false;

  // A user-initiated cancellation is terminal: never retry it. It surfaces as
  // an AbortError (which for an INTERNAL timeout is retryable), so it must be
  // distinguished by an explicit flag, not by the error name.
  if ((error as { userCancelled?: unknown }).userCancelled === true) return false;

  // Capacity contention needs admission/backoff, not a cheaper generation.
  if (isAIBudgetError(error)) return false;

  // The classification informs explicit retries, never automatic compaction.
  if (error instanceof ModelJsonError) return error.retryable;

  const code = (error as { code?: unknown }).code;
  if (code === 'astra_not_configured' || code === 'proxy_not_configured'
    || code === 'stale_ai_configuration' || code === 'byo_not_enabled'
    || code === 'byo_availability_unknown' || code === 'byo_unverified'
    || code === 'byo_key_required' || code === 'byo_profile_missing'
    || code === 'byo_invalid_profile' || code === 'invalid_byo_configuration') return false;
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
   * Legacy compatibility field. A classifier no longer authorizes an automatic
   * compact retry; changing requested quality requires an explicit user action.
   */
  isRetryable?: (error: unknown) => boolean;
}

/**
 * Legacy entry point retained for caller compatibility. Never compact or lower
 * reasoning automatically: timeouts and invalid outputs require a user retry.
 * Classified throttles are retried by the transport with the identical body.
 */
export async function runWithCompactRetry<T>(options: CompactRetryOptions<T>): Promise<T> {
  return options.attempt(false, options.override);
}

// ── Fence-tolerant, refusal-aware model JSON parsing ────────────────────────
//
// Models can wrap JSON in ```json fences, prepend a sentence of prose, or
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
