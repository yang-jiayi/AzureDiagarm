// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

export interface AIConcurrencyBudget {
  concurrentLimit: number;
  concurrentRequests: number;
}

export type AIQueueTaskState<T> =
  | { status: 'pending' | 'running' }
  | { status: 'success'; value: T }
  | { status: 'error'; error: unknown };

export type AIBudgetQueueErrorCode =
  | 'invalid_configuration'
  | 'invalid_budget'
  | 'budget_unavailable'
  | 'capacity_timeout'
  | 'contention_limit';

export class AIBudgetQueueError extends Error {
  constructor(
    public readonly code: AIBudgetQueueErrorCode,
    message: string,
    public readonly cause?: unknown,
  ) {
    super(message);
    this.name = 'AIBudgetQueueError';
  }
}

export interface AIBudgetQueueOptions<T> {
  getBudget: (signal: AbortSignal) => Promise<unknown>;
  signal?: AbortSignal;
  onStateChange?: (index: number, state: AIQueueTaskState<T>) => void;
  pollIntervalMs?: number;
  maxCapacityWaitMs?: number;
  budgetRequestTimeoutMs?: number;
  maxContentionRetries?: number;
}

export function isAIConcurrencyLimitError(error: unknown): boolean {
  return !!error && typeof error === 'object'
    && 'code' in error && error.code === 'ai_concurrency_limit';
}

export function validateAIConcurrencyBudget(value: unknown): AIConcurrencyBudget {
  if (!value || typeof value !== 'object' || Array.isArray(value)
    || !('concurrentLimit' in value) || typeof value.concurrentLimit !== 'number'
    || !Number.isSafeInteger(value.concurrentLimit) || value.concurrentLimit < 1
    || !('concurrentRequests' in value) || typeof value.concurrentRequests !== 'number'
    || !Number.isSafeInteger(value.concurrentRequests) || value.concurrentRequests < 0) {
    throw new AIBudgetQueueError(
      'invalid_budget',
      'The AI concurrency budget is invalid. Check the server configuration.',
    );
  }
  return { concurrentLimit: value.concurrentLimit, concurrentRequests: value.concurrentRequests };
}

function cancellationError(): Error & { userCancelled: true } {
  return Object.assign(new Error('AI comparison cancelled.'), {
    name: 'AbortError',
    userCancelled: true as const,
  });
}

function throwIfAborted(signal: AbortSignal): void {
  if (signal.aborted) throw cancellationError();
}

function forwardAbort(source: AbortSignal | undefined, target: AbortController): () => void {
  if (!source) return () => {};
  const abort = () => target.abort();
  if (source.aborted) abort();
  else source.addEventListener('abort', abort, { once: true });
  return () => source.removeEventListener('abort', abort);
}

function abortable<T>(promise: Promise<T>, signal: AbortSignal): Promise<T> {
  return new Promise<T>((resolve, reject) => {
    const cleanup = () => signal.removeEventListener('abort', abort);
    const abort = () => {
      cleanup();
      reject(cancellationError());
    };
    // Observe the underlying operation even if it ignores cancellation and
    // resolves or rejects after the caller has closed the dialog.
    promise.then(
      value => {
        cleanup();
        if (signal.aborted) reject(cancellationError());
        else resolve(value);
      },
      error => {
        cleanup();
        reject(signal.aborted ? cancellationError() : error);
      },
    );
    if (signal.aborted) abort();
    else signal.addEventListener('abort', abort, { once: true });
  });
}

async function delay(milliseconds: number, signal: AbortSignal): Promise<void> {
  let timer: ReturnType<typeof setTimeout> | undefined;
  try {
    await abortable(new Promise<void>(resolve => {
      timer = setTimeout(resolve, milliseconds);
    }), signal);
  } finally {
    if (timer !== undefined) clearTimeout(timer);
  }
}

function duration(value: number, name: string): number {
  if (!Number.isSafeInteger(value) || value < 1 || value > 2_147_483_647) {
    throw new AIBudgetQueueError('invalid_configuration', `${name} must be a positive timer-safe integer.`);
  }
  return value;
}

/**
 * Admit work against both the server snapshot and local in-flight requests.
 * The server remains the final authority: admission races are requeued with
 * bounded backoff, not converted into cheaper model-generation attempts.
 */
export async function runAIBudgetQueue<T>(
  tasks: ReadonlyArray<(signal: AbortSignal) => Promise<T>>,
  options: AIBudgetQueueOptions<T>,
): Promise<PromiseSettledResult<T>[]> {
  const pollInterval = duration(options.pollIntervalMs ?? 1_000, 'pollIntervalMs');
  const capacityWait = duration(options.maxCapacityWaitMs ?? 60_000, 'maxCapacityWaitMs');
  const requestTimeout = duration(options.budgetRequestTimeoutMs ?? 10_000, 'budgetRequestTimeoutMs');
  const retryLimit = options.maxContentionRetries ?? 3;
  if (!Number.isSafeInteger(retryLimit) || retryLimit < 0) {
    throw new AIBudgetQueueError('invalid_configuration', 'maxContentionRetries must be a nonnegative safe integer.');
  }

  const controller = new AbortController();
  const signal = controller.signal;
  const unlinkAbort = forwardAbort(options.signal, controller);
  const jobs = [...tasks];
  const pending = jobs.map((_, index) => index);
  const results: Array<PromiseSettledResult<T> | undefined> = jobs.map(() => undefined);
  const contentionRetries = jobs.map(() => 0);
  const active = new Map<number, Promise<void>>();
  let limit = 1;
  let needsBackoff = false;
  let completed = false;
  let failure: { error: unknown } | undefined;
  let capacityTimer: ReturnType<typeof setTimeout> | undefined;

  const fail = (error: unknown) => {
    failure ??= { error };
    controller.abort();
  };
  const clearCapacityTimer = () => {
    if (capacityTimer !== undefined) clearTimeout(capacityTimer);
    capacityTimer = undefined;
  };
  const startCapacityTimer = () => {
    if (capacityTimer !== undefined) return;
    capacityTimer = setTimeout(() => fail(new AIBudgetQueueError(
      'capacity_timeout',
      'AI request capacity stayed busy. Try the comparison again when capacity is available.',
    )), capacityWait);
  };
  const publish = (index: number, state: AIQueueTaskState<T>) => {
    if (!signal.aborted && !failure) options.onStateChange?.(index, state);
  };
  const readBudget = async (): Promise<AIConcurrencyBudget> => {
    const request = new AbortController();
    const unlinkRequest = forwardAbort(signal, request);
    const timer = setTimeout(() => request.abort(), requestTimeout);
    let value: unknown;
    try {
      value = await abortable(Promise.resolve().then(() => {
        throwIfAborted(request.signal);
        return options.getBudget(request.signal);
      }), request.signal);
    } catch (error) {
      throwIfAborted(signal);
      throw new AIBudgetQueueError(
        'budget_unavailable',
        'The AI concurrency budget could not be checked. Please try again.',
        error,
      );
    } finally {
      clearTimeout(timer);
      unlinkRequest();
    }
    return validateAIConcurrencyBudget(value);
  };
  const start = (index: number) => {
    publish(index, { status: 'running' });
    const run = async () => {
      let result: PromiseSettledResult<T>;
      try {
        const value = await abortable(Promise.resolve().then(() => {
          throwIfAborted(signal);
          return jobs[index](signal);
        }), signal);
        result = { status: 'fulfilled', value };
      } catch (error) {
        if (signal.aborted) return;
        if (isAIConcurrencyLimitError(error) && contentionRetries[index] < retryLimit) {
          contentionRetries[index] += 1;
          pending.push(index);
          needsBackoff = true;
          publish(index, { status: 'pending' });
          return;
        }
        result = {
          status: 'rejected',
          reason: isAIConcurrencyLimitError(error)
            ? new AIBudgetQueueError(
                'contention_limit',
                'AI capacity was repeatedly claimed by other requests. Try this model again.',
                error,
              )
            : error,
        };
      }
      if (signal.aborted) return;
      results[index] = result;
      publish(index, result.status === 'fulfilled'
        ? { status: 'success', value: result.value }
        : { status: 'error', error: result.reason });
    };
    const work = run().catch(fail).finally(() => active.delete(index));
    active.set(index, work);
  };

  try {
    throwIfAborted(signal);
    pending.forEach(index => publish(index, { status: 'pending' }));
    while (pending.length || active.size) {
      throwIfAborted(signal);
      if (!pending.length || active.size >= limit) {
        await abortable(Promise.race(active.values()), signal);
        continue;
      }
      if (needsBackoff) {
        needsBackoff = false;
        await delay(pollInterval, signal);
      }
      const budget = await readBudget();
      throwIfAborted(signal);
      limit = budget.concurrentLimit;
      const slots = Math.max(0, Math.min(
        pending.length,
        limit - active.size,
        limit - budget.concurrentRequests,
      ));
      if (slots === 0) {
        if (active.size) {
          // Our own provider request can release capacity; it already has a
          // transport timeout. Do not poll or expire queued work while it runs.
          clearCapacityTimer();
          await abortable(Promise.race(active.values()), signal);
        } else {
          startCapacityTimer();
          await delay(pollInterval, signal);
        }
        continue;
      }
      clearCapacityTimer();
      for (let count = 0; count < slots && !signal.aborted; count += 1) {
        const index = pending.shift();
        if (index === undefined) break;
        start(index);
      }
    }
    throwIfAborted(signal);
    const settled = results.map(result => {
      if (!result) throw new Error('AI comparison queue did not settle every task.');
      return result;
    });
    completed = true;
    return settled;
  } catch (error) {
    throw failure ? failure.error : error;
  } finally {
    clearCapacityTimer();
    unlinkAbort();
    if (!completed) controller.abort();
  }
}
