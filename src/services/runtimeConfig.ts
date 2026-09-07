// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

import { useEffect, useState } from 'react';

export type RuntimeConfigStatus = 'unknown' | 'loading' | 'ready' | 'error';
export interface RuntimeConfigSnapshot {
  status: RuntimeConfigStatus;
  bringYourOwnAI: boolean;
  error?: string;
}

let currentSnapshot: RuntimeConfigSnapshot = { status: 'unknown', bringYourOwnAI: false };
let pendingRequest: Promise<RuntimeConfigSnapshot> | null = null;
let requestRevision = 0;
let lastResultRevision = 0;
const listeners = new Set<() => void>();

function publish(value: RuntimeConfigSnapshot): RuntimeConfigSnapshot {
  currentSnapshot = { ...value };
  listeners.forEach(listener => listener());
  return { ...currentSnapshot };
}

export function getRuntimeConfigSnapshot(): RuntimeConfigSnapshot { return { ...currentSnapshot }; }
export function isBYOAIEnabledOnServer(): boolean {
  return currentSnapshot.status === 'ready' && currentSnapshot.bringYourOwnAI;
}

export function runtimeConfigCancellationError(): Error & { userCancelled: true } {
  return Object.assign(new DOMException('Connection check cancelled.', 'AbortError'), { userCancelled: true as const });
}

/** Also bounds mocks/transports that ignore AbortSignal, without unhandled late replies. */
export function awaitWithAISignal<T>(operation: Promise<T>, signal: AbortSignal): Promise<T> {
  return new Promise<T>((resolve, reject) => {
    const abort = () => { cleanup(); reject(runtimeConfigCancellationError()); };
    const cleanup = () => signal.removeEventListener('abort', abort);
    operation.then(value => {
      cleanup();
      if (signal.aborted) reject(runtimeConfigCancellationError());
      else resolve(value);
    }, error => { cleanup(); reject(signal.aborted ? runtimeConfigCancellationError() : error); });
    if (signal.aborted) abort();
    else signal.addEventListener('abort', abort, { once: true });
  });
}

export async function loadRuntimeConfig(
  force = false,
  options: { signal?: AbortSignal } = {},
): Promise<RuntimeConfigSnapshot> {
  if (options.signal?.aborted) throw runtimeConfigCancellationError();
  if (!force && currentSnapshot.status === 'ready') return getRuntimeConfigSnapshot();
  if (!force && pendingRequest) {
    return options.signal ? awaitWithAISignal(pendingRequest, options.signal) : pendingRequest;
  }
  const revision = ++requestRevision;
  const controller = new AbortController();
  const abort = () => controller.abort();
  options.signal?.addEventListener('abort', abort, { once: true });
  const timer = setTimeout(abort, 10_000);
  publish({ status: 'loading', bringYourOwnAI: false });
  const work = (async () => {
    try {
      const response = await awaitWithAISignal(fetch('/api/runtime-config', {
        method: 'GET', headers: { Accept: 'application/json' }, cache: 'no-store',
        credentials: 'same-origin', redirect: 'error', signal: controller.signal,
      }), controller.signal);
      if (!response.ok || response.redirected
        || !/^application\/(?:[a-z0-9!#$&^_.+-]+\+)?json(?:\s*;|$)/i.test(response.headers.get('content-type') || '')) {
        throw new Error('Runtime configuration is unavailable.');
      }
      const payload = await awaitWithAISignal(response.json(), controller.signal) as { features?: { bringYourOwnAI?: unknown } };
      if (typeof payload?.features?.bringYourOwnAI !== 'boolean') throw new Error('Invalid runtime configuration.');
      const result: RuntimeConfigSnapshot = { status: 'ready', bringYourOwnAI: payload.features.bringYourOwnAI };
      if (revision > lastResultRevision) {
        lastResultRevision = revision;
        return publish(result);
      }
      return getRuntimeConfigSnapshot();
    } catch {
      const result: RuntimeConfigSnapshot = {
        status: 'error', bringYourOwnAI: false, error: 'The application server could not confirm bring-your-own AI availability.',
      };
      if (revision > lastResultRevision) {
        lastResultRevision = revision;
        publish(result);
      }
      if (options.signal?.aborted) throw runtimeConfigCancellationError();
      return getRuntimeConfigSnapshot();
    } finally {
      clearTimeout(timer);
      options.signal?.removeEventListener('abort', abort);
      if (revision === requestRevision) pendingRequest = null;
    }
  })();
  pendingRequest = work;
  return work;
}

export function useRuntimeConfig(): RuntimeConfigSnapshot {
  const [value, setValue] = useState(getRuntimeConfigSnapshot);
  useEffect(() => {
    const listener = () => setValue(getRuntimeConfigSnapshot());
    listeners.add(listener);
    listener();
    void loadRuntimeConfig();
    return () => { listeners.delete(listener); };
  }, []);
  return value;
}

export function resetRuntimeConfigForTests(): void {
  requestRevision++;
  lastResultRevision = requestRevision;
  pendingRequest = null;
  publish({ status: 'unknown', bringYourOwnAI: false });
}
