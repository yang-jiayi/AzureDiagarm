// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

import { useEffect, useState } from 'react';

export type RuntimeConfigStatus = 'unknown' | 'loading' | 'ready' | 'error';

export interface RuntimeConfigSnapshot {
  status: RuntimeConfigStatus;
  bringYourOwnAI: boolean;
  error?: string;
}

let currentSnapshot: RuntimeConfigSnapshot = {
  status: 'unknown',
  bringYourOwnAI: false,
};
let pendingRequest: Promise<RuntimeConfigSnapshot> | null = null;
const listeners = new Set<() => void>();

function publish(snapshot: RuntimeConfigSnapshot): RuntimeConfigSnapshot {
  currentSnapshot = snapshot;
  listeners.forEach(listener => listener());
  return snapshot;
}

export function getRuntimeConfigSnapshot(): RuntimeConfigSnapshot {
  return { ...currentSnapshot };
}

export function isBYOAIEnabledOnServer(): boolean {
  return currentSnapshot.status === 'ready' && currentSnapshot.bringYourOwnAI;
}

export async function loadRuntimeConfig(force = false): Promise<RuntimeConfigSnapshot> {
  if (!force && currentSnapshot.status === 'ready') return getRuntimeConfigSnapshot();
  if (!force && pendingRequest) return pendingRequest;

  publish({ status: 'loading', bringYourOwnAI: false });
  pendingRequest = (async () => {
    try {
      const response = await fetch('/api/runtime-config', {
        method: 'GET',
        headers: { Accept: 'application/json' },
        cache: 'no-store',
      });
      if (!response.ok) {
        throw new Error(`Runtime configuration request failed (${response.status}).`);
      }
      const payload = await response.json() as {
        features?: { bringYourOwnAI?: unknown };
      };
      if (typeof payload.features?.bringYourOwnAI !== 'boolean') {
        throw new Error('Runtime configuration response is invalid.');
      }
      return publish({
        status: 'ready',
        bringYourOwnAI: payload.features.bringYourOwnAI,
      });
    } catch (error) {
      return publish({
        status: 'error',
        bringYourOwnAI: false,
        error: error instanceof Error ? error.message : 'Runtime configuration is unavailable.',
      });
    } finally {
      pendingRequest = null;
    }
  })();
  return pendingRequest;
}

export function useRuntimeConfig(): RuntimeConfigSnapshot {
  const [snapshot, setSnapshot] = useState(getRuntimeConfigSnapshot);

  useEffect(() => {
    const listener = () => setSnapshot(getRuntimeConfigSnapshot());
    listeners.add(listener);
    void loadRuntimeConfig();
    return () => {
      listeners.delete(listener);
    };
  }, []);

  return snapshot;
}

export function resetRuntimeConfigForTests(): void {
  pendingRequest = null;
  publish({ status: 'unknown', bringYourOwnAI: false });
}
