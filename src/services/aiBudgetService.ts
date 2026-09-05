// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.
export interface AIBudget {
  available: true;
  limitTokens: number;
  usedTokens: number;
  reservedTokens: number;
  remainingTokens: number;
  concurrentRequests: number;
  concurrentLimit: number;
  resetAt: string;
  mode: 'public' | 'local';
}

export async function getAIBudget(signal?: AbortSignal): Promise<AIBudget> {
  const response = await fetch('/api/ai/budget', {
    credentials: 'same-origin', cache: 'no-store', signal,
  });
  if (!response.ok) throw new Error('AI budget unavailable.');
  const data = await response.json() as AIBudget;
  if (data.available !== true || !Number.isFinite(data.remainingTokens)
    || !Number.isFinite(data.limitTokens) || !Number.isFinite(data.concurrentRequests)
    || !Number.isFinite(data.concurrentLimit) || !Number.isFinite(Date.parse(data.resetAt))) {
    throw new Error('Invalid AI budget response.');
  }
  return data;
}
