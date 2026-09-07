// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

/** Tab-local credentials and verification. This module never writes to storage. */
export interface BYOAIConnectionError {
  code: string;
  source: string;
  message: string;
  status?: number;
  requestId?: string;
}

export interface BYOAIConnectionState {
  status: 'missing-profile' | 'key-required' | 'unverified' | 'testing' | 'verified' | 'failed';
  hasApiKey: boolean;
  verified: boolean;
  revision: number;
  error?: BYOAIConnectionError;
}

interface ConnectionSession {
  apiKey: string;
  revision: number;
  verified: boolean;
  error?: BYOAIConnectionError;
  test?: { token: object; controller: AbortController };
}

let nextRevision = 0;
const sessions = new Map<string, ConnectionSession>();
const listeners = new Set<() => void>();
const notify = () => listeners.forEach(listener => listener());

function sessionFor(id: string): ConnectionSession {
  let session = sessions.get(id);
  if (!session) {
    session = { apiKey: '', revision: ++nextRevision, verified: false };
    sessions.set(id, session);
  }
  return session;
}

export function subscribeBYOAIConnectionSessions(listener: () => void): () => void {
  listeners.add(listener);
  return () => { listeners.delete(listener); };
}

export function isValidBYOAIApiKey(key: unknown): key is string {
  return typeof key === 'string' && /^[\x21-\x7e]{8,512}$/.test(key);
}

export function getBYOAIConnectionSessionState(id: string): BYOAIConnectionState {
  const session = sessionFor(id);
  const hasApiKey = isValidBYOAIApiKey(session.apiKey);
  return {
    status: !hasApiKey ? 'key-required' : session.test ? 'testing'
      : session.verified ? 'verified' : session.error ? 'failed' : 'unverified',
    hasApiKey,
    verified: hasApiKey && session.verified && !session.test,
    revision: session.revision,
    ...(session.error ? { error: { ...session.error } } : {}),
  };
}

export function invalidateBYOAIConnection(id: string): void {
  const session = sessionFor(id);
  session.test?.controller.abort();
  session.test = undefined;
  session.error = undefined;
  session.verified = false;
  session.revision = ++nextRevision;
  notify();
}

export function setBYOAIConnectionSecret(id: string, apiKey: string): void {
  const session = sessionFor(id);
  if (session.apiKey === apiKey) return;
  session.apiKey = apiKey;
  invalidateBYOAIConnection(id);
}

/** Internal transport access only; never re-export from the public settings store. */
export function readBYOAIConnectionSecret(id: string, revision: number): string | undefined {
  const session = sessions.get(id);
  return session?.revision === revision ? session.apiKey : undefined;
}

export function removeBYOAIConnectionSession(id: string): void {
  sessions.get(id)?.test?.controller.abort();
  sessions.delete(id);
  notify();
}

export function clearBYOAIConnectionSessions(): void {
  for (const session of sessions.values()) session.test?.controller.abort();
  sessions.clear();
  notify();
}

export function beginBYOAIConnectionTest(id: string, controller: AbortController): {
  token: object;
  revision: number;
} {
  const session = sessionFor(id);
  session.test?.controller.abort();
  session.verified = false;
  session.error = undefined;
  const token = Object.freeze({});
  session.test = { token, controller };
  notify();
  return { token, revision: session.revision };
}

export function isBYOAIConnectionTestCurrent(id: string, token: object, revision: number): boolean {
  const session = sessions.get(id);
  return session?.revision === revision && session.test?.token === token;
}

export function finishBYOAIConnectionTest(
  id: string,
  token: object,
  revision: number,
  result: { verified: boolean; error?: BYOAIConnectionError },
): boolean {
  if (!isBYOAIConnectionTestCurrent(id, token, revision)) return false;
  const session = sessions.get(id)!;
  if (result.verified && session.test!.controller.signal.aborted) return false;
  session.test = undefined;
  session.verified = result.verified;
  session.error = result.error ? { ...result.error } : undefined;
  notify();
  return true;
}
