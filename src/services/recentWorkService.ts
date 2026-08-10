// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

import type { CloudDiagramPayload } from './cloudDiagramService';

export type RecentWorkSyncState =
  | 'local'
  | 'saving'
  | 'synced'
  | 'readonly'
  | 'offline'
  | 'unavailable'
  | 'conflict'
  | 'error';

export interface RecentWorkPayload extends CloudDiagramPayload {
  viewport?: {
    x: number;
    y: number;
    zoom: number;
  };
}

export interface RecentWorkRecord {
  id: string;
  lineageId: string;
  sessionId: string;
  diagramName: string;
  updatedAt: number;
  payload: RecentWorkPayload;
  syncState: RecentWorkSyncState;
  cloudDocumentId?: string;
  cloudRevision?: number;
}

const DB_NAME = 'AzureDiagramRecentWork';
const STORE_NAME = 'work';
const DB_VERSION = 1;
const SESSION_KEY = 'azurediagarm.recent-work-session.v1';
const MAX_RECORDS = 12;
const MAX_AGE_MS = 45 * 24 * 60 * 60 * 1000;

function isRecord(value: unknown): value is Record<string, unknown> {
  return Boolean(value) && typeof value === 'object' && !Array.isArray(value);
}

function isRecentWorkRecord(value: unknown): value is RecentWorkRecord {
  if (!isRecord(value) || !isRecord(value.payload)) return false;
  return (
    typeof value.id === 'string'
    && value.id.length > 0
    && value.id.length <= 320
    && typeof value.lineageId === 'string'
    && value.lineageId.length > 0
    && value.lineageId.length <= 320
    && typeof value.sessionId === 'string'
    && value.sessionId.length > 0
    && value.sessionId.length <= 120
    && typeof value.diagramName === 'string'
    && value.diagramName.length <= 200
    && typeof value.updatedAt === 'number'
    && Number.isFinite(value.updatedAt)
    && Array.isArray(value.payload.nodes)
    && value.payload.nodes.length <= 2000
    && Array.isArray(value.payload.edges)
    && value.payload.edges.length <= 5000
  );
}

function openDatabase(): Promise<IDBDatabase> {
  if (typeof indexedDB === 'undefined') {
    return Promise.reject(new Error('Local recovery storage is unavailable.'));
  }

  return new Promise((resolve, reject) => {
    const request = indexedDB.open(DB_NAME, DB_VERSION);
    request.onerror = () => reject(request.error || new Error('Local recovery storage could not be opened.'));
    request.onblocked = () => reject(new Error('Local recovery storage is blocked by another tab.'));
    request.onsuccess = () => resolve(request.result);
    request.onupgradeneeded = () => {
      const database = request.result;
      if (database.objectStoreNames.contains(STORE_NAME)) return;
      const store = database.createObjectStore(STORE_NAME, { keyPath: 'id' });
      store.createIndex('updatedAt', 'updatedAt', { unique: false });
      store.createIndex('lineageId', 'lineageId', { unique: false });
    };
  });
}

async function readAllRecords(): Promise<unknown[]> {
  const database = await openDatabase();
  try {
    return await new Promise((resolve, reject) => {
      const transaction = database.transaction(STORE_NAME, 'readonly');
      const request = transaction.objectStore(STORE_NAME).getAll();
      request.onerror = () => reject(request.error || new Error('Recent work could not be read.'));
      request.onsuccess = () => resolve(request.result);
    });
  } finally {
    database.close();
  }
}

async function deleteRecordIds(ids: string[]): Promise<void> {
  if (ids.length === 0) return;
  const database = await openDatabase();
  try {
    await new Promise<void>((resolve, reject) => {
      const transaction = database.transaction(STORE_NAME, 'readwrite');
      const store = transaction.objectStore(STORE_NAME);
      ids.forEach(id => store.delete(id));
      transaction.onerror = () => reject(transaction.error || new Error('Recent work could not be pruned.'));
      transaction.onabort = () => reject(transaction.error || new Error('Recent work pruning was aborted.'));
      transaction.oncomplete = () => resolve();
    });
  } finally {
    database.close();
  }
}

/**
 * Random suffix for a session ID. Uses the platform CSPRNG; `Math.random` is
 * deliberately avoided so the value is never a predictable identifier.
 */
function randomSessionSuffix(): string {
  if (typeof crypto !== 'undefined' && typeof crypto.getRandomValues === 'function') {
    const bytes = new Uint8Array(16);
    crypto.getRandomValues(bytes);
    return Array.from(bytes, (byte) => byte.toString(16).padStart(2, '0')).join('');
  }
  // No CSPRNG at all: fall back to a high-resolution timestamp. This only
  // groups one browser session's own records and is never used as a token.
  const monotonic = typeof performance !== 'undefined' ? performance.now() : 0;
  return `${Date.now().toString(36)}-${Math.trunc(monotonic * 1000).toString(36)}`;
}

function createSessionId(): string {
  if (typeof crypto !== 'undefined' && typeof crypto.randomUUID === 'function') {
    return crypto.randomUUID();
  }
  return `session-${randomSessionSuffix()}`;
}

export function getRecentWorkSessionId(): string {
  try {
    const stored = window.sessionStorage.getItem(SESSION_KEY);
    if (stored && stored.length <= 120) return stored;
    const next = createSessionId();
    window.sessionStorage.setItem(SESSION_KEY, next);
    return next;
  } catch {
    return createSessionId();
  }
}

export function isRecentWorkUnsynced(record: RecentWorkRecord): boolean {
  return (
    record.syncState === 'saving'
    || record.syncState === 'offline'
    || record.syncState === 'unavailable'
    || record.syncState === 'conflict'
    || record.syncState === 'error'
  );
}

export async function listRecentWork(): Promise<RecentWorkRecord[]> {
  const cutoff = Date.now() - MAX_AGE_MS;
  const rawRecords = await readAllRecords();
  const validRecords = rawRecords
    .filter(isRecentWorkRecord)
    .sort((left, right) => right.updatedAt - left.updatedAt);
  const retained = validRecords
    .filter(record => record.updatedAt >= cutoff)
    .slice(0, MAX_RECORDS);
  const retainedIds = new Set(retained.map(record => record.id));
  const staleIds = validRecords
    .filter(record => !retainedIds.has(record.id))
    .map(record => record.id);
  await deleteRecordIds(staleIds);
  return retained;
}

export async function saveRecentWork(record: RecentWorkRecord): Promise<void> {
  if (!isRecentWorkRecord(record)) {
    throw new Error('The current diagram is not valid for local recovery.');
  }
  const database = await openDatabase();
  try {
    await new Promise<void>((resolve, reject) => {
      const transaction = database.transaction(STORE_NAME, 'readwrite');
      transaction.objectStore(STORE_NAME).put(record);
      transaction.onerror = () => reject(transaction.error || new Error('Recent work could not be saved.'));
      transaction.onabort = () => reject(transaction.error || new Error('Recent work saving was aborted.'));
      transaction.oncomplete = () => resolve();
    });
  } finally {
    database.close();
  }
  await listRecentWork();
}

export async function deleteRecentWork(id: string): Promise<void> {
  await deleteRecordIds([id]);
}
