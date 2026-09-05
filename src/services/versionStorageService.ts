// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

import type { PricingScenario } from '../types/pricing';
import type { IaCBaseline } from './iacRoundTrip';
import type { EditorDocument } from './editorHistory';

/**
 * Version Storage Service
 * Manages diagram version history using IndexedDB for local persistence
 */

export interface DiagramVersion {
  versionId: string;
  lineageId?: string;
  timestamp: number;
  diagramName: string;
  architecturePrompt?: string;
  /** The first prompt of the diagram lineage (survives chat refinements). */
  originalPrompt?: string;
  validationScore?: number;
  parentVersionId?: string;
  improvementsApplied?: string[];
  notes?: string;
  nodes: any[];
  edges: any[];
  metadata?: any;
  workflow?: any[];
  titleBlockData?: any;
  pricingScenarios?: PricingScenario[];
  iacBaseline?: IaCBaseline | null;
  settings?: EditorDocument['settings'];
  reviewHistory?: unknown[];
  validationSourceFingerprint?: string | null;
}

const DB_NAME = 'AzureDiagramVersions';
const STORE_NAME = 'versions';
const DB_VERSION = 1;

/**
 * Initialize IndexedDB database
 */
const initDB = (): Promise<IDBDatabase> => {
  return new Promise((resolve, reject) => {
    const request = indexedDB.open(DB_NAME, DB_VERSION);
    let blocked = false;

    request.onerror = () => reject(request.error);
    request.onblocked = () => {
      blocked = true;
      reject(new Error('Version storage is blocked by another tab.'));
    };
    request.onsuccess = () => {
      if (blocked) {
        request.result.close();
        return;
      }
      request.result.onversionchange = () => request.result.close();
      resolve(request.result);
    };

    request.onupgradeneeded = (event) => {
      const db = (event.target as IDBOpenDBRequest).result;
      
      if (!db.objectStoreNames.contains(STORE_NAME)) {
        const objectStore = db.createObjectStore(STORE_NAME, { keyPath: 'versionId' });
        objectStore.createIndex('timestamp', 'timestamp', { unique: false });
        objectStore.createIndex('diagramName', 'diagramName', { unique: false });
      }
    };
  });
};

async function versionTransaction<T>(mode: IDBTransactionMode, operation: (store: IDBObjectStore) => IDBRequest<T>): Promise<T> {
  const db = await initDB();
  return new Promise((resolve, reject) => {
    const transaction = db.transaction(STORE_NAME, mode);
    let result: T;
    let request: IDBRequest<T>;
    try {
      request = operation(transaction.objectStore(STORE_NAME));
    } catch (error) {
      transaction.abort();
      db.close();
      reject(error);
      return;
    }
    request.onsuccess = () => { result = request.result; };
    transaction.oncomplete = () => { db.close(); resolve(result); };
    transaction.onabort = () => {
      db.close();
      reject(transaction.error ?? request.error ?? new Error('Version storage transaction aborted.'));
    };
  });
}

/**
 * Save a new diagram version
 */
export const saveVersion = async (version: DiagramVersion): Promise<void> => {
  await versionTransaction('readwrite', store => store.put(version));
};

/**
 * Get all versions sorted by timestamp (newest first)
 */
export const getAllVersions = async (): Promise<DiagramVersion[]> => {
  const versions: DiagramVersion[] = await versionTransaction('readonly', store => store.getAll());
  return versions.sort((a, b) => b.timestamp - a.timestamp);
};

/**
 * Get a specific version by ID
 */
export const getVersion = async (versionId: string): Promise<DiagramVersion | null> => {
  return await versionTransaction('readonly', store => store.get(versionId)) ?? null;
};

/**
 * Delete a specific version
 */
export const deleteVersion = async (versionId: string): Promise<void> => {
  await versionTransaction('readwrite', store => store.delete(versionId));
};

/**
 * Delete all versions
 */
export const clearAllVersions = async (): Promise<void> => {
  await versionTransaction('readwrite', store => store.clear());
};

/**
 * Generate a unique version ID
 */
export const generateVersionId = (): string => {
  return `v-${Date.now()}-${Math.random().toString(36).substr(2, 9)}`;
};

/**
 * Create a snapshot of current diagram state
 */
export const createSnapshot = async (
  nodes: any[],
  edges: any[],
  diagramName: string,
  options?: {
    lineageId?: string;
    architecturePrompt?: string;
    originalPrompt?: string;
    validationScore?: number;
    parentVersionId?: string;
    improvementsApplied?: string[];
    notes?: string;
    metadata?: any;
    workflow?: any[];
    titleBlockData?: any;
    pricingScenarios?: PricingScenario[];
    iacBaseline?: IaCBaseline | null;
    settings?: EditorDocument['settings'];
    reviewHistory?: unknown[];
    validationSourceFingerprint?: string | null;
  }
): Promise<DiagramVersion> => {
  const version: DiagramVersion = {
    versionId: generateVersionId(),
    timestamp: Date.now(),
    diagramName,
    nodes: JSON.parse(JSON.stringify(nodes)), // Deep clone
    edges: JSON.parse(JSON.stringify(edges)), // Deep clone
    ...JSON.parse(JSON.stringify(options ?? {})),
  };

  await saveVersion(version);
  return version;
};

/**
 * Get version count
 */
export const getVersionCount = async (): Promise<number> => {
  const versions = await getAllVersions();
  return versions.length;
};
