// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

import { useCallback, useEffect, useRef, useState } from 'react';
import {
  CloudDiagramApiError,
  CloudDiagramDocument,
  CloudDiagramPayload,
  CloudDiagramVersion,
  CloudDocumentContext,
  createCloudDiagram,
  createCloudVersion,
  getCloudDiagram,
  getSharedCloudDiagram,
  updateCloudDiagram,
  updateSharedCloudDiagram,
} from '../services/cloudDiagramService';
import { OperationGeneration } from '../utils/operationGeneration';

const CONTEXT_KEY = 'azurediagarm.cloud-document.v1';
const SHARE_HASH_PREFIX = '#share-';
const SHARE_TOKEN_PATTERN = /^[A-Za-z0-9_-]{43}$/;

export type CloudSyncStatus =
  | 'idle'
  | 'loading'
  | 'saving'
  | 'saved'
  | 'readonly'
  | 'offline'
  | 'unavailable'
  | 'conflict'
  | 'error';

export class CloudDiagramOperationCancelledError extends Error {
  constructor() {
    super('The cloud diagram operation was superseded by a newer document action.');
    this.name = 'CloudDiagramOperationCancelledError';
  }
}

interface PendingSave {
  diagramName: string;
  payload: CloudDiagramPayload;
  serialized: string;
}

interface UseCloudDiagramSyncOptions {
  diagramName: string;
  payload: CloudDiagramPayload;
  enabled: boolean;
  onLoad: (payload: CloudDiagramPayload) => void;
}

function readStoredContext(): CloudDocumentContext | null {
  try {
    const raw = window.sessionStorage.getItem(CONTEXT_KEY);
    if (!raw) return null;
    const parsed = JSON.parse(raw) as Partial<CloudDocumentContext>;
    if (
      typeof parsed.documentId !== 'string'
      || (parsed.access !== 'owner' && parsed.access !== 'shared')
      || (parsed.role !== 'owner' && parsed.role !== 'viewer' && parsed.role !== 'editor')
    ) {
      return null;
    }
    if (parsed.access === 'shared' && typeof parsed.shareToken !== 'string') return null;
    return parsed as CloudDocumentContext;
  } catch {
    return null;
  }
}

function writeStoredContext(context: CloudDocumentContext | null): void {
  try {
    if (context) {
      window.sessionStorage.setItem(CONTEXT_KEY, JSON.stringify(context));
    } else {
      window.sessionStorage.removeItem(CONTEXT_KEY);
    }
  } catch {
    // Session storage is an optimization; cloud persistence still works without it.
  }
}

function normalizeDocument(
  document: CloudDiagramDocument,
  context: CloudDocumentContext,
): CloudDiagramDocument {
  return {
    ...document,
    comments: Array.isArray(document.comments) ? document.comments : [],
    shares: Array.isArray(document.shares) ? document.shares : [],
    access: context.access,
    role: context.role,
  };
}

function contextForDocument(
  document: CloudDiagramDocument,
  fallback?: CloudDocumentContext,
): CloudDocumentContext {
  const access = document.access === 'shared' || fallback?.access === 'shared'
    ? 'shared'
    : 'owner';
  const role = document.role === 'viewer' || document.role === 'editor'
    ? document.role
    : access === 'owner'
      ? 'owner'
      : fallback?.role || 'viewer';
  return {
    documentId: document.id,
    access,
    role,
    shareToken: access === 'shared' ? fallback?.shareToken : undefined,
  };
}

export function useCloudDiagramSync({
  diagramName,
  payload,
  enabled,
  onLoad,
}: UseCloudDiagramSyncOptions) {
  const [document, setDocument] = useState<CloudDiagramDocument | null>(null);
  const [context, setContext] = useState<CloudDocumentContext | null>(null);
  const [status, setStatus] = useState<CloudSyncStatus>('loading');
  const [errorMessage, setErrorMessage] = useState('');
  const [lastSavedAt, setLastSavedAt] = useState<string | null>(null);
  const [initialized, setInitialized] = useState(false);

  const documentRef = useRef<CloudDiagramDocument | null>(null);
  const contextRef = useRef<CloudDocumentContext | null>(null);
  const latestRef = useRef({ diagramName, payload, serialized: JSON.stringify(payload) });
  const lastSavedSerializedRef = useRef('');
  const pendingSaveRef = useRef<PendingSave | null>(null);
  const saveInFlightRef = useRef<Promise<void> | null>(null);
  const debounceTimerRef = useRef<number | null>(null);
  const retryTimerRef = useRef<number | null>(null);
  const lastSaveErrorRef = useRef<Error | null>(null);
  const mountedRef = useRef(true);
  const documentGenerationRef = useRef(new OperationGeneration());
  const serializedPayload = JSON.stringify(payload);

  latestRef.current = {
    diagramName,
    payload,
    serialized: serializedPayload,
  };

  const clearTimers = useCallback(() => {
    if (debounceTimerRef.current !== null) {
      window.clearTimeout(debounceTimerRef.current);
      debounceTimerRef.current = null;
    }
    if (retryTimerRef.current !== null) {
      window.clearTimeout(retryTimerRef.current);
      retryTimerRef.current = null;
    }
  }, []);

  useEffect(() => () => {
    documentGenerationRef.current.advance();
    mountedRef.current = false;
    clearTimers();
  }, [clearTimers]);

  const beginDocumentGeneration = useCallback(
    () => documentGenerationRef.current.advance(),
    [],
  );

  const isCurrentDocumentGeneration = useCallback(
    (generation: number) => (
      mountedRef.current
      && documentGenerationRef.current.isCurrent(generation)
    ),
    [],
  );

  const storeDocument = useCallback((
    nextDocument: CloudDiagramDocument,
    nextContext: CloudDocumentContext,
  ) => {
    const normalized = normalizeDocument(nextDocument, nextContext);
    documentRef.current = normalized;
    contextRef.current = nextContext;
    setDocument(normalized);
    setContext(nextContext);
    writeStoredContext(nextContext);
    return normalized;
  }, []);

  const clearDocument = useCallback((expectedDocumentId?: string) => {
    if (
      expectedDocumentId
      && contextRef.current?.documentId !== expectedDocumentId
    ) {
      return false;
    }
    beginDocumentGeneration();
    clearTimers();
    documentRef.current = null;
    contextRef.current = null;
    pendingSaveRef.current = null;
    lastSavedSerializedRef.current = '';
    lastSaveErrorRef.current = null;
    setDocument(null);
    setContext(null);
    setLastSavedAt(null);
    setErrorMessage('');
    setStatus('idle');
    writeStoredContext(null);
    return true;
  }, [beginDocumentGeneration, clearTimers]);

  const activateDocument = useCallback((
    nextDocument: CloudDiagramDocument,
    nextContext?: CloudDocumentContext,
    applyPayload = true,
    expectedGeneration?: number,
  ) => {
    const generation = expectedGeneration ?? beginDocumentGeneration();
    if (!isCurrentDocumentGeneration(generation)) return null;
    clearTimers();
    pendingSaveRef.current = null;
    const resolvedContext = nextContext || contextForDocument(nextDocument);
    const normalized = storeDocument(nextDocument, resolvedContext);
    lastSavedSerializedRef.current = JSON.stringify(normalized.payload);
    lastSaveErrorRef.current = null;
    setErrorMessage('');
    setLastSavedAt(normalized.updatedAt || new Date().toISOString());
    setStatus(resolvedContext.role === 'viewer' ? 'readonly' : 'saved');
    if (applyPayload) onLoad(normalized.payload);
    return normalized;
  }, [
    beginDocumentGeneration,
    clearTimers,
    isCurrentDocumentGeneration,
    onLoad,
    storeDocument,
  ]);

  const loadContext = useCallback(async (
    nextContext: CloudDocumentContext,
    applyPayload = true,
  ) => {
    const generation = beginDocumentGeneration();
    clearTimers();
    pendingSaveRef.current = null;
    setStatus('loading');
    setErrorMessage('');
    try {
      const nextDocument = nextContext.access === 'shared' && nextContext.shareToken
        ? await getSharedCloudDiagram(nextContext.shareToken)
        : await getCloudDiagram(nextContext.documentId);
      if (!isCurrentDocumentGeneration(generation)) return null;
      return activateDocument(nextDocument, {
        ...nextContext,
        documentId: nextDocument.id,
        role: nextDocument.role || nextContext.role,
      }, applyPayload, generation);
    } catch (error) {
      if (!isCurrentDocumentGeneration(generation)) return null;
      const apiError = error instanceof CloudDiagramApiError ? error : null;
      if (apiError?.status === 404 || apiError?.status === 403) {
        clearDocument();
      } else {
        setStatus(apiError?.status === 503 ? 'unavailable' : 'offline');
        setErrorMessage(error instanceof Error ? error.message : 'Cloud storage is unavailable.');
      }
      return null;
    }
  }, [
    activateDocument,
    beginDocumentGeneration,
    clearDocument,
    clearTimers,
    isCurrentDocumentGeneration,
  ]);

  useEffect(() => {
    let cancelled = false;
    const initialize = async () => {
      const hash = window.location.hash;
      if (hash.startsWith('#version-')) {
        clearDocument();
        if (!cancelled) {
          setInitialized(true);
        }
        return;
      }

      let initialContext = readStoredContext();
      if (hash.startsWith(SHARE_HASH_PREFIX)) {
        let token = '';
        try {
          token = decodeURIComponent(hash.slice(SHARE_HASH_PREFIX.length));
        } catch {
          token = '';
        }
        window.history.replaceState(
          null,
          '',
          window.location.pathname + window.location.search,
        );
        if (!SHARE_TOKEN_PATTERN.test(token)) {
          if (!cancelled) {
            setStatus('error');
            setErrorMessage('The shared diagram link is invalid.');
            setInitialized(true);
          }
          return;
        }
        initialContext = {
          documentId: '',
          access: 'shared',
          role: 'viewer',
          shareToken: token,
        };
      }

      if (initialContext) {
        await loadContext(initialContext, true);
      } else if (!cancelled) {
        setStatus('idle');
      }
      if (!cancelled) setInitialized(true);
    };

    void initialize();
    return () => {
      cancelled = true;
    };
  }, [clearDocument, loadContext]);

  const drainPendingSave = useCallback(async (): Promise<void> => {
    if (saveInFlightRef.current) return saveInFlightRef.current;

    const operation = (async () => {
      while (pendingSaveRef.current && mountedRef.current) {
        const candidate = pendingSaveRef.current;
        pendingSaveRef.current = null;
        if (candidate.serialized === lastSavedSerializedRef.current) continue;

        const generation = documentGenerationRef.current.current();
        const currentContext = contextRef.current;
        const currentDocument = documentRef.current;
        if (currentContext?.role === 'viewer') {
          setStatus('readonly');
          continue;
        }

        setStatus('saving');
        setErrorMessage('');
        lastSaveErrorRef.current = null;
        try {
          let saved: CloudDiagramDocument;
          let savedContext: CloudDocumentContext;
          if (!currentDocument || !currentContext) {
            saved = await createCloudDiagram(candidate.diagramName, candidate.payload);
            savedContext = {
              documentId: saved.id,
              access: 'owner',
              role: 'owner',
            };
          } else if (currentContext.access === 'shared' && currentContext.shareToken) {
            saved = await updateSharedCloudDiagram(
              currentContext.shareToken,
              currentDocument.etag,
              candidate.diagramName,
              candidate.payload,
            );
            savedContext = currentContext;
          } else {
            saved = await updateCloudDiagram(
              currentDocument.id,
              currentDocument.etag,
              candidate.diagramName,
              candidate.payload,
            );
            savedContext = currentContext;
          }

          if (!isCurrentDocumentGeneration(generation)) continue;
          const normalized = storeDocument(saved, savedContext);
          lastSavedSerializedRef.current = candidate.serialized;
          setLastSavedAt(normalized.updatedAt || new Date().toISOString());
          setStatus(savedContext.role === 'viewer' ? 'readonly' : 'saved');
        } catch (error) {
          if (!isCurrentDocumentGeneration(generation)) continue;
          const apiError = error instanceof CloudDiagramApiError ? error : null;
          lastSaveErrorRef.current = error instanceof Error
            ? error
            : new Error('Cloud save failed.');
          setErrorMessage(lastSaveErrorRef.current.message);

          if (apiError?.status === 412 || apiError?.status === 409) {
            setStatus('conflict');
            pendingSaveRef.current = null;
            break;
          }
          if (apiError?.status === 404 && currentContext?.access === 'owner') {
            documentRef.current = null;
            contextRef.current = null;
            setDocument(null);
            setContext(null);
            writeStoredContext(null);
            pendingSaveRef.current = candidate;
            continue;
          }

          setStatus(apiError?.status === 503 ? 'unavailable' : 'offline');
          if (!pendingSaveRef.current) pendingSaveRef.current = candidate;
          if (retryTimerRef.current === null) {
            retryTimerRef.current = window.setTimeout(() => {
              retryTimerRef.current = null;
              void drainPendingSave();
            }, 15_000);
          }
          break;
        }
      }
    })().finally(() => {
      if (saveInFlightRef.current === operation) saveInFlightRef.current = null;
    });

    saveInFlightRef.current = operation;
    return operation;
  }, [isCurrentDocumentGeneration, storeDocument]);

  useEffect(() => {
    if (!initialized || !enabled) return;
    if (contextRef.current?.role === 'viewer') {
      setStatus('readonly');
      return;
    }
    const latest = latestRef.current;
    if (latest.serialized === lastSavedSerializedRef.current) return;

    pendingSaveRef.current = latest;
    if (debounceTimerRef.current !== null) {
      window.clearTimeout(debounceTimerRef.current);
    }
    debounceTimerRef.current = window.setTimeout(() => {
      debounceTimerRef.current = null;
      void drainPendingSave();
    }, 2_000);

    return () => {
      if (debounceTimerRef.current !== null) {
        window.clearTimeout(debounceTimerRef.current);
        debounceTimerRef.current = null;
      }
    };
  }, [
    diagramName,
    drainPendingSave,
    enabled,
    initialized,
    serializedPayload,
  ]);

  const saveNow = useCallback(async (): Promise<CloudDiagramDocument | null> => {
    if (!enabled && !documentRef.current) return null;
    const generation = documentGenerationRef.current.current();
    if (debounceTimerRef.current !== null) {
      window.clearTimeout(debounceTimerRef.current);
      debounceTimerRef.current = null;
    }
    pendingSaveRef.current = latestRef.current;
    await drainPendingSave();
    if (!isCurrentDocumentGeneration(generation)) {
      throw new CloudDiagramOperationCancelledError();
    }
    if (lastSaveErrorRef.current) throw lastSaveErrorRef.current;
    return documentRef.current;
  }, [drainPendingSave, enabled, isCurrentDocumentGeneration]);

  const saveSnapshot = useCallback(async (notes: string): Promise<CloudDiagramVersion | null> => {
    const currentDocument = await saveNow();
    const currentContext = contextRef.current;
    if (!currentDocument || !currentContext || currentContext.role === 'viewer') return null;
    return createCloudVersion(currentContext, notes);
  }, [saveNow]);

  const reloadRemote = useCallback(async () => {
    const currentContext = contextRef.current;
    if (!currentContext) return null;
    return loadContext(currentContext, true);
  }, [loadContext]);

  const saveAsCopy = useCallback(async () => {
    let operationGeneration = beginDocumentGeneration();
    clearTimers();
    pendingSaveRef.current = null;
    const candidate = latestRef.current;
    setStatus('saving');
    setErrorMessage('');
    lastSaveErrorRef.current = null;

    try {
      const saved = await createCloudDiagram(candidate.diagramName, candidate.payload);
      if (!isCurrentDocumentGeneration(operationGeneration)) {
        throw new CloudDiagramOperationCancelledError();
      }
      const savedContext: CloudDocumentContext = {
        documentId: saved.id,
        access: 'owner',
        role: 'owner',
      };
      const normalized = activateDocument(saved, savedContext, false);
      if (!normalized) throw new CloudDiagramOperationCancelledError();
      lastSavedSerializedRef.current = candidate.serialized;
      setLastSavedAt(normalized.updatedAt || new Date().toISOString());
      setStatus('saved');

      const activationGeneration = documentGenerationRef.current.current();
      operationGeneration = activationGeneration;
      const latestAfterCopy = latestRef.current;
      if (latestAfterCopy.serialized !== candidate.serialized) {
        pendingSaveRef.current = latestAfterCopy;
        await drainPendingSave();
        if (!isCurrentDocumentGeneration(activationGeneration)) {
          throw new CloudDiagramOperationCancelledError();
        }
        if (lastSaveErrorRef.current) throw lastSaveErrorRef.current;
      }
      return documentRef.current;
    } catch (error) {
      if (error instanceof CloudDiagramOperationCancelledError) throw error;
      if (!isCurrentDocumentGeneration(operationGeneration)) {
        throw new CloudDiagramOperationCancelledError();
      }
      const apiError = error instanceof CloudDiagramApiError ? error : null;
      lastSaveErrorRef.current = error instanceof Error
        ? error
        : new Error('Cloud copy failed.');
      setErrorMessage(lastSaveErrorRef.current.message);
      setStatus(apiError?.status === 503 ? 'unavailable' : 'offline');
      throw lastSaveErrorRef.current;
    }
  }, [
    activateDocument,
    beginDocumentGeneration,
    clearTimers,
    drainPendingSave,
    isCurrentDocumentGeneration,
  ]);

  const openDocument = useCallback((
    nextDocument: CloudDiagramDocument,
    nextContext?: CloudDocumentContext,
  ) => {
    activateDocument(nextDocument, nextContext || {
      documentId: nextDocument.id,
      access: 'owner',
      role: 'owner',
    }, true);
  }, [activateDocument]);

  const restoreVersion = useCallback((
    version: CloudDiagramVersion,
    baseDocument: CloudDiagramDocument,
    baseContext: CloudDocumentContext,
  ) => {
    beginDocumentGeneration();
    clearTimers();
    const normalized = storeDocument(baseDocument, baseContext);
    lastSavedSerializedRef.current = JSON.stringify(normalized.payload);
    pendingSaveRef.current = null;
    lastSaveErrorRef.current = null;
    setErrorMessage('');
    setStatus(baseContext.role === 'viewer' ? 'readonly' : 'saved');
    onLoad(version.payload);
  }, [beginDocumentGeneration, clearTimers, onLoad, storeDocument]);

  const replaceCurrentDocument = useCallback((nextDocument: CloudDiagramDocument) => {
    const currentContext = contextRef.current;
    if (!currentContext || nextDocument.id !== currentContext.documentId) return;
    const normalized = storeDocument(nextDocument, currentContext);
    lastSavedSerializedRef.current = JSON.stringify(normalized.payload);
    setLastSavedAt(normalized.updatedAt || new Date().toISOString());
    setStatus(currentContext.role === 'viewer' ? 'readonly' : 'saved');
  }, [storeDocument]);

  return {
    document,
    context,
    status,
    errorMessage,
    lastSavedAt,
    saveNow,
    saveSnapshot,
    reset: clearDocument,
    reloadRemote,
    saveAsCopy,
    openDocument,
    restoreVersion,
    replaceCurrentDocument,
  };
}
