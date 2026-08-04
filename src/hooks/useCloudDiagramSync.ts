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
  getCloudDiagramRetryDelay,
  getSharedCloudDiagram,
  updateCloudDiagram,
  updateSharedCloudDiagram,
} from '../services/cloudDiagramService';
import { OperationGeneration } from '../utils/operationGeneration';

const CONTEXT_KEY = 'azurediagarm.cloud-document.v1';
const SHARE_HASH_PREFIX = '#share-';
const SHARE_TOKEN_PATTERN = /^[A-Za-z0-9_-]{43}$/;
const MAX_DIAGRAM_NAME_LENGTH = 200;

function normalizeDiagramName(value: string): string {
  return (value.trim() || 'Untitled Architecture').slice(0, MAX_DIAGRAM_NAME_LENGTH);
}

function isNonRetryableClientError(error: CloudDiagramApiError | null): boolean {
  return Boolean(
    error
    && error.status >= 400
    && error.status < 500
    && error.status !== 408
    && error.status !== 409
    && error.status !== 412
    && error.status !== 429,
  );
}

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
  force?: boolean;
}

interface PendingCreate {
  idempotencyKey: string;
  candidate: PendingSave;
}

function createIdempotencyKey(): string {
  if (typeof crypto !== 'undefined' && typeof crypto.randomUUID === 'function') {
    return crypto.randomUUID();
  }
  const random = Math.random().toString(36).slice(2);
  return `create-${Date.now().toString(36)}-${random}`;
}

function sameSaveCandidate(left: PendingSave, right: PendingSave): boolean {
  return (
    left.diagramName === right.diagramName
    && left.serialized === right.serialized
  );
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
  const latestRef = useRef({
    diagramName: normalizeDiagramName(diagramName),
    payload,
    serialized: JSON.stringify(payload),
  });
  const lastSavedDiagramNameRef = useRef('');
  const lastSavedSerializedRef = useRef('');
  const pendingSaveRef = useRef<PendingSave | null>(null);
  const pendingCreateRef = useRef<PendingCreate | null>(null);
  const saveInFlightRef = useRef<Promise<void> | null>(null);
  const restoreWriteInFlightRef = useRef<Promise<void> | null>(null);
  const debounceTimerRef = useRef<number | null>(null);
  const retryTimerRef = useRef<number | null>(null);
  const loadAbortControllerRef = useRef<AbortController | null>(null);
  const lastSaveErrorRef = useRef<Error | null>(null);
  const conflictRef = useRef(false);
  const conflictEpochRef = useRef(0);
  const enabledRef = useRef(enabled);
  const autosaveBlockedRef = useRef(false);
  const viewerBaselineDocumentIdRef = useRef<string | null>(null);
  const invalidShareLinkRef = useRef(false);
  const autosaveSuspendedGenerationRef = useRef<number | null>(null);
  const mountedRef = useRef(true);
  const documentGenerationRef = useRef(new OperationGeneration());
  const serializedPayload = JSON.stringify(payload);

  latestRef.current = {
    diagramName: normalizeDiagramName(diagramName),
    payload,
    serialized: serializedPayload,
  };
  enabledRef.current = enabled;

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

  const abortActiveLoad = useCallback(() => {
    loadAbortControllerRef.current?.abort();
    loadAbortControllerRef.current = null;
  }, []);

  const enterConflict = useCallback((error?: unknown) => {
    const conflictError = error instanceof Error
      ? error
      : new Error(
          'The cloud diagram changed in another session. Reload it or save your work as a copy.',
        );
    clearTimers();
    pendingSaveRef.current = null;
    lastSaveErrorRef.current = conflictError;
    conflictRef.current = true;
    conflictEpochRef.current += 1;
    setErrorMessage(conflictError.message);
    setStatus('conflict');
  }, [clearTimers]);

  const reportConflict = useCallback((
    documentId: string,
    error?: unknown,
    expectedRevision?: number,
    expectedEtag?: string,
  ) => {
    if (contextRef.current?.documentId !== documentId) return false;
    const currentDocument = documentRef.current;
    if (
      currentDocument
      && expectedRevision !== undefined
      && (
        currentDocument.revision > expectedRevision
        || (
          currentDocument.revision === expectedRevision
          && Boolean(expectedEtag)
          && currentDocument.etag !== expectedEtag
        )
      )
    ) return false;
    enterConflict(error);
    return true;
  }, [enterConflict]);

  useEffect(() => {
    // React StrictMode intentionally mounts, cleans up, and remounts effects in
    // development. Restore the flag on every effect setup so the simulated
    // cleanup cannot permanently invalidate all cloud operations in CI/dev.
    const documentGeneration = documentGenerationRef.current;
    mountedRef.current = true;
    return () => {
      documentGeneration.advance();
      mountedRef.current = false;
      abortActiveLoad();
      clearTimers();
    };
  }, [abortActiveLoad, clearTimers]);

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
    autosaveBlockedRef.current = false;
    invalidShareLinkRef.current = false;
    documentRef.current = normalized;
    contextRef.current = nextContext;
    pendingCreateRef.current = null;
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
    abortActiveLoad();
    clearTimers();
    documentRef.current = null;
    contextRef.current = null;
    pendingSaveRef.current = null;
    pendingCreateRef.current = null;
    lastSavedDiagramNameRef.current = '';
    lastSavedSerializedRef.current = '';
    lastSaveErrorRef.current = null;
    conflictRef.current = false;
    autosaveBlockedRef.current = false;
    viewerBaselineDocumentIdRef.current = null;
    invalidShareLinkRef.current = false;
    setDocument(null);
    setContext(null);
    setLastSavedAt(null);
    setErrorMessage('');
    setStatus('idle');
    writeStoredContext(null);
    return true;
  }, [abortActiveLoad, beginDocumentGeneration, clearTimers]);

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
    autosaveBlockedRef.current = false;
    viewerBaselineDocumentIdRef.current = applyPayload && resolvedContext.role === 'viewer'
      ? normalized.id
      : null;
    lastSavedDiagramNameRef.current = normalized.diagramName;
    lastSavedSerializedRef.current = JSON.stringify(normalized.payload);
    lastSaveErrorRef.current = null;
    conflictRef.current = false;
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
    const hadDocument = Boolean(documentRef.current);
    const generation = beginDocumentGeneration();
    const conflictEpoch = conflictEpochRef.current;
    abortActiveLoad();
    const loadController = new AbortController();
    loadAbortControllerRef.current = loadController;
    clearTimers();
    pendingSaveRef.current = null;
    setStatus('loading');
    setErrorMessage('');
    try {
      const nextDocument = nextContext.access === 'shared' && nextContext.shareToken
        ? await getSharedCloudDiagram(nextContext.shareToken, { signal: loadController.signal })
        : await getCloudDiagram(nextContext.documentId, { signal: loadController.signal });
      if (
        !isCurrentDocumentGeneration(generation)
        || conflictEpoch !== conflictEpochRef.current
      ) return null;
      const resolvedAccess = nextDocument.access === 'owner'
        ? 'owner'
        : nextContext.access;
      return activateDocument(nextDocument, {
        documentId: nextDocument.id,
        access: resolvedAccess,
        role: nextDocument.role || nextContext.role,
        shareToken: resolvedAccess === 'shared' ? nextContext.shareToken : undefined,
      }, applyPayload, generation);
    } catch (error) {
      if (
        !isCurrentDocumentGeneration(generation)
        || conflictEpoch !== conflictEpochRef.current
      ) return null;
      const apiError = error instanceof CloudDiagramApiError ? error : null;
      if (conflictRef.current) {
        setStatus('conflict');
        setErrorMessage(error instanceof Error ? error.message : 'Cloud storage is unavailable.');
      } else if (apiError?.status === 404 || apiError?.status === 403) {
        if (nextContext.access === 'shared') {
          autosaveBlockedRef.current = true;
          setStatus('error');
          setErrorMessage(error instanceof Error ? error.message : 'The shared diagram is unavailable.');
        } else {
          clearDocument();
        }
      } else {
        if (!hadDocument) autosaveBlockedRef.current = true;
        setStatus(apiError?.status === 503 ? 'unavailable' : 'offline');
        setErrorMessage(error instanceof Error ? error.message : 'Cloud storage is unavailable.');
      }
      return null;
    } finally {
      if (loadAbortControllerRef.current === loadController) {
        loadAbortControllerRef.current = null;
      }
    }
  }, [
    activateDocument,
    abortActiveLoad,
    beginDocumentGeneration,
    clearDocument,
    clearTimers,
    isCurrentDocumentGeneration,
  ]);

  useEffect(() => {
    let cancelled = false;
    const initialize = async () => {
      if (invalidShareLinkRef.current) {
        if (!cancelled) {
          setStatus('error');
          setErrorMessage('The shared diagram link is invalid.');
          setInitialized(true);
        }
        return;
      }
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
            invalidShareLinkRef.current = true;
            autosaveBlockedRef.current = true;
            writeStoredContext(null);
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
        writeStoredContext(initialContext);
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
    while (restoreWriteInFlightRef.current || saveInFlightRef.current) {
      const activeWrite = restoreWriteInFlightRef.current || saveInFlightRef.current;
      if (activeWrite) await activeWrite;
    }
    if (!pendingSaveRef.current || !mountedRef.current) return;

    const operation = (async () => {
      while (pendingSaveRef.current && mountedRef.current) {
        const candidate = pendingSaveRef.current;
        pendingSaveRef.current = null;
        if (!documentRef.current && !enabledRef.current) {
          lastSaveErrorRef.current = null;
          setErrorMessage('');
          setStatus('idle');
          break;
        }
        if (
          !candidate.force
          && candidate.diagramName === lastSavedDiagramNameRef.current
          && candidate.serialized === lastSavedSerializedRef.current
        ) continue;

        const generation = documentGenerationRef.current.current();
        const conflictEpoch = conflictEpochRef.current;
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
          let savedCandidate = candidate;
          if (!currentDocument || !currentContext) {
            let createAttempt = pendingCreateRef.current;
            if (!createAttempt) {
              createAttempt = {
                idempotencyKey: createIdempotencyKey(),
                candidate,
              };
              pendingCreateRef.current = createAttempt;
            } else if (!sameSaveCandidate(createAttempt.candidate, candidate)) {
              pendingSaveRef.current = candidate;
            }
            savedCandidate = createAttempt.candidate;
            saved = await createCloudDiagram(
              savedCandidate.diagramName,
              savedCandidate.payload,
              createAttempt.idempotencyKey,
            );
            if (pendingCreateRef.current === createAttempt) {
              pendingCreateRef.current = null;
            }
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

          if (
            !isCurrentDocumentGeneration(generation)
            || conflictEpoch !== conflictEpochRef.current
          ) continue;
          const normalized = storeDocument(saved, savedContext);
          lastSavedDiagramNameRef.current = savedCandidate.diagramName;
          lastSavedSerializedRef.current = savedCandidate.serialized;
          const latest = latestRef.current;
          if (
            !sameSaveCandidate(savedCandidate, latest)
            && !pendingSaveRef.current
          ) {
            pendingSaveRef.current = latest;
          }
          conflictRef.current = false;
          setLastSavedAt(normalized.updatedAt || new Date().toISOString());
          setStatus(savedContext.role === 'viewer' ? 'readonly' : 'saved');
        } catch (error) {
          if (
            !isCurrentDocumentGeneration(generation)
            || conflictEpoch !== conflictEpochRef.current
          ) continue;
          const apiError = error instanceof CloudDiagramApiError ? error : null;
          lastSaveErrorRef.current = error instanceof Error
            ? error
            : new Error('Cloud save failed.');
          setErrorMessage(lastSaveErrorRef.current.message);

          if (apiError?.status === 412 || apiError?.status === 409) {
            enterConflict(lastSaveErrorRef.current);
            break;
          }
          if (apiError?.status === 404 && currentContext) {
            const latestCandidate = latestRef.current;
            pendingSaveRef.current = null;
            if (
              currentContext.access === 'shared'
              || (latestCandidate.payload.nodes.length === 0 && enabledRef.current)
            ) {
              enterConflict(new Error(
                'The cloud diagram was deleted in another session. Save your work as a copy or discard it.',
              ));
              break;
            }
            documentRef.current = null;
            contextRef.current = null;
            setDocument(null);
            setContext(null);
            writeStoredContext(null);
            if (latestCandidate.payload.nodes.length === 0) {
              lastSavedDiagramNameRef.current = latestCandidate.diagramName;
              lastSavedSerializedRef.current = latestCandidate.serialized;
              lastSaveErrorRef.current = null;
              conflictRef.current = false;
              setLastSavedAt(null);
              setErrorMessage('');
              setStatus('idle');
              break;
            }
            lastSavedDiagramNameRef.current = '';
            lastSavedSerializedRef.current = '';
            pendingSaveRef.current = latestCandidate;
            continue;
          }
          if (isNonRetryableClientError(apiError)) {
            pendingCreateRef.current = null;
            pendingSaveRef.current = null;
            setStatus('error');
            break;
          }

          setStatus(apiError?.status === 503 ? 'unavailable' : 'offline');
          if (!pendingSaveRef.current) pendingSaveRef.current = candidate;
          if (retryTimerRef.current === null) {
            const retryDelayMs = getCloudDiagramRetryDelay(apiError, 15_000);
            retryTimerRef.current = window.setTimeout(() => {
              retryTimerRef.current = null;
              void drainPendingSave();
            }, retryDelayMs);
          }
          break;
        }
      }
    })().finally(() => {
      if (saveInFlightRef.current === operation) saveInFlightRef.current = null;
    });

    saveInFlightRef.current = operation;
    return operation;
  }, [enterConflict, isCurrentDocumentGeneration, storeDocument]);

  useEffect(() => {
    if (
      !initialized
      || conflictRef.current
      || autosaveBlockedRef.current
      || autosaveSuspendedGenerationRef.current !== null
    ) return;
    if (!enabled && !documentRef.current && !saveInFlightRef.current) {
      pendingSaveRef.current = null;
      if (retryTimerRef.current !== null) {
        window.clearTimeout(retryTimerRef.current);
        retryTimerRef.current = null;
      }
      lastSaveErrorRef.current = null;
      setErrorMessage('');
      setStatus('idle');
      return;
    }
    if (contextRef.current?.role === 'viewer') {
      if (viewerBaselineDocumentIdRef.current === contextRef.current.documentId) {
        lastSavedDiagramNameRef.current = latestRef.current.diagramName;
        lastSavedSerializedRef.current = latestRef.current.serialized;
        viewerBaselineDocumentIdRef.current = null;
      }
      setStatus('readonly');
      return;
    }
    const latest = latestRef.current;
    if (
      latest.diagramName === lastSavedDiagramNameRef.current
      && latest.serialized === lastSavedSerializedRef.current
    ) {
      if (saveInFlightRef.current) {
        pendingSaveRef.current = { ...latest, force: true };
      } else {
        pendingSaveRef.current = null;
        if (retryTimerRef.current !== null) {
          window.clearTimeout(retryTimerRef.current);
          retryTimerRef.current = null;
        }
        lastSaveErrorRef.current = null;
        setErrorMessage('');
        setStatus(documentRef.current ? 'saved' : 'idle');
      }
      return;
    }

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

  useEffect(() => {
    const hasUnsavedWork = () => {
      const latest = latestRef.current;
      return Boolean(
        conflictRef.current
        || pendingSaveRef.current
        || saveInFlightRef.current
        || (
          (documentRef.current || enabledRef.current)
          && (
            latest.diagramName !== lastSavedDiagramNameRef.current
            || latest.serialized !== lastSavedSerializedRef.current
          )
        )
      );
    };
    const handleBeforeUnload = (event: BeforeUnloadEvent) => {
      if (!hasUnsavedWork()) return;
      event.preventDefault();
      event.returnValue = '';
    };
    const handleVisibilityChange = () => {
      if (
        window.document.visibilityState !== 'hidden'
        || contextRef.current?.role === 'viewer'
        || conflictRef.current
        || autosaveBlockedRef.current
        || autosaveSuspendedGenerationRef.current !== null
        || (!documentRef.current && !enabledRef.current)
      ) return;
      const latest = latestRef.current;
      if (
        !pendingSaveRef.current
        && latest.diagramName === lastSavedDiagramNameRef.current
        && latest.serialized === lastSavedSerializedRef.current
      ) return;
      if (debounceTimerRef.current !== null) {
        window.clearTimeout(debounceTimerRef.current);
        debounceTimerRef.current = null;
      }
      pendingSaveRef.current = {
        ...latest,
        force: true,
      };
      void drainPendingSave();
    };

    window.addEventListener('beforeunload', handleBeforeUnload);
    window.document.addEventListener('visibilitychange', handleVisibilityChange);
    return () => {
      window.removeEventListener('beforeunload', handleBeforeUnload);
      window.document.removeEventListener('visibilitychange', handleVisibilityChange);
    };
  }, [drainPendingSave]);

  const saveNow = useCallback(async (
    options?: { force?: boolean },
  ): Promise<CloudDiagramDocument | null> => {
    if (!enabled && !documentRef.current && !saveInFlightRef.current) return null;
    if (conflictRef.current) {
      throw lastSaveErrorRef.current || new Error(
        'The cloud diagram changed in another session. Reload it or save your work as a copy.',
      );
    }
    const generation = documentGenerationRef.current.current();
    if (debounceTimerRef.current !== null) {
      window.clearTimeout(debounceTimerRef.current);
      debounceTimerRef.current = null;
    }
    pendingSaveRef.current = {
      ...latestRef.current,
      force: options?.force === true,
    };
    await drainPendingSave();
    if (!isCurrentDocumentGeneration(generation)) {
      throw new CloudDiagramOperationCancelledError();
    }
    if (lastSaveErrorRef.current) throw lastSaveErrorRef.current;
    return documentRef.current;
  }, [drainPendingSave, enabled, isCurrentDocumentGeneration]);

  const saveSnapshot = useCallback(async (notes: string): Promise<CloudDiagramVersion | null> => {
    const currentDocument = await saveNow({ force: true });
    const currentContext = contextRef.current;
    if (!currentDocument || !currentContext || currentContext.role === 'viewer') return null;
    return createCloudVersion(currentContext, notes);
  }, [saveNow]);

  const discardPendingSave = useCallback(() => {
    beginDocumentGeneration();
    clearTimers();
    pendingSaveRef.current = null;
  }, [beginDocumentGeneration, clearTimers]);

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
      lastSavedDiagramNameRef.current = candidate.diagramName;
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
      setStatus(conflictRef.current
        ? 'conflict'
        : apiError?.status === 503
          ? 'unavailable'
          : 'offline');
      throw lastSaveErrorRef.current;
    }
  }, [
    activateDocument,
    beginDocumentGeneration,
    clearTimers,
    drainPendingSave,
    isCurrentDocumentGeneration,
  ]);

  const saveAsDetachedCopy = useCallback(async () => {
    const generation = documentGenerationRef.current.current();
    const conflictEpoch = conflictEpochRef.current;
    let candidate = latestRef.current;
    setStatus('saving');
    setErrorMessage('');
    lastSaveErrorRef.current = null;

    try {
      let saved = await createCloudDiagram(candidate.diagramName, candidate.payload);
      while (
        isCurrentDocumentGeneration(generation)
        && conflictEpoch === conflictEpochRef.current
      ) {
        const latest = latestRef.current;
        if (latest.serialized === candidate.serialized) {
          setStatus(contextRef.current?.role === 'viewer' ? 'readonly' : 'saved');
          return normalizeDocument(saved, {
            documentId: saved.id,
            access: 'owner',
            role: 'owner',
          });
        }
        candidate = latest;
        saved = await updateCloudDiagram(
          saved.id,
          saved.etag,
          candidate.diagramName,
          candidate.payload,
        );
      }
      throw new CloudDiagramOperationCancelledError();
    } catch (error) {
      if (error instanceof CloudDiagramOperationCancelledError) throw error;
      if (
        !isCurrentDocumentGeneration(generation)
        || conflictEpoch !== conflictEpochRef.current
      ) {
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
  }, [isCurrentDocumentGeneration]);

  const openDocument = useCallback((
    nextDocument: CloudDiagramDocument,
    nextContext?: CloudDocumentContext,
  ) => {
    abortActiveLoad();
    activateDocument(nextDocument, nextContext || {
      documentId: nextDocument.id,
      access: 'owner',
      role: 'owner',
    }, true);
  }, [abortActiveLoad, activateDocument]);

  const resumeSuspendedAutosave = useCallback(async (
    generation: number,
    options: { mode: 'debounce' | 'retry' | 'none'; force?: boolean },
  ) => {
    if (autosaveSuspendedGenerationRef.current !== generation) return;
    autosaveSuspendedGenerationRef.current = null;
    const generationIsCurrent = isCurrentDocumentGeneration(generation);
    if (
      conflictRef.current
      || autosaveBlockedRef.current
      || contextRef.current?.role === 'viewer'
      || (!enabledRef.current && !documentRef.current)
    ) return;

    if (generationIsCurrent && options.mode === 'none') return;
    const latest = latestRef.current;
    const force = generationIsCurrent && options.force === true;
    if (
      !force
      && latest.diagramName === lastSavedDiagramNameRef.current
      && latest.serialized === lastSavedSerializedRef.current
    ) return;
    pendingSaveRef.current = {
      ...latest,
      force,
    };
    if (!generationIsCurrent || options.mode === 'debounce') {
      if (debounceTimerRef.current !== null) {
        window.clearTimeout(debounceTimerRef.current);
      }
      debounceTimerRef.current = window.setTimeout(() => {
        debounceTimerRef.current = null;
        void drainPendingSave();
      }, 2_000);
      return;
    }
    if (retryTimerRef.current === null) {
      retryTimerRef.current = window.setTimeout(() => {
        retryTimerRef.current = null;
        void drainPendingSave();
      }, 15_000);
    }
  }, [drainPendingSave, isCurrentDocumentGeneration]);

  const restoreVersion = useCallback(async (
    version: CloudDiagramVersion,
    baseDocument: CloudDiagramDocument,
    baseContext: CloudDocumentContext,
  ): Promise<boolean> => {
    const startingGeneration = documentGenerationRef.current.current();
    while (restoreWriteInFlightRef.current || saveInFlightRef.current) {
      const activeWrite = restoreWriteInFlightRef.current || saveInFlightRef.current;
      if (activeWrite) await activeWrite;
    }
    if (!isCurrentDocumentGeneration(startingGeneration)) {
      throw new CloudDiagramOperationCancelledError();
    }
    if (conflictRef.current) return false;
    const generation = beginDocumentGeneration();
    const conflictEpoch = conflictEpochRef.current;
    autosaveSuspendedGenerationRef.current = generation;
    clearTimers();
    const normalized = storeDocument(baseDocument, baseContext);
    lastSavedDiagramNameRef.current = normalized.diagramName;
    lastSavedSerializedRef.current = JSON.stringify(normalized.payload);
    pendingSaveRef.current = null;
    lastSaveErrorRef.current = null;
    conflictRef.current = false;
    setErrorMessage('');
    setStatus(baseContext.role === 'viewer' ? 'readonly' : 'saving');
    const restoredDiagramName = normalizeDiagramName(version.diagramName);
    latestRef.current = {
      diagramName: restoredDiagramName,
      payload: version.payload,
      serialized: JSON.stringify(version.payload),
    };
    let restoreWriteBlocker: Promise<void> | null = null;
    let releaseRestoreWrite: (() => void) | null = null;
    const finishRestoreWrite = () => {
      if (!restoreWriteBlocker) return;
      if (restoreWriteInFlightRef.current === restoreWriteBlocker) {
        restoreWriteInFlightRef.current = null;
      }
      releaseRestoreWrite?.();
      restoreWriteBlocker = null;
      releaseRestoreWrite = null;
    };
    try {
      onLoad(version.payload);
      if (baseContext.role === 'viewer') {
        await resumeSuspendedAutosave(generation, { mode: 'debounce' });
        if (!isCurrentDocumentGeneration(generation)) {
          throw new CloudDiagramOperationCancelledError();
        }
        return true;
      }

      restoreWriteBlocker = new Promise<void>((resolve) => {
        releaseRestoreWrite = resolve;
      });
      restoreWriteInFlightRef.current = restoreWriteBlocker;
      const verified = baseContext.access === 'shared' && baseContext.shareToken
        ? await updateSharedCloudDiagram(
            baseContext.shareToken,
            baseDocument.etag,
            restoredDiagramName,
            version.payload,
          )
        : await updateCloudDiagram(
            baseDocument.id,
            baseDocument.etag,
            restoredDiagramName,
            version.payload,
          );
      if (
        !isCurrentDocumentGeneration(generation)
        || conflictEpoch !== conflictEpochRef.current
      ) {
        throw new CloudDiagramOperationCancelledError();
      }
      const verifiedDocument = storeDocument(verified, baseContext);
      lastSavedDiagramNameRef.current = verifiedDocument.diagramName;
      lastSavedSerializedRef.current = JSON.stringify(verifiedDocument.payload);
      lastSaveErrorRef.current = null;
      setLastSavedAt(verifiedDocument.updatedAt || new Date().toISOString());
      setStatus('saved');
      finishRestoreWrite();
    } catch (error) {
      finishRestoreWrite();
      if (
        error instanceof CloudDiagramOperationCancelledError
        || !isCurrentDocumentGeneration(generation)
        || conflictEpoch !== conflictEpochRef.current
      ) {
        await resumeSuspendedAutosave(generation, { mode: 'retry' });
        throw new CloudDiagramOperationCancelledError();
      }
      const apiError = error instanceof CloudDiagramApiError ? error : null;
      const terminalClientError = isNonRetryableClientError(apiError);
      lastSaveErrorRef.current = error instanceof Error
        ? error
        : new Error('Cloud snapshot restore failed.');
      setErrorMessage(lastSaveErrorRef.current.message);
      if (apiError?.status === 412 || apiError?.status === 409) {
        enterConflict(lastSaveErrorRef.current);
      } else {
        setStatus(terminalClientError
          ? 'error'
          : apiError?.status === 503
            ? 'unavailable'
            : 'offline');
      }
      await resumeSuspendedAutosave(generation, {
        mode: terminalClientError ? 'none' : 'retry',
        force: !terminalClientError && !conflictRef.current,
      });
      throw lastSaveErrorRef.current;
    }
    await resumeSuspendedAutosave(generation, { mode: 'debounce' });
    if (
      !isCurrentDocumentGeneration(generation)
      || conflictEpoch !== conflictEpochRef.current
    ) {
      throw new CloudDiagramOperationCancelledError();
    }
    return true;
  }, [
    beginDocumentGeneration,
    clearTimers,
    enterConflict,
    isCurrentDocumentGeneration,
    onLoad,
    resumeSuspendedAutosave,
    storeDocument,
  ]);

  const replaceCurrentDocument = useCallback((nextDocument: CloudDiagramDocument) => {
    const currentContext = contextRef.current;
    if (!currentContext || nextDocument.id !== currentContext.documentId) return;
    if (conflictRef.current) return;
    const normalized = normalizeDocument(nextDocument, currentContext);
    const remoteSerialized = JSON.stringify(normalized.payload);
    const currentDocument = documentRef.current;
    if (
      currentDocument
      && (
        normalized.revision < currentDocument.revision
        || (
          normalized.revision === currentDocument.revision
          && Boolean(currentDocument.etag)
          && normalized.etag !== currentDocument.etag
        )
      )
    ) return;
    const knownRemoteSerialized = currentDocument
      ? JSON.stringify(currentDocument.payload)
      : '';
    if (
      remoteSerialized !== latestRef.current.serialized
      && remoteSerialized !== knownRemoteSerialized
    ) {
      if (currentDocument) {
        const metadataOnlyDocument = {
          ...currentDocument,
          comments: normalized.comments,
          shares: normalized.shares,
        };
        documentRef.current = metadataOnlyDocument;
        setDocument(metadataOnlyDocument);
      }
      const conflictError = new Error(
        'The cloud diagram changed in another session. Reload it or save your work as a copy.',
      );
      enterConflict(conflictError);
      return;
    }
    storeDocument(normalized, currentContext);
    if (
      normalized.diagramName === latestRef.current.diagramName
      && remoteSerialized === latestRef.current.serialized
    ) {
      lastSavedDiagramNameRef.current = normalized.diagramName;
      lastSavedSerializedRef.current = remoteSerialized;
    }
    lastSaveErrorRef.current = null;
    setErrorMessage('');
    setLastSavedAt(normalized.updatedAt || new Date().toISOString());
    setStatus(currentContext.role === 'viewer' ? 'readonly' : 'saved');
  }, [enterConflict, storeDocument]);

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
    saveAsDetachedCopy,
    openDocument,
    restoreVersion,
    replaceCurrentDocument,
    reportConflict,
    discardPendingSave,
  };
}
