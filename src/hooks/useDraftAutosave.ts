import { useCallback, useEffect, useRef, useState } from 'react';
import { cloneEditorDocument, type EditorDocument } from '../services/editorHistory';
import { discardDraft, draftFingerprint, readDraft, writeDraft, DraftConflictError, type DiagramDraft } from '../services/draftStorage';

export type DraftSaveStatus = 'loading' | 'idle' | 'saving' | 'saved' | 'error';
interface DraftSession {
  key: string;
  revision: number | null;
  fingerprint: string;
  loaded: boolean;
  writing: boolean;
  conflicted: boolean;
  pending: Promise<DiagramDraft | null> | null;
  saved: DiagramDraft | null;
}

export function useDraftAutosave(
  value: EditorDocument,
  onRestore: (document: EditorDocument) => void,
  options: { key: string; enabled: boolean; skipRecovery?: boolean; paused?: boolean },
) {
  const latest = useRef(value);
  latest.current = value;
  const baseline = useRef<{ key: string | null; fingerprint: string }>();
  baseline.current ??= { key: null, fingerprint: draftFingerprint(value) };
  const restore = useRef(onRestore);
  restore.current = onRestore;
  const session = useRef<DraftSession | null>(null);
  const saveTimer = useRef<ReturnType<typeof setTimeout>>();
  const [status, setStatus] = useState<DraftSaveStatus>('loading');
  const [error, setError] = useState<Error | null>(null);
  const [recovery, setRecovery] = useState<DiagramDraft | null>(null);
  const [ready, setReady] = useState(false);
  const [savedAt, setSavedAt] = useState<number>();
  const [savedDraft, setSavedDraft] = useState<DiagramDraft | null>(null);
  const [saveEpoch, setSaveEpoch] = useState(0);
  const [retryCount, setRetryCount] = useState(0);

  useEffect(() => {
    if (!options.enabled) return;
    const initial = baseline.current!;
    // Identity lookup can finish after a cloud or URL restore. Those changes
    // still need saving; they are not a new, already-persisted baseline.
    if (initial.key !== null && initial.key !== options.key) {
      initial.fingerprint = draftFingerprint(latest.current);
    }
    initial.key = options.key;
    const current: DraftSession = {
      key: options.key, revision: null, fingerprint: initial.fingerprint, loaded: false, writing: false, conflicted: false,
      pending: null, saved: null,
    };
    session.current = current;
    setReady(false);
    setStatus('loading');
    setError(null);
    setRecovery(null);
    setSavedDraft(null);
    setSavedAt(undefined);
    readDraft(current.key).then(draft => {
      if (session.current !== current) return;
      current.revision = draft?.revision ?? null;
      current.saved = draft;
      current.loaded = true;
      if (draft && !options.skipRecovery) {
        setRecovery(draft);
        setStatus('idle');
      } else {
        setRecovery(null);
        setReady(true);
        setStatus('idle');
      }
    }).catch(cause => {
      if (session.current !== current) return;
      setError(cause instanceof Error ? cause : new Error(String(cause)));
      setStatus('error');
      // Editing and JSON downloads remain available, but no unknown draft is overwritten.
      setReady(true);
    });
    return () => {
      if (session.current === current) session.current = null;
      clearTimeout(saveTimer.current);
    };
  }, [options.enabled, options.key, options.skipRecovery, retryCount]);

  const saveLatest = useCallback(async (): Promise<DiagramDraft | null> => {
    const current = session.current;
    if (!current?.loaded) throw new Error('Local draft storage is not ready. Download your diagram before switching work.');
    while (current.pending) await current.pending;
    if (session.current !== current) return null;
    if (current.conflicted) throw new DraftConflictError();
    const snapshot = cloneEditorDocument(latest.current);
    const fingerprint = draftFingerprint(snapshot);
    if (fingerprint === current.fingerprint) return current.saved;
    current.writing = true;
    setStatus('saving');
    setError(null);
    const operation = (async () => {
      try {
        const saved = await writeDraft(current.key, snapshot, current.revision);
        current.revision = saved.revision;
        current.fingerprint = fingerprint;
        current.saved = saved;
        if (session.current !== current) return saved;
        setSavedAt(saved.updatedAt);
        setSavedDraft(saved);
        setSaveEpoch(epoch => epoch + 1);
        setStatus(draftFingerprint(latest.current) === fingerprint ? 'saved' : 'saving');
        return saved;
      } catch (cause) {
        if (session.current === current) {
          if (cause instanceof DraftConflictError) current.conflicted = true;
          setError(cause instanceof Error ? cause : new Error(String(cause)));
          setStatus('error');
        }
        throw cause;
      } finally {
        current.writing = false;
        current.pending = null;
      }
    })();
    current.pending = operation;
    return operation;
  }, []);

  useEffect(() => {
    const current = session.current;
    if (!ready || recovery || !current?.loaded || current.conflicted || options.paused) return;
    if (draftFingerprint(value) === current.fingerprint) {
      if (!current.writing) {
        setStatus(current.revision === null ? 'idle' : 'saved');
        setError(null);
      }
      return;
    }
    setStatus('saving');
    clearTimeout(saveTimer.current);
    saveTimer.current = setTimeout(() => { void saveLatest().catch(() => undefined); }, 500);
    return () => clearTimeout(saveTimer.current);
  }, [value, ready, recovery, saveLatest, saveEpoch, options.paused]);

  const flush = useCallback(async () => {
    clearTimeout(saveTimer.current);
    if (!ready || recovery || options.paused) throw new Error('Finish local draft recovery before switching work.');
    const current = session.current;
    for (;;) {
      const saved = await saveLatest();
      if (session.current !== current || !current) throw new Error('The active draft changed while saving. Retry against the current document.');
      if (draftFingerprint(latest.current) === current.fingerprint) return saved;
    }
  }, [ready, recovery, options.paused, saveLatest]);

  useEffect(() => {
    const preserve = () => {
      if (document.visibilityState === 'hidden' && ready && !recovery && !options.paused) {
        void saveLatest().catch(() => undefined);
      }
    };
    document.addEventListener('visibilitychange', preserve);
    return () => document.removeEventListener('visibilitychange', preserve);
  }, [ready, recovery, options.paused, saveLatest]);

  useEffect(() => {
    const protectUnsavedWork = (event: BeforeUnloadEvent) => {
      const current = session.current;
      if (ready && current && (current.writing || current.conflicted || draftFingerprint(latest.current) !== current.fingerprint)) {
        event.preventDefault();
        event.returnValue = '';
      }
    };
    window.addEventListener('beforeunload', protectUnsavedWork);
    return () => window.removeEventListener('beforeunload', protectUnsavedWork);
  }, [ready]);

  const restoreDraft = useCallback(() => {
    const current = session.current;
    if (!recovery || !current) return;
    try {
      restore.current(recovery.document);
      current.fingerprint = draftFingerprint(recovery.document);
      setSavedAt(recovery.updatedAt);
      setSavedDraft(recovery);
      setError(null);
      setStatus('saved');
      setRecovery(null);
      setReady(true);
    } catch (cause) {
      setError(cause instanceof Error ? cause : new Error(String(cause)));
      setStatus('error');
    }
  }, [recovery]);

  const startFresh = useCallback(async () => {
    const current = session.current;
    if (!recovery || !current) return;
    try {
      await discardDraft(current.key, recovery.revision);
      current.revision = null;
      current.saved = null;
      current.fingerprint = draftFingerprint(latest.current);
      setRecovery(null);
      setSavedDraft(null);
      setError(null);
      setStatus('idle');
      setReady(true);
    } catch (cause) {
      setError(cause instanceof Error ? cause : new Error(String(cause)));
      setStatus('error');
    }
  }, [recovery]);

  const retry = useCallback(() => {
    if (!session.current?.loaded) setRetryCount(count => count + 1);
    else void saveLatest().catch(() => undefined);
  }, [saveLatest]);

  return { status, error, recovery, ready, savedAt, savedDraft, restoreDraft, startFresh, retry, flush };
}
