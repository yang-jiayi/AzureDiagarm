// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

import { useCallback, useEffect, useRef, useState } from 'react';
import { DiagramHistory } from '../services/diagramHistoryService';

export function useDiagramHistory<T>(
  state: T,
  onRestore: (state: T) => void,
  options?: {
    delayMs?: number;
    limit?: number;
  },
) {
  const delayMs = options?.delayMs ?? 250;
  const limit = options?.limit ?? 50;
  const stateRef = useRef(state);
  stateRef.current = state;
  const onRestoreRef = useRef(onRestore);
  onRestoreRef.current = onRestore;
  const historyRef = useRef<DiagramHistory<T>>();
  if (!historyRef.current) {
    historyRef.current = new DiagramHistory(state, limit);
  }
  const timerRef = useRef<number | null>(null);
  const [revision, setRevision] = useState(0);

  const clearTimer = useCallback(() => {
    if (timerRef.current !== null) {
      window.clearTimeout(timerRef.current);
      timerRef.current = null;
    }
  }, []);

  const commitCurrent = useCallback(() => {
    clearTimer();
    if (historyRef.current?.record(stateRef.current)) {
      setRevision((current) => current + 1);
    }
  }, [clearTimer]);

  useEffect(() => {
    clearTimer();
    timerRef.current = window.setTimeout(() => {
      timerRef.current = null;
      if (historyRef.current?.record(stateRef.current)) {
        setRevision((current) => current + 1);
      }
    }, delayMs);
    return clearTimer;
  }, [clearTimer, delayMs, state]);

  const undo = useCallback(() => {
    commitCurrent();
    const restored = historyRef.current?.undo();
    if (!restored) return;
    stateRef.current = restored;
    onRestoreRef.current(restored);
    setRevision((current) => current + 1);
  }, [commitCurrent]);

  const redo = useCallback(() => {
    commitCurrent();
    const restored = historyRef.current?.redo();
    if (!restored) return;
    stateRef.current = restored;
    onRestoreRef.current(restored);
    setRevision((current) => current + 1);
  }, [commitCurrent]);

  const reset = useCallback((nextState?: T) => {
    clearTimer();
    const restored = nextState ?? stateRef.current;
    stateRef.current = restored;
    historyRef.current?.reset(restored);
    setRevision((current) => current + 1);
  }, [clearTimer]);

  return {
    canUndo: historyRef.current.canUndo(),
    canRedo: historyRef.current.canRedo(),
    undo,
    redo,
    reset,
    revision,
  };
}
