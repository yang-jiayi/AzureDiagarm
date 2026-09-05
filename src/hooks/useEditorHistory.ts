import { useCallback, useLayoutEffect, useRef, useState } from 'react';
import { EditorHistory, editorFingerprint, type EditorDocument } from '../services/editorHistory';
import type { Node } from 'reactflow';
import type { NodePricingConfig } from '../types/pricing';

export const EDITOR_GESTURE_EVENT = 'azurediagarm:editor-gesture';

export function setEditorGesture(active: boolean): void {
  window.dispatchEvent(new CustomEvent(EDITOR_GESTURE_EVENT, { detail: active }));
}

function editingText(): boolean {
  const element = document.activeElement;
  return element instanceof HTMLTextAreaElement
    || (element instanceof HTMLInputElement && !['checkbox', 'radio', 'button', 'file'].includes(element.type))
    || (element instanceof HTMLElement && element.isContentEditable);
}

export function useEditorHistory(
  value: EditorDocument,
  onRestore: (document: EditorDocument) => void,
  enabled: boolean,
  suspended = false,
) {
  const history = useRef<EditorHistory>();
  history.current ??= new EditorHistory(value);
  const restoreRef = useRef(onRestore);
  const restoredFingerprint = useRef<string | null>(null);
  const wasEnabled = useRef(false);
  const explicitGesture = useRef(false);
  const latest = useRef(value);
  latest.current = value;
  const [, setRenderVersion] = useState(0);
  const refresh = useCallback(() => setRenderVersion(value => value + 1), []);

  useLayoutEffect(() => { restoreRef.current = onRestore; }, [onRestore]);

  useLayoutEffect(() => {
    const current = history.current!;
    if (suspended) return;
    if (!enabled) {
      wasEnabled.current = false;
      return;
    }
    if (!wasEnabled.current) {
      current.reset(value);
      wasEnabled.current = true;
      refresh();
      return;
    }
    if (restoredFingerprint.current === editorFingerprint(value)) {
      restoredFingerprint.current = null;
      return;
    }
    const transient = explicitGesture.current || value.nodes.some(node => node.dragging || node.resizing) || editingText();
    if (current.record(value, transient)) refresh();
  }, [value, enabled, suspended, refresh]);

  useLayoutEffect(() => {
    const finish = () => {
      explicitGesture.current = false;
      if (history.current!.finishGesture()) refresh();
    };
    const gesture = (event: Event) => {
      explicitGesture.current = Boolean((event as CustomEvent<boolean>).detail);
      if (!explicitGesture.current) finish();
    };
    window.addEventListener('blur', finish);
    window.addEventListener(EDITOR_GESTURE_EVENT, gesture);
    document.addEventListener('focusout', finish);
    return () => {
      window.removeEventListener('blur', finish);
      window.removeEventListener(EDITOR_GESTURE_EVENT, gesture);
      document.removeEventListener('focusout', finish);
    };
  }, [refresh]);

  const move = useCallback((direction: 'undo' | 'redo') => {
    if (!enabled || suspended) return;
    const result = history.current![direction]();
    if (!result) return;
    restoredFingerprint.current = editorFingerprint(result);
    restoreRef.current(result);
    refresh();
  }, [enabled, suspended, refresh]);

  const reset = useCallback((next = latest.current) => {
    history.current!.reset(next);
    restoredFingerprint.current = editorFingerprint(next);
    explicitGesture.current = false;
    refresh();
  }, [refresh]);

  const enrichPricing = useCallback((node: Node, pricing: NodePricingConfig) => {
    if (history.current!.enrichPricing(node, pricing)) refresh();
  }, [refresh]);

  return {
    revision: history.current.revision,
    canUndo: enabled && !suspended && history.current.canUndo,
    canRedo: enabled && !suspended && history.current.canRedo,
    undo: useCallback(() => move('undo'), [move]),
    redo: useCallback(() => move('redo'), [move]),
    reset,
    enrichPricing,
    getRevision: useCallback(() => history.current!.revision, []),
  };
}
