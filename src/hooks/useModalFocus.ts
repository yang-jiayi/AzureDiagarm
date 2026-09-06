import { useRef, type RefObject } from 'react';
import { useEscapeKey, type EscapeKeyOptions } from './useEscapeKey';

export type ModalFocusOptions = Omit<EscapeKeyOptions, 'scopeRef' | 'returnFocusTarget'>;

/** Attach the returned ref to the dialog, not the backdrop or an ordinary side panel. */
export function useModalFocus<T extends HTMLElement = HTMLDivElement>(
  active: boolean,
  onEscapeOrReturnFocus: (() => void) | HTMLElement | null = null,
  options: ModalFocusOptions = {},
): RefObject<T> {
  const dialogRef = useRef<T>(null);
  const onEscape = typeof onEscapeOrReturnFocus === 'function' ? onEscapeOrReturnFocus : undefined;
  useEscapeKey(active, onEscape, {
    ...options,
    scopeRef: dialogRef,
    initialFocusRef: options.initialFocusRef ?? (onEscape ? undefined : dialogRef),
    returnFocusTarget: typeof onEscapeOrReturnFocus === 'function' ? undefined : onEscapeOrReturnFocus,
  });
  return dialogRef;
}
