import { useEffect, useRef, type RefObject } from 'react';

export interface EscapeKeyOptions {
  scopeRef?: RefObject<HTMLElement>;
  initialFocusRef?: RefObject<HTMLElement>;
  closeOnEscape?: boolean;
  returnFocusTarget?: HTMLElement | null;
}

interface DismissLayer {
  options: () => EscapeKeyOptions;
  dismiss: () => void;
  canDismiss: () => boolean;
  opener: HTMLElement | null;
  scope: HTMLElement | null;
}

const layers: DismissLayer[] = [];
const focusableSelector = [
  'a[href]', 'area[href]', 'button', 'input:not([type="hidden"])', 'select',
  'textarea', '[tabindex]', '[contenteditable="true"]', 'summary',
  'audio[controls]', 'video[controls]', 'iframe', 'object', 'embed',
].join(',');

function isVisible(element: HTMLElement): boolean {
  return element.isConnected && !element.closest('[hidden], [inert], [aria-hidden="true"]') &&
    element.getClientRects().length > 0 && getComputedStyle(element).visibility !== 'hidden';
}

function isFocusable(element: HTMLElement): boolean {
  return isVisible(element) && !element.matches(':disabled');
}

function getScope(layer: DismissLayer): HTMLElement | null {
  return layer.scope;
}

function topLayer(includeHidden = false): DismissLayer | undefined {
  const dialogs = layers.filter(layer => {
    const scope = getScope(layer);
    return scope && (includeHidden || isVisible(scope));
  });
  // A nested dialog wins even when React runs the child's effect first.
  return dialogs.reverse().find(layer => !dialogs.some(other =>
    other !== layer && getScope(layer)?.contains(getScope(other)),
  )) ?? [...layers].reverse().find(layer => !layer.options().scopeRef && layer.canDismiss());
}

function escapeLayer(top: DismissLayer): DismissLayer | undefined {
  if (top.canDismiss()) return top;
  // Legacy callers pair useModalFocus(active, returnTarget) with a separate
  // useEscapeKey. Only that scope's companion may dismiss it, never its parent.
  const companions: DismissLayer[] = [];
  for (const layer of layers.slice(layers.indexOf(top) + 1)) {
    if (layer.options().scopeRef) break;
    if (layer.canDismiss()) companions.push(layer);
  }
  return companions[companions.length - 1];
}

function tabStops(scope: HTMLElement): HTMLElement[] {
  return Array.from(scope.querySelectorAll<HTMLElement>(focusableSelector))
    .filter(element => element.tabIndex >= 0 && isFocusable(element))
    .filter((element, _, elements) => {
      if (!(element instanceof HTMLInputElement) || element.type !== 'radio' || !element.name) return true;
      const group = elements.filter((candidate): candidate is HTMLInputElement =>
        candidate instanceof HTMLInputElement && candidate.type === 'radio' &&
        candidate.name === element.name && candidate.form === element.form,
      );
      return element === (group.find(candidate => candidate.checked) ?? group[0]);
    })
    .sort((a, b) => (a.tabIndex || Infinity) - (b.tabIndex || Infinity));
}

function focusScope(layer: DismissLayer, last = false, initial = true): void {
  const scope = getScope(layer);
  if (!scope) return;
  const preferred = layer.options().initialFocusRef?.current ??
    scope.querySelector<HTMLElement>('[data-modal-initial-focus], [autofocus]');
  const stops = tabStops(scope);
  const target = last ? stops[stops.length - 1] : initial && preferred && scope.contains(preferred) && isFocusable(preferred)
    ? preferred : stops[0];
  (target ?? scope).focus({ preventScroll: true });
}

function handleKeyDown(event: KeyboardEvent): void {
  const top = topLayer();
  if (!top) return;
  if (event.key === 'Escape' && !event.isComposing) {
    event.preventDefault();
    event.stopImmediatePropagation();
    if (top.options().closeOnEscape !== false) {
      const dismissLayer = escapeLayer(top);
      if (dismissLayer?.options().closeOnEscape !== false) dismissLayer?.dismiss();
    }
    return;
  }
  const scope = getScope(top);
  if (event.key !== 'Tab' || !scope) return;
  const stops = tabStops(scope);
  const index = stops.indexOf(document.activeElement as HTMLElement);
  if (index < 0 || (event.shiftKey ? index === 0 : index === stops.length - 1)) {
    event.preventDefault();
    focusScope(top, event.shiftKey, false);
  }
}

function handleFocusIn(event: FocusEvent): void {
  const top = topLayer();
  const scope = top && getScope(top);
  if (top && scope && event.target instanceof Node && !scope.contains(event.target)) {
    focusScope(top);
  }
}

export function useEscapeKey(active: boolean, onEscape?: () => void, options: EscapeKeyOptions = {}): void {
  const latest = useRef({ onEscape, options });
  const opening = useRef(false);
  const opener = useRef<HTMLElement | null>(null);
  latest.current = { onEscape, options };
  if (active && !opening.current && typeof document !== 'undefined') {
    // Capture before React commits descendants with autoFocus.
    opener.current = document.activeElement instanceof HTMLElement ? document.activeElement : null;
  }
  opening.current = active;

  useEffect(() => {
    if (!active) return;
    const layer: DismissLayer = {
      options: () => latest.current.options,
      dismiss: () => latest.current.onEscape?.(),
      canDismiss: () => typeof latest.current.onEscape === 'function',
      opener: latest.current.options.returnFocusTarget ?? opener.current,
      scope: latest.current.options.scopeRef?.current ?? null,
    };
    const scope = getScope(layer);
    const addedTabIndex = scope && !scope.hasAttribute('tabindex');
    if (addedTabIndex) scope.tabIndex = -1;
    layers.push(layer);
    if (layers.length === 1) {
      document.addEventListener('keydown', handleKeyDown, true);
      document.addEventListener('focusin', handleFocusIn, true);
    }
    queueMicrotask(() => {
      if (topLayer() === layer && scope && !scope.contains(document.activeElement)) focusScope(layer);
    });
    return () => {
      const wasTop = topLayer(true) === layer;
      layers.splice(layers.indexOf(layer), 1);
      if (!layers.length) {
        document.removeEventListener('keydown', handleKeyDown, true);
        document.removeEventListener('focusin', handleFocusIn, true);
      }
      if (addedTabIndex) scope.removeAttribute('tabindex');
      if (scope && wasTop) {
        queueMicrotask(() => {
          const top = topLayer();
          const nextScope = top && getScope(top);
          if (layer.opener && isFocusable(layer.opener) && (!nextScope || nextScope.contains(layer.opener))) {
            layer.opener.focus({ preventScroll: true });
          } else if (top && nextScope) {
            focusScope(top);
          } else {
            const fallback = document.querySelector<HTMLElement>('[data-modal-focus-fallback]');
            if (fallback && isFocusable(fallback)) fallback.focus({ preventScroll: true });
          }
        });
      }
    };
  }, [active]);
}
