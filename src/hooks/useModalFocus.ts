import { useEffect, useRef } from 'react';

const FOCUSABLE_SELECTOR = [
  'a[href]',
  'area[href]',
  'button:not([disabled])',
  'input:not([disabled]):not([type="hidden"])',
  'select:not([disabled])',
  'textarea:not([disabled])',
  'iframe',
  'object',
  'embed',
  '[contenteditable="true"]',
  '[tabindex]:not([tabindex="-1"])',
].join(',');

function getFocusableElements(container: HTMLElement): HTMLElement[] {
  return Array.from(container.querySelectorAll<HTMLElement>(FOCUSABLE_SELECTOR))
    .filter((element) => (
      element.getAttribute('aria-hidden') !== 'true'
      && window.getComputedStyle(element).visibility !== 'hidden'
      && element.getClientRects().length > 0
    ));
}

/**
 * Keeps keyboard focus inside an active modal and restores it to the opener
 * when the modal closes. Dialog containers should also have tabIndex={-1}.
 */
export function useModalFocus<T extends HTMLElement>(
  active: boolean,
  returnFocusTarget: HTMLElement | null = null,
) {
  const dialogRef = useRef<T>(null);

  useEffect(() => {
    if (!active) return;

    const activeElement = document.activeElement;
    const returnFocus = returnFocusTarget?.isConnected
      ? returnFocusTarget
      : activeElement instanceof HTMLElement && activeElement !== document.body
        ? activeElement
        : null;
    const focusFrame = window.requestAnimationFrame(() => {
      dialogRef.current?.focus();
    });

    const handleKeyDown = (event: KeyboardEvent) => {
      if (event.key !== 'Tab') return;
      const dialog = dialogRef.current;
      if (!dialog) return;

      const focusable = getFocusableElements(dialog);
      if (focusable.length === 0) {
        event.preventDefault();
        dialog.focus();
        return;
      }

      const activeElement = document.activeElement;
      const currentIndex = focusable.indexOf(activeElement as HTMLElement);
      const movingBackward = event.shiftKey;
      const shouldWrapBackward = movingBackward && currentIndex <= 0;
      const shouldWrapForward = !movingBackward && currentIndex === focusable.length - 1;
      const focusOutsideDialog = !(activeElement instanceof Node) || !dialog.contains(activeElement);

      if (focusOutsideDialog || shouldWrapBackward || shouldWrapForward) {
        event.preventDefault();
        focusable[movingBackward ? focusable.length - 1 : 0].focus();
      }
    };

    document.addEventListener('keydown', handleKeyDown, true);
    return () => {
      window.cancelAnimationFrame(focusFrame);
      document.removeEventListener('keydown', handleKeyDown, true);
      window.requestAnimationFrame(() => {
        const focusTarget = returnFocus?.isConnected
          ? returnFocus
          : document.querySelector<HTMLElement>('[data-modal-focus-fallback]');
        focusTarget?.focus();
      });
    };
  }, [active, returnFocusTarget]);

  return dialogRef;
}
