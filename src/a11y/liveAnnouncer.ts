/*
 * Global screen-reader announcer.
 *
 * Several important state changes in this app happen far away from the control
 * that triggered them: a keyboard-created connection appears on the canvas, an
 * AI generation finishes minutes after the button was pressed, an export
 * completes asynchronously. None of those move focus, so without an ARIA live
 * region a screen-reader user gets no feedback at all (WCAG 4.1.3 Status
 * Messages).
 *
 * The store lives at module scope rather than in React context so that
 * non-component code (services, event handlers, hooks used outside the tree)
 * can announce without threading a provider through every call site.
 */

export type Politeness = 'polite' | 'assertive';

export interface LiveMessage {
  /** Monotonic id — re-announces identical text, which assistive tech otherwise skips. */
  id: number;
  text: string;
}

type Listener = () => void;

const listeners = new Set<Listener>();

let nextId = 1;

const state: Record<Politeness, LiveMessage> = {
  polite: { id: 0, text: '' },
  assertive: { id: 0, text: '' },
};

function emit(): void {
  listeners.forEach(listener => listener());
}

export function subscribeToAnnouncements(listener: Listener): () => void {
  listeners.add(listener);
  return () => {
    listeners.delete(listener);
  };
}

export function getAnnouncement(politeness: Politeness): LiveMessage {
  return state[politeness];
}

/**
 * Queue `text` for the given live region. Empty/whitespace-only text is ignored
 * so callers can pass a conditional string without guarding.
 */
export function announce(text: string, politeness: Politeness = 'polite'): void {
  const trimmed = typeof text === 'string' ? text.trim() : '';
  if (!trimmed) return;
  state[politeness] = { id: nextId++, text: trimmed };
  emit();
}

/** Test helper: clears both regions without emitting a spoken message. */
export function resetAnnouncements(): void {
  state.polite = { id: 0, text: '' };
  state.assertive = { id: 0, text: '' };
  emit();
}
