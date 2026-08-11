/*
 * Keyboard-only edge creation.
 *
 * React Flow's connection UX is drag-only: the source and target handles are
 * 8px dots that must be dragged between with a pointer. That made building a
 * diagram impossible without a mouse (WCAG 2.1.1 Keyboard) — a node could be
 * focused, moved and renamed from the keyboard, but never connected.
 *
 * This module holds the "pending source" for a two-step keyboard connection:
 * press C on the source node, focus the target node, press C again. The state
 * lives at module scope because the two presses happen in two different
 * `AzureNode` instances, which have no common React ancestor other than the
 * canvas itself.
 *
 * Completion is dispatched as a DOM event rather than mutating edges here, so
 * that App's existing `onConnect` remains the single place that decides what a
 * new edge looks like (type, marker, label styling, flow animation, …).
 */

import { useSyncExternalStore } from 'react';

export const KEYBOARD_CONNECT_EVENT = 'azd:keyboard-connect';

export interface KeyboardConnectDetail {
  source: string;
  target: string;
}

export interface PendingConnection {
  nodeId: string;
  label: string;
}

type Listener = () => void;

const listeners = new Set<Listener>();

let pending: PendingConnection | null = null;

function emit(): void {
  listeners.forEach(listener => listener());
}

function subscribe(listener: Listener): () => void {
  listeners.add(listener);
  return () => {
    listeners.delete(listener);
  };
}

function getPending(): PendingConnection | null {
  return pending;
}

export function beginKeyboardConnection(nodeId: string, label: string): void {
  pending = { nodeId, label };
  emit();
}

export function cancelKeyboardConnection(): void {
  if (!pending) return;
  pending = null;
  emit();
}

export function getPendingConnection(): PendingConnection | null {
  return pending;
}

/**
 * Finishes the pending connection at `targetId`.
 *
 * Returns the source node id when an edge was requested, or `null` when the
 * attempt was rejected (no pending source, or the target is the source — React
 * Flow silently drops self-connections, which would look like a dead key).
 */
export function completeKeyboardConnection(targetId: string): string | null {
  const source = pending;
  if (!source) return null;
  pending = null;
  emit();
  if (source.nodeId === targetId) return null;

  if (typeof window !== 'undefined') {
    window.dispatchEvent(new CustomEvent<KeyboardConnectDetail>(KEYBOARD_CONNECT_EVENT, {
      detail: { source: source.nodeId, target: targetId },
    }));
  }
  return source.nodeId;
}

/** Reactive read of the pending source, used to highlight the origin node. */
export function usePendingConnection(): PendingConnection | null {
  return useSyncExternalStore(subscribe, getPending, getPending);
}
