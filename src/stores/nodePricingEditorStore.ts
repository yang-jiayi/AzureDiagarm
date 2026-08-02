// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

/**
 * Node pricing editor state
 *
 * Which node (if any) currently has the per-node cost editor open. Lives in a
 * module store rather than node data because the editor is a page-level modal:
 * rendering it inside a React Flow node would inherit the canvas pan/zoom
 * transform.
 *
 * AzureNode calls `openNodePricingEditor(id)` from the cost badge; App
 * subscribes via `useNodePricingEditor()` and renders the modal.
 */

import { useState, useEffect } from 'react';

let openNodeId: string | null = null;
let stateVersion = 0;
const listeners: Set<(nodeId: string | null) => void> = new Set();

function notifyListeners() {
  listeners.forEach(listener => listener(openNodeId));
}

/** Open the cost editor for a node. */
export function openNodePricingEditor(nodeId: string): void {
  stateVersion += 1;
  openNodeId = nodeId;
  notifyListeners();
}

/** Close the cost editor. */
export function closeNodePricingEditor(): void {
  stateVersion += 1;
  openNodeId = null;
  notifyListeners();
}

/** Monotonic version used to invalidate pending asynchronous open requests. */
export function getNodePricingEditorStateVersion(): number {
  return stateVersion;
}

/** React hook returning the node id whose editor is open, or null. */
export function useNodePricingEditor(): string | null {
  const [nodeId, setNodeId] = useState<string | null>(openNodeId);

  useEffect(() => {
    const listener = (next: string | null) => setNodeId(next);
    listeners.add(listener);
    return () => {
      listeners.delete(listener);
    };
  }, []);

  return nodeId;
}
