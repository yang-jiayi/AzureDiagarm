import { useCallback } from 'react';
import { useReactFlow } from 'reactflow';

import { announce } from '../a11y/liveAnnouncer';
import { useLanguage } from '../i18n/LanguageContext';
import { localize } from '../i18n/localization';
import {
  beginKeyboardConnection,
  cancelKeyboardConnection,
  completeKeyboardConnection,
  getPendingConnection,
} from './useKeyboardConnection';

const ARROW_DELTAS: Record<string, { x: number; y: number }> = {
  ArrowUp: { x: 0, y: -5 },
  ArrowRight: { x: 5, y: 0 },
  ArrowDown: { x: 0, y: 5 },
  ArrowLeft: { x: -5, y: 0 },
};

export function useNodeKeyboardInteraction(
  nodeId: string,
  startEditing: () => void,
  nodeLabel?: string,
) {
  const { setNodes, setEdges, getNode } = useReactFlow();
  const { language } = useLanguage();

  const describe = useCallback((id: string, fallback?: string): string => {
    const node = getNode(id);
    const label = (node?.data?.label as string | undefined) || fallback || id;
    return String(label);
  }, [getNode]);

  const handleFocus = useCallback((event: React.FocusEvent<HTMLElement>) => {
    if (
      event.relatedTarget instanceof Node
      && event.currentTarget.contains(event.relatedTarget)
    ) return;

    setNodes(nodes => nodes.map(node => ({
      ...node,
      selected: node.id === nodeId,
    })));
    setEdges(edges => edges.map(edge => ({ ...edge, selected: false })));
  }, [nodeId, setEdges, setNodes]);

  // Two-step keyboard connection: C on the source, then C on the target.
  // Without this there is no pointer-free way to draw an edge, because React
  // Flow's connection handles are drag-only.
  const handleConnectKey = useCallback(() => {
    const pending = getPendingConnection();
    const cancelled = localize(language, {
      en: 'Connection cancelled.',
      ja: '接続を取り消しました。',
    });

    if (!pending) {
      const label = describe(nodeId, nodeLabel);
      beginKeyboardConnection(nodeId, label);
      announce(localize(language, {
        en: `Connection started from ${label}. Move to another node and press C to connect, or press Escape to cancel.`,
        ja: `${label} から接続を開始しました。別のノードへ移動して C を押すと接続、Escape で取り消します。`,
      }));
      return;
    }

    if (pending.nodeId === nodeId) {
      cancelKeyboardConnection();
      announce(cancelled);
      return;
    }

    const sourceLabel = pending.label;
    const targetLabel = describe(nodeId, nodeLabel);
    const source = completeKeyboardConnection(nodeId);
    announce(source
      ? localize(language, {
          en: `Connected ${sourceLabel} to ${targetLabel}.`,
          ja: `${sourceLabel} から ${targetLabel} へ接続しました。`,
        })
      : cancelled);
  }, [describe, language, nodeId, nodeLabel]);

  const handleKeyDown = useCallback((event: React.KeyboardEvent<HTMLElement>) => {
    if (event.key === 'Enter' || event.key === ' ' || event.key === 'F2') {
      event.preventDefault();
      event.stopPropagation();
      startEditing();
      return;
    }

    // Guard the modifiers so Ctrl/Cmd+C still copies.
    if ((event.key === 'c' || event.key === 'C') && !event.ctrlKey && !event.metaKey && !event.altKey) {
      event.preventDefault();
      event.stopPropagation();
      handleConnectKey();
      return;
    }

    const delta = ARROW_DELTAS[event.key];
    if (delta) {
      const multiplier = event.shiftKey ? 4 : 1;
      event.preventDefault();
      event.stopPropagation();
      setNodes(nodes => nodes.map(node => (
        node.id === nodeId
          ? {
              ...node,
              position: {
                x: node.position.x + delta.x * multiplier,
                y: node.position.y + delta.y * multiplier,
              },
            }
          : node
      )));
    } else if (event.key === 'Escape') {
      event.stopPropagation();
      if (getPendingConnection()) {
        cancelKeyboardConnection();
        announce(localize(language, {
          en: 'Connection cancelled.',
          ja: '接続を取り消しました。',
        }));
        return;
      }
      setNodes(nodes => nodes.map(node => (
        node.id === nodeId ? { ...node, selected: false } : node
      )));
    }
  }, [handleConnectKey, language, nodeId, setNodes, startEditing]);

  return { handleFocus, handleKeyDown };
}
