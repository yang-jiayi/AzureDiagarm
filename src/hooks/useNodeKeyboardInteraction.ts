import { useCallback } from 'react';
import { useReactFlow } from 'reactflow';

const ARROW_DELTAS: Record<string, { x: number; y: number }> = {
  ArrowUp: { x: 0, y: -5 },
  ArrowRight: { x: 5, y: 0 },
  ArrowDown: { x: 0, y: 5 },
  ArrowLeft: { x: -5, y: 0 },
};

export function useNodeKeyboardInteraction(
  nodeId: string,
  startEditing: () => void,
) {
  const { setNodes, setEdges } = useReactFlow();

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

  const handleKeyDown = useCallback((event: React.KeyboardEvent<HTMLElement>) => {
    if (event.key === 'Enter' || event.key === ' ' || event.key === 'F2') {
      event.preventDefault();
      event.stopPropagation();
      startEditing();
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
      setNodes(nodes => nodes.map(node => (
        node.id === nodeId ? { ...node, selected: false } : node
      )));
    }
  }, [nodeId, setNodes, startEditing]);

  return { handleFocus, handleKeyDown };
}
