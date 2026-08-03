// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

import type { Edge, Node } from 'reactflow';

function samePosition(
  left: { x: number; y: number } | undefined,
  right: { x: number; y: number } | undefined,
): boolean {
  return left?.x === right?.x && left?.y === right?.y;
}

function dimension(style: Node['style'], name: 'width' | 'height'): unknown {
  return style && typeof style === 'object' ? style[name] : undefined;
}

function geometryIsUnchanged(current: Node, source: Node): boolean {
  return samePosition(current.position, source.position)
    && current.parentNode === source.parentNode
    && current.extent === source.extent
    && current.width === source.width
    && current.height === source.height
    && dimension(current.style, 'width') === dimension(source.style, 'width')
    && dimension(current.style, 'height') === dimension(source.style, 'height');
}

function mergeLayoutDimensions(current: Node, laidOut: Node): Node['style'] {
  const nextStyle = { ...(current.style ?? {}) };
  const laidOutStyle = laidOut.style;
  if (!laidOutStyle || typeof laidOutStyle !== 'object') return nextStyle;

  if (laidOutStyle.width === undefined) delete nextStyle.width;
  else nextStyle.width = laidOutStyle.width;
  if (laidOutStyle.height === undefined) delete nextStyle.height;
  else nextStyle.height = laidOutStyle.height;
  return nextStyle;
}

/**
 * Apply layout-only fields to the current graph. Concurrent additions,
 * deletions, label edits, and selection changes remain intact. A node whose
 * geometry changed while layout was running keeps the newer user geometry.
 */
export function mergeLayoutNodes(
  currentNodes: Node[],
  sourceNodes: Node[],
  laidOutNodes: Node[],
): Node[] {
  const sourceById = new Map(sourceNodes.map(node => [node.id, node]));
  const laidOutById = new Map(laidOutNodes.map(node => [node.id, node]));

  return currentNodes.map((current) => {
    const source = sourceById.get(current.id);
    const laidOut = laidOutById.get(current.id);
    if (!source || !laidOut || !geometryIsUnchanged(current, source)) return current;

    return {
      ...current,
      position: { ...laidOut.position },
      positionAbsolute: undefined,
      parentNode: laidOut.parentNode,
      extent: laidOut.extent,
      style: mergeLayoutDimensions(current, laidOut),
    };
  });
}

function mergeLayoutDataField(
  nextData: Record<string, unknown>,
  currentData: Record<string, unknown>,
  sourceData: Record<string, unknown>,
  laidOutData: Record<string, unknown>,
  name: 'pathStyle' | 'primaryPath',
) {
  if (currentData[name] !== sourceData[name]) return;
  if (laidOutData[name] === undefined) delete nextData[name];
  else nextData[name] = laidOutData[name];
}

/**
 * Merge edge styling produced by layout without recreating deleted edges or
 * replacing labels and other edge data edited while layout was running.
 */
export function mergeLayoutEdges(
  currentEdges: Edge[],
  sourceEdges: Edge[],
  laidOutEdges: Edge[],
): Edge[] {
  const sourceById = new Map(sourceEdges.map(edge => [edge.id, edge]));
  const laidOutById = new Map(laidOutEdges.map(edge => [edge.id, edge]));

  return currentEdges.map((current) => {
    const source = sourceById.get(current.id);
    const laidOut = laidOutById.get(current.id);
    if (
      !source
      || !laidOut
      || current.source !== source.source
      || current.target !== source.target
    ) {
      return current;
    }

    const currentData = (current.data ?? {}) as Record<string, unknown>;
    const sourceData = (source.data ?? {}) as Record<string, unknown>;
    const laidOutData = (laidOut.data ?? {}) as Record<string, unknown>;
    const nextData = { ...currentData };
    mergeLayoutDataField(nextData, currentData, sourceData, laidOutData, 'pathStyle');
    mergeLayoutDataField(nextData, currentData, sourceData, laidOutData, 'primaryPath');
    return { ...current, data: nextData };
  });
}
