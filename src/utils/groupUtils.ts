// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

import { Node } from 'reactflow';

/**
 * Constants shared by GroupNode and App-level group operations.
 */
export const GROUP_PADDING = 40;
export const GROUP_HEADER_HEIGHT = 50;

export type GroupLayoutSnapshot = Map<string, {
  position: { x: number; y: number };
  width?: number;
  height?: number;
}>;

function numericDimension(value: unknown): number | undefined {
  if (typeof value === 'number' && Number.isFinite(value) && value > 0) return value;
  if (typeof value === 'string') {
    const parsed = Number.parseFloat(value);
    if (Number.isFinite(parsed) && parsed > 0) return parsed;
  }
  return undefined;
}

function nodeDimensions(node: Node): { width: number; height: number } {
  const style = node.style as Record<string, unknown> | undefined;
  return {
    width: numericDimension(node.width)
      ?? numericDimension(style?.width)
      ?? (node.type === 'groupNode' ? 400 : 160),
    height: numericDimension(node.height)
      ?? numericDimension(style?.height)
      ?? (node.type === 'groupNode' ? 300 : 100),
  };
}

function nodeDepth(node: Node, nodesById: Map<string, Node>): number {
  let depth = 0;
  let current = node;
  const visiting = new Set<string>();
  while (current.parentNode && !visiting.has(current.id)) {
    visiting.add(current.id);
    const parent = nodesById.get(current.parentNode);
    if (!parent) break;
    depth += 1;
    current = parent;
  }
  return depth;
}

export function captureGroupLayout(nodes: Node[]): GroupLayoutSnapshot {
  return new Map(nodes.map(node => {
    const dimensions = node.type === 'groupNode' ? nodeDimensions(node) : {};
    return [
      node.id,
      {
        position: { ...node.position },
        ...dimensions,
      },
    ];
  }));
}

export function restoreGroupLayout(
  nodes: Node[],
  snapshot: GroupLayoutSnapshot,
): Node[] {
  return nodes.map(node => {
    const saved = snapshot.get(node.id);
    if (!saved) return node;
    return {
      ...node,
      position: { ...saved.position },
      positionAbsolute: undefined,
      style: node.type === 'groupNode'
        ? {
            ...node.style,
            width: saved.width,
            height: saved.height,
          }
        : node.style,
    };
  });
}

/**
 * Compute the bounding box of a group's children and return
 * the updated node array with the group tightly fitted and
 * children repositioned relative to the new origin without
 * changing their absolute canvas positions.
 *
 * Returns null if the group has no children (no-op).
 */
export function fitGroupToContent(
  allNodes: Node[],
  groupId: string
): Node[] | null {
  const children = allNodes.filter(n => n.parentNode === groupId);
  if (children.length === 0) return null;

  let minX = Infinity,
    minY = Infinity,
    maxX = -Infinity,
    maxY = -Infinity;

  children.forEach(child => {
    const x = child.position.x;
    const y = child.position.y;
    const { width: w, height: h } = nodeDimensions(child);
    minX = Math.min(minX, x);
    minY = Math.min(minY, y);
    maxX = Math.max(maxX, x + w);
    maxY = Math.max(maxY, y + h);
  });

  const offsetX = minX - GROUP_PADDING;
  const offsetY = minY - GROUP_PADDING - GROUP_HEADER_HEIGHT;
  const newWidth = maxX - minX + GROUP_PADDING * 2;
  const newHeight = maxY - minY + GROUP_PADDING * 2 + GROUP_HEADER_HEIGHT;

  return allNodes.map(n => {
    if (n.id === groupId) {
      return {
        ...n,
        position: {
          x: n.position.x + offsetX,
          y: n.position.y + offsetY,
        },
        style: { ...n.style, width: newWidth, height: newHeight },
      };
    }
    if (n.parentNode === groupId) {
      return {
        ...n,
        position: { x: n.position.x - offsetX, y: n.position.y - offsetY },
      };
    }
    return n;
  });
}

/**
 * Apply fitGroupToContent to every group node in the array.
 * Returns the modified node array.
 */
export function fitAllGroupsToContent(nodes: Node[]): Node[] {
  const nodesById = new Map(nodes.map(node => [node.id, node]));
  const groupIds = nodes
    .filter(node => node.type === 'groupNode')
    .sort((left, right) => nodeDepth(right, nodesById) - nodeDepth(left, nodesById))
    .map(node => node.id);

  let result = [...nodes];
  for (const gid of groupIds) {
    const updated = fitGroupToContent(result, gid);
    if (updated) result = updated;
  }
  return result;
}

function buildAbsolutePositions(nodes: Node[]): Map<string, { x: number; y: number }> {
  const nodesById = new Map(nodes.map(node => [node.id, node]));
  const positions = new Map<string, { x: number; y: number }>();

  const resolve = (node: Node, visiting = new Set<string>()): { x: number; y: number } => {
    const cached = positions.get(node.id);
    if (cached) return cached;
    if (visiting.has(node.id)) return node.position;

    const nextVisiting = new Set(visiting);
    nextVisiting.add(node.id);
    const parent = node.parentNode ? nodesById.get(node.parentNode) : undefined;
    const position = parent
      ? {
          x: resolve(parent, nextVisiting).x + node.position.x,
          y: resolve(parent, nextVisiting).y + node.position.y,
        }
      : node.position;
    positions.set(node.id, position);
    return position;
  };

  nodes.forEach(node => resolve(node));
  return positions;
}

/**
 * Detach direct children from groups that are being deleted while preserving
 * each child's absolute canvas position.
 */
export function detachChildrenFromGroups(nodes: Node[], groupIds: Iterable<string>): Node[] {
  const deletedGroups = new Set(groupIds);
  if (deletedGroups.size === 0) return nodes;
  const absolutePositions = buildAbsolutePositions(nodes);

  return nodes.map(node => {
    if (!node.parentNode || !deletedGroups.has(node.parentNode)) return node;
    return {
      ...node,
      parentNode: undefined,
      extent: undefined,
      position: absolutePositions.get(node.id) || node.position,
    };
  });
}

/**
 * Remove selected nodes without implicitly removing or hiding the contents of
 * a selected group.
 */
export function deleteNodesPreservingGroupChildren(nodes: Node[], nodeIds: Iterable<string>): Node[] {
  const deletedIds = new Set(nodeIds);
  const deletedGroupIds = nodes
    .filter(node => deletedIds.has(node.id) && node.type === 'groupNode')
    .map(node => node.id);
  return detachChildrenFromGroups(nodes, deletedGroupIds)
    .filter(node => !deletedIds.has(node.id));
}

/**
 * Return the requested nodes plus every nested descendant they contain.
 */
export function collectNodeAndDescendantIds(
  nodes: Node[],
  rootIds: Iterable<string>,
): Set<string> {
  const collected = new Set(rootIds);
  let added = true;

  while (added) {
    added = false;
    for (const node of nodes) {
      if (
        node.parentNode
        && collected.has(node.parentNode)
        && !collected.has(node.id)
      ) {
        collected.add(node.id);
        added = true;
      }
    }
  }

  return collected;
}
