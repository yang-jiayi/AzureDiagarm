// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

import { Node } from 'reactflow';

/**
 * Constants shared by GroupNode and App-level group operations.
 */
export const GROUP_PADDING = 40;
export const GROUP_HEADER_HEIGHT = 50;

/**
 * Compute the bounding box of a group's children and return
 * the updated node array with the group tightly fitted and
 * children repositioned relative to the new origin.
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
    const w = (child.width as number) || 160;
    const h = (child.height as number) || 100;
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
      return { ...n, style: { ...n.style, width: newWidth, height: newHeight } };
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
  const groupIds = nodes
    .filter(n => n.type === 'groupNode')
    .map(n => n.id);

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
