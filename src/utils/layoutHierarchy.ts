// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

interface HierarchyNode {
  id: string;
  type?: string;
  parentNode?: string;
}

interface SizedHierarchyNode extends HierarchyNode {
  position?: { x: number; y: number };
  width?: number | null;
  height?: number | null;
  style?: unknown;
}

export interface NestedHierarchyLayoutUnit {
  rootId: string;
  surrogateId: string;
  width: number;
  height: number;
  offsetX: number;
  offsetY: number;
}

function numericDimension(value: unknown): number | undefined {
  if (typeof value === 'number' && Number.isFinite(value) && value > 0) return value;
  if (typeof value === 'string') {
    const parsed = Number.parseFloat(value);
    if (Number.isFinite(parsed) && parsed > 0) return parsed;
  }
  return undefined;
}

export function layoutNodeDimensions(node: SizedHierarchyNode): {
  width: number;
  height: number;
} {
  const style = node.style && typeof node.style === 'object'
    ? node.style as Record<string, unknown>
    : {};
  return {
    width: numericDimension(node.width)
      ?? numericDimension(style.width)
      ?? (node.type === 'groupNode' ? 420 : 180),
    height: numericDimension(node.height)
      ?? numericDimension(style.height)
      ?? (node.type === 'groupNode' ? 260 : 100),
  };
}

function absolutePosition<T extends SizedHierarchyNode>(
  node: T,
  nodesById: Map<string, T>,
  cache: Map<string, { x: number; y: number }>,
  visiting = new Set<string>(),
): { x: number; y: number } {
  const cached = cache.get(node.id);
  if (cached) return cached;
  const local = node.position ?? { x: 0, y: 0 };
  if (!node.parentNode || visiting.has(node.id)) {
    cache.set(node.id, local);
    return local;
  }
  const parent = nodesById.get(node.parentNode);
  if (!parent) {
    cache.set(node.id, local);
    return local;
  }
  const nextVisiting = new Set(visiting);
  nextVisiting.add(node.id);
  const parentPosition = absolutePosition(parent, nodesById, cache, nextVisiting);
  const result = {
    x: parentPosition.x + local.x,
    y: parentPosition.y + local.y,
  };
  cache.set(node.id, result);
  return result;
}

/**
 * Returns every node that belongs to a nested-group hierarchy.
 *
 * Automatic layout engines currently understand one group level. Preserving
 * the complete nested subtree is safer than reinterpreting nested relative
 * coordinates as top-level positions.
 */
export function collectNestedHierarchyNodeIds<T extends HierarchyNode>(
  nodes: T[],
): Set<string> {
  const nodesById = new Map(nodes.map(node => [node.id, node]));
  const protectedGroupIds = new Set<string>();

  for (const node of nodes) {
    if (node.type !== 'groupNode' || !node.parentNode) continue;
    const parent = nodesById.get(node.parentNode);
    if (parent?.type !== 'groupNode') continue;

    let current: T | undefined = node;
    const visiting = new Set<string>();
    while (current?.type === 'groupNode' && !visiting.has(current.id)) {
      protectedGroupIds.add(current.id);
      visiting.add(current.id);
      current = current.parentNode ? nodesById.get(current.parentNode) : undefined;
    }
  }

  if (protectedGroupIds.size === 0) return new Set();

  const protectedNodeIds = new Set(protectedGroupIds);
  let added = true;
  while (added) {
    added = false;
    for (const node of nodes) {
      if (
        node.parentNode
        && protectedNodeIds.has(node.parentNode)
        && !protectedNodeIds.has(node.id)
      ) {
        protectedNodeIds.add(node.id);
        added = true;
      }
    }
  }

  return protectedNodeIds;
}

export function buildNestedHierarchyLayout<T extends SizedHierarchyNode>(
  nodes: T[],
): {
  protectedNodeIds: Set<string>;
  units: NestedHierarchyLayoutUnit[];
  unitByNodeId: Map<string, NestedHierarchyLayoutUnit>;
} {
  const protectedNodeIds = collectNestedHierarchyNodeIds(nodes);
  if (protectedNodeIds.size === 0) {
    return {
      protectedNodeIds,
      units: [],
      unitByNodeId: new Map(),
    };
  }

  const nodesById = new Map(nodes.map(node => [node.id, node]));
  const rootNodes = nodes.filter(node => (
    node.type === 'groupNode'
    && protectedNodeIds.has(node.id)
    && (!node.parentNode || !protectedNodeIds.has(node.parentNode))
  ));
  const rootIds = new Set(rootNodes.map(node => node.id));
  const rootIdByNodeId = new Map<string, string>();
  for (const node of nodes) {
    if (!protectedNodeIds.has(node.id)) continue;
    let current: T | undefined = node;
    const visiting = new Set<string>();
    while (
      current?.parentNode
      && protectedNodeIds.has(current.parentNode)
      && !visiting.has(current.id)
    ) {
      visiting.add(current.id);
      current = nodesById.get(current.parentNode);
    }
    if (current && rootIds.has(current.id)) {
      rootIdByNodeId.set(node.id, current.id);
    }
  }

  const usedIds = new Set(nodesById.keys());
  const positionCache = new Map<string, { x: number; y: number }>();
  const units = rootNodes.map(root => {
    let surrogateId = `__nested_hierarchy__${root.id}`;
    while (usedIds.has(surrogateId)) surrogateId = `_${surrogateId}`;
    usedIds.add(surrogateId);

    const rootPosition = absolutePosition(root, nodesById, positionCache);
    const rootDimensions = layoutNodeDimensions(root);
    let minX = rootPosition.x;
    let minY = rootPosition.y;
    let maxX = rootPosition.x + rootDimensions.width;
    let maxY = rootPosition.y + rootDimensions.height;
    for (const node of nodes) {
      if (rootIdByNodeId.get(node.id) !== root.id) continue;
      const position = absolutePosition(node, nodesById, positionCache);
      const dimensions = layoutNodeDimensions(node);
      minX = Math.min(minX, position.x);
      minY = Math.min(minY, position.y);
      maxX = Math.max(maxX, position.x + dimensions.width);
      maxY = Math.max(maxY, position.y + dimensions.height);
    }

    return {
      rootId: root.id,
      surrogateId,
      width: maxX - minX,
      height: maxY - minY,
      offsetX: minX - rootPosition.x,
      offsetY: minY - rootPosition.y,
    };
  });
  const unitByRootId = new Map(units.map(unit => [unit.rootId, unit]));
  const unitByNodeId = new Map<string, NestedHierarchyLayoutUnit>();

  for (const [nodeId, rootId] of rootIdByNodeId) {
    const unit = unitByRootId.get(rootId);
    if (unit) unitByNodeId.set(nodeId, unit);
  }

  return { protectedNodeIds, units, unitByNodeId };
}
