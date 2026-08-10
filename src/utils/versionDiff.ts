// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

import type { Edge, Node } from 'reactflow';
import { deleteNodesPreservingGroupChildren } from './groupUtils';

export type VersionDiffStatus = 'added' | 'removed' | 'changed';
export type VersionDiffKind = 'node' | 'edge';

export interface VersionDiffItem {
  key: string;
  kind: VersionDiffKind;
  status: VersionDiffStatus;
  id: string;
  label: string;
  before?: Node | Edge;
  after?: Node | Edge;
}

export interface VersionDiff {
  items: VersionDiffItem[];
  nodes: VersionDiffItem[];
  edges: VersionDiffItem[];
  counts: Record<VersionDiffStatus, number>;
}

export interface SelectiveVersionRestoreResult {
  nodes: Node[];
  edges: Edge[];
  appliedKeys: string[];
  skippedKeys: string[];
  autoAddedNodeIds: string[];
  autoRemovedEdgeIds: string[];
}

const OMITTED_RUNTIME_KEYS = new Set([
  'selected',
  'dragging',
  'resizing',
  'positionAbsolute',
  'measured',
  'zIndex',
]);

function isRecord(value: unknown): value is Record<string, unknown> {
  return Boolean(value) && typeof value === 'object' && !Array.isArray(value);
}

function canonicalize(value: unknown, depth = 0): unknown {
  if (depth > 20) return '[depth-limit]';
  if (
    value === null
    || typeof value === 'string'
    || typeof value === 'number'
    || typeof value === 'boolean'
  ) return value;
  if (typeof value === 'function' || typeof value === 'symbol' || value === undefined) {
    return undefined;
  }
  if (Array.isArray(value)) {
    return value
      .map(item => canonicalize(item, depth + 1))
      .filter(item => item !== undefined);
  }
  if (!isRecord(value)) return String(value);

  const normalized: Record<string, unknown> = {};
  for (const entryKey of Object.keys(value).sort()) {
    if (OMITTED_RUNTIME_KEYS.has(entryKey)) continue;
    const normalizedValue = canonicalize(value[entryKey], depth + 1);
    if (normalizedValue !== undefined) normalized[entryKey] = normalizedValue;
  }
  return normalized;
}

function fingerprint(value: unknown): string {
  return JSON.stringify(canonicalize(value));
}

function nodeLabel(node: Node | undefined, fallback: string): string {
  if (!node) return fallback;
  const data = isRecord(node.data) ? node.data : {};
  const label = data.label ?? data.serviceName;
  return typeof label === 'string' && label.trim() ? label.trim() : fallback;
}

function edgeLabel(
  edge: Edge | undefined,
  nodesById: Map<string, Node>,
  fallback: string,
): string {
  if (!edge) return fallback;
  const source = nodeLabel(nodesById.get(edge.source), edge.source);
  const target = nodeLabel(nodesById.get(edge.target), edge.target);
  const label = typeof edge.label === 'string' && edge.label.trim()
    ? ` · ${edge.label.trim()}`
    : '';
  return `${source} → ${target}${label}`;
}

function clone<T>(value: T): T {
  return JSON.parse(JSON.stringify(value)) as T;
}

export function compareDiagramVersions(
  currentNodes: Node[],
  currentEdges: Edge[],
  targetNodes: Node[],
  targetEdges: Edge[],
): VersionDiff {
  const currentNodeMap = new Map(currentNodes.map(node => [node.id, node]));
  const targetNodeMap = new Map(targetNodes.map(node => [node.id, node]));
  const currentEdgeMap = new Map(currentEdges.map(edge => [edge.id, edge]));
  const targetEdgeMap = new Map(targetEdges.map(edge => [edge.id, edge]));
  const allNodeIds = [...new Set([...currentNodeMap.keys(), ...targetNodeMap.keys()])].sort();
  const allEdgeIds = [...new Set([...currentEdgeMap.keys(), ...targetEdgeMap.keys()])].sort();

  const nodeItems = allNodeIds.flatMap<VersionDiffItem>((id) => {
    const before = currentNodeMap.get(id);
    const after = targetNodeMap.get(id);
    if (!before && after) {
      return [{
        key: `node:${id}`,
        kind: 'node',
        status: 'added',
        id,
        label: nodeLabel(after, id),
        after,
      }];
    }
    if (before && !after) {
      return [{
        key: `node:${id}`,
        kind: 'node',
        status: 'removed',
        id,
        label: nodeLabel(before, id),
        before,
      }];
    }
    if (before && after && fingerprint(before) !== fingerprint(after)) {
      return [{
        key: `node:${id}`,
        kind: 'node',
        status: 'changed',
        id,
        label: nodeLabel(after, id),
        before,
        after,
      }];
    }
    return [];
  });

  const combinedNodeMap = new Map([...currentNodeMap, ...targetNodeMap]);
  const edgeItems = allEdgeIds.flatMap<VersionDiffItem>((id) => {
    const before = currentEdgeMap.get(id);
    const after = targetEdgeMap.get(id);
    if (!before && after) {
      return [{
        key: `edge:${id}`,
        kind: 'edge',
        status: 'added',
        id,
        label: edgeLabel(after, combinedNodeMap, id),
        after,
      }];
    }
    if (before && !after) {
      return [{
        key: `edge:${id}`,
        kind: 'edge',
        status: 'removed',
        id,
        label: edgeLabel(before, combinedNodeMap, id),
        before,
      }];
    }
    if (before && after && fingerprint(before) !== fingerprint(after)) {
      return [{
        key: `edge:${id}`,
        kind: 'edge',
        status: 'changed',
        id,
        label: edgeLabel(after, combinedNodeMap, id),
        before,
        after,
      }];
    }
    return [];
  });

  const items = [...nodeItems, ...edgeItems];
  return {
    items,
    nodes: nodeItems,
    edges: edgeItems,
    counts: {
      added: items.filter(item => item.status === 'added').length,
      removed: items.filter(item => item.status === 'removed').length,
      changed: items.filter(item => item.status === 'changed').length,
    },
  };
}

export function applySelectedVersionChanges(
  currentNodes: Node[],
  currentEdges: Edge[],
  targetNodes: Node[],
  targetEdges: Edge[],
  selectedKeys: Iterable<string>,
): SelectiveVersionRestoreResult {
  const selected = new Set(selectedKeys);
  const diff = compareDiagramVersions(currentNodes, currentEdges, targetNodes, targetEdges);
  const selectedItems = diff.items.filter(item => selected.has(item.key));
  const appliedKeys: string[] = [];
  const skippedKeys: string[] = [];
  const autoAddedNodeIds: string[] = [];
  const targetNodeMap = new Map(targetNodes.map(node => [node.id, node]));
  const targetEdgeMap = new Map(targetEdges.map(edge => [edge.id, edge]));
  const selectedNodeKeys = new Set(
    selectedItems.filter(item => item.kind === 'node').map(item => item.key),
  );

  const removedNodeIds = new Set(
    selectedItems
      .filter(item => item.kind === 'node' && item.status === 'removed')
      .map(item => item.id),
  );
  const baseNodes = deleteNodesPreservingGroupChildren(clone(currentNodes), removedNodeIds);
  const nodeOrder = baseNodes.map(node => node.id);
  const nodeMap = new Map(baseNodes.map(node => [node.id, node]));
  const visiting = new Set<string>();

  const ensureTargetNode = (nodeId: string, requestedKey?: string): boolean => {
    const target = targetNodeMap.get(nodeId);
    if (!target || visiting.has(nodeId)) {
      if (requestedKey) skippedKeys.push(requestedKey);
      return false;
    }
    visiting.add(nodeId);
    if (
      target.parentNode
      && !nodeMap.has(target.parentNode)
      && !ensureTargetNode(target.parentNode)
    ) {
      visiting.delete(nodeId);
      if (requestedKey) skippedKeys.push(requestedKey);
      return false;
    }
    const alreadyPresent = nodeMap.has(nodeId);
    nodeMap.set(nodeId, {
      ...clone(target),
      selected: false,
      dragging: false,
      positionAbsolute: undefined,
    });
    if (!alreadyPresent) nodeOrder.push(nodeId);
    if (!requestedKey && !selectedNodeKeys.has(`node:${nodeId}`)) {
      autoAddedNodeIds.push(nodeId);
    }
    visiting.delete(nodeId);
    if (requestedKey) appliedKeys.push(requestedKey);
    return true;
  };

  for (const item of selectedItems) {
    if (item.kind !== 'node') continue;
    if (item.status === 'removed') {
      appliedKeys.push(item.key);
      continue;
    }
    ensureTargetNode(item.id, item.key);
  }

  const nextNodes = nodeOrder
    .map(id => nodeMap.get(id))
    .filter((node): node is Node => Boolean(node));
  const retainedNodeIds = new Set(nextNodes.map(node => node.id));
  const autoRemovedEdgeIds: string[] = [];
  const edgeOrder: string[] = [];
  const edgeMap = new Map<string, Edge>();

  for (const edge of clone(currentEdges)) {
    if (!retainedNodeIds.has(edge.source) || !retainedNodeIds.has(edge.target)) {
      autoRemovedEdgeIds.push(edge.id);
      continue;
    }
    edgeOrder.push(edge.id);
    edgeMap.set(edge.id, edge);
  }

  for (const item of selectedItems) {
    if (item.kind !== 'edge') continue;
    if (item.status === 'removed') {
      edgeMap.delete(item.id);
      appliedKeys.push(item.key);
      continue;
    }
    const target = targetEdgeMap.get(item.id);
    if (
      !target
      || !retainedNodeIds.has(target.source)
      || !retainedNodeIds.has(target.target)
    ) {
      skippedKeys.push(item.key);
      continue;
    }
    if (!edgeMap.has(item.id)) edgeOrder.push(item.id);
    edgeMap.set(item.id, {
      ...clone(target),
      selected: false,
    });
    appliedKeys.push(item.key);
  }

  return {
    nodes: nextNodes,
    edges: edgeOrder
      .map(id => edgeMap.get(id))
      .filter((edge): edge is Edge => Boolean(edge)),
    appliedKeys: [...new Set(appliedKeys)],
    skippedKeys: [...new Set(skippedKeys)],
    autoAddedNodeIds: [...new Set(autoAddedNodeIds)],
    autoRemovedEdgeIds: [...new Set(autoRemovedEdgeIds)],
  };
}
