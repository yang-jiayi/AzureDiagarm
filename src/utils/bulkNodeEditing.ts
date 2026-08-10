// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

import type { Node } from 'reactflow';
import type { NodePricingConfig } from '../types/pricing';
import { fitGroupToContent } from './groupUtils';
import { buildAbsolutePositionMap } from './preserveManualLayout';

export type BulkAlignmentType =
  | 'left'
  | 'center-h'
  | 'right'
  | 'top'
  | 'center-v'
  | 'bottom'
  | 'distribute-h'
  | 'distribute-v';

export type BulkTagMode = 'add' | 'replace' | 'remove';

export interface BulkGroupColor {
  name: string;
  bg: string;
  border: string;
  header: string;
}

export const BULK_GROUP_COLORS: BulkGroupColor[] = [
  { name: 'Gray', bg: 'rgba(107, 114, 128, 0.10)', border: '#6b7280', header: '#6b7280' },
  { name: 'Blue', bg: 'rgba(0, 120, 212, 0.10)', border: '#0078d4', header: '#0078d4' },
  { name: 'Green', bg: 'rgba(16, 185, 129, 0.10)', border: '#10b981', header: '#10b981' },
  { name: 'Orange', bg: 'rgba(245, 158, 11, 0.10)', border: '#f59e0b', header: '#f59e0b' },
  { name: 'Red', bg: 'rgba(239, 68, 68, 0.10)', border: '#ef4444', header: '#ef4444' },
  { name: 'Purple', bg: 'rgba(139, 92, 246, 0.10)', border: '#8b5cf6', header: '#8b5cf6' },
  { name: 'Cyan', bg: 'rgba(6, 182, 212, 0.10)', border: '#06b6d4', header: '#06b6d4' },
  { name: 'Pink', bg: 'rgba(236, 72, 153, 0.10)', border: '#ec4899', header: '#ec4899' },
  { name: 'Teal', bg: 'rgba(20, 184, 166, 0.10)', border: '#14b8a6', header: '#14b8a6' },
];

export interface BulkNodeEdits {
  targetGroupId?: string | null;
  stylePreset?: 'detailed' | 'presentation';
  groupColor?: BulkGroupColor | null;
  tags?: {
    mode: BulkTagMode;
    values: string[];
  };
  pricingByNodeId?: ReadonlyMap<string, NodePricingConfig>;
}

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
      ?? (node.type === 'groupNode' ? 300 : 128),
  };
}

function hasSelectedAncestor(
  node: Node,
  selectedIds: ReadonlySet<string>,
  nodesById: ReadonlyMap<string, Node>,
): boolean {
  let parentId = node.parentNode;
  const visited = new Set<string>();
  while (parentId && !visited.has(parentId)) {
    if (selectedIds.has(parentId)) return true;
    visited.add(parentId);
    parentId = nodesById.get(parentId)?.parentNode;
  }
  return false;
}

export function alignSelectedNodes(
  nodes: Node[],
  type: BulkAlignmentType,
): Node[] {
  const allSelectedIds = new Set(nodes.filter(node => node.selected).map(node => node.id));
  if (allSelectedIds.size < 2) return nodes;
  const nodesById = new Map(nodes.map(node => [node.id, node]));
  const selected = nodes.filter(node => (
    allSelectedIds.has(node.id)
    && !hasSelectedAncestor(node, allSelectedIds, nodesById)
  ));
  if (selected.length < 2) return nodes;

  const absolute = buildAbsolutePositionMap(nodes);
  const rectangles = selected.map(node => {
    const position = absolute.get(node.id) || node.position;
    const size = nodeDimensions(node);
    return {
      node,
      x: position.x,
      y: position.y,
      width: size.width,
      height: size.height,
    };
  });
  const targetPositions = new Map<string, { x: number; y: number }>();

  if (type === 'left' || type === 'center-h' || type === 'right') {
    const left = Math.min(...rectangles.map(rect => rect.x));
    const right = Math.max(...rectangles.map(rect => rect.x + rect.width));
    const centerX = (left + right) / 2;
    rectangles.forEach(rect => {
      const x = type === 'left'
        ? left
        : type === 'right'
          ? right - rect.width
          : centerX - rect.width / 2;
      targetPositions.set(rect.node.id, { x, y: rect.y });
    });
  } else if (type === 'top' || type === 'center-v' || type === 'bottom') {
    const top = Math.min(...rectangles.map(rect => rect.y));
    const bottom = Math.max(...rectangles.map(rect => rect.y + rect.height));
    const centerY = (top + bottom) / 2;
    rectangles.forEach(rect => {
      const y = type === 'top'
        ? top
        : type === 'bottom'
          ? bottom - rect.height
          : centerY - rect.height / 2;
      targetPositions.set(rect.node.id, { x: rect.x, y });
    });
  } else {
    const horizontal = type === 'distribute-h';
    const sorted = [...rectangles].sort((left, right) => (
      horizontal ? left.x - right.x : left.y - right.y
    ));
    const start = horizontal ? sorted[0].x : sorted[0].y;
    const end = horizontal
      ? sorted[sorted.length - 1].x
      : sorted[sorted.length - 1].y;
    const spacing = (end - start) / (sorted.length - 1);
    sorted.forEach((rect, index) => {
      targetPositions.set(rect.node.id, {
        x: horizontal ? start + spacing * index : rect.x,
        y: horizontal ? rect.y : start + spacing * index,
      });
    });
  }

  return nodes.map(node => {
    const target = targetPositions.get(node.id);
    if (!target) return node;
    const parentPosition = node.parentNode
      ? absolute.get(node.parentNode) || { x: 0, y: 0 }
      : { x: 0, y: 0 };
    return {
      ...node,
      position: {
        x: target.x - parentPosition.x,
        y: target.y - parentPosition.y,
      },
      positionAbsolute: undefined,
    };
  });
}

export function normalizeBulkTags(values: Iterable<string>): string[] {
  const normalized: string[] = [];
  const seen = new Set<string>();
  for (const value of values) {
    const tag = String(value).trim().replace(/\s+/g, ' ').slice(0, 40);
    const key = tag.toLowerCase();
    if (!tag || seen.has(key)) continue;
    seen.add(key);
    normalized.push(tag);
    if (normalized.length >= 12) break;
  }
  return normalized;
}

function applyTags(
  currentValue: unknown,
  mode: BulkTagMode,
  values: string[],
): string[] {
  const current = normalizeBulkTags(Array.isArray(currentValue) ? currentValue : []);
  if (mode === 'replace') return values;
  const targetKeys = new Set(values.map(value => value.toLowerCase()));
  if (mode === 'remove') {
    return current.filter(value => !targetKeys.has(value.toLowerCase()));
  }
  return normalizeBulkTags([...current, ...values]);
}

export function applyBulkNodeEdits(
  nodes: Node[],
  selectedNodeIds: ReadonlySet<string>,
  edits: BulkNodeEdits,
): Node[] {
  if (selectedNodeIds.size === 0) return nodes;
  const absolute = buildAbsolutePositionMap(nodes);
  const targetGroup = edits.targetGroupId
    ? nodes.find(node => node.id === edits.targetGroupId && node.type === 'groupNode')
    : undefined;
  const targetGroupPosition = targetGroup
    ? absolute.get(targetGroup.id) || targetGroup.position
    : undefined;
  const normalizedTags = edits.tags
    ? normalizeBulkTags(edits.tags.values)
    : [];
  const affectedGroupIds = new Set<string>();

  let updated = nodes.map(node => {
    if (!selectedNodeIds.has(node.id)) return node;
    const data = { ...node.data };
    let position = node.position;
    let parentNode = node.parentNode;
    let extent = node.extent;

    if (edits.tags) {
      data.tags = applyTags(data.tags, edits.tags.mode, normalizedTags);
    }
    if (node.type === 'azureNode') {
      if (edits.stylePreset) data.stylePreset = edits.stylePreset;
      const pricing = edits.pricingByNodeId?.get(node.id);
      if (pricing) data.pricing = pricing;

      if (edits.targetGroupId !== undefined) {
        if (node.parentNode) affectedGroupIds.add(node.parentNode);
        const nodeAbsolute = absolute.get(node.id) || node.position;
        if (edits.targetGroupId === null) {
          position = nodeAbsolute;
          parentNode = undefined;
          extent = undefined;
          delete data.groupId;
        } else if (targetGroup && targetGroupPosition) {
          position = {
            x: nodeAbsolute.x - targetGroupPosition.x,
            y: nodeAbsolute.y - targetGroupPosition.y,
          };
          parentNode = targetGroup.id;
          extent = 'parent';
          data.groupId = targetGroup.id;
          affectedGroupIds.add(targetGroup.id);
        }
      }
    } else if (node.type === 'groupNode' && edits.groupColor !== undefined) {
      if (edits.groupColor === null) {
        delete data.customColor;
      } else {
        data.customColor = { ...edits.groupColor };
      }
    }

    return {
      ...node,
      position,
      parentNode,
      extent,
      data,
    };
  });

  for (const groupId of affectedGroupIds) {
    const fitted = fitGroupToContent(updated, groupId);
    if (fitted) updated = fitted;
  }
  return updated;
}
