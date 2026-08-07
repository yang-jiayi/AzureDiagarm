// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

import type { Edge, Node, NodeChange } from 'reactflow';
import { layoutNodeDimensions } from './layoutHierarchy';
import { buildAbsolutePositionMap } from './preserveManualLayout';

interface Rect {
  x: number;
  y: number;
  width: number;
  height: number;
}

interface LabelOffset {
  x: number;
  y: number;
}

function overlapArea(left: Rect, right: Rect): number {
  const width = Math.min(left.x + left.width, right.x + right.width) - Math.max(left.x, right.x);
  const height = Math.min(left.y + left.height, right.y + right.height) - Math.max(left.y, right.y);
  return Math.max(0, width) * Math.max(0, height);
}

function estimateLabelSize(label: string): { width: number; height: number } {
  const normalizedLength = Math.max(1, label.trim().length);
  const width = Math.min(220, Math.max(76, normalizedLength * 7.2 + 22));
  const estimatedLines = Math.min(3, Math.max(1, Math.ceil((normalizedLength * 7.2) / 190)));
  return { width, height: 14 + estimatedLines * 17 };
}

function nodeRects(nodes: Node[]): Map<string, Rect> {
  const absolute = buildAbsolutePositionMap(nodes);
  return new Map(
    nodes
      .filter(node => node.type === 'azureNode')
      .flatMap((node) => {
        const position = absolute.get(node.id);
        if (!position) return [];
        const dimensions = layoutNodeDimensions(node);
        return [[node.id, {
          x: position.x - 10,
          y: position.y - 10,
          width: dimensions.width + 20,
          height: dimensions.height + 20,
        }] as const];
      }),
  );
}

export function calculateAutomaticEdgeLabelOffsets(
  nodes: Node[],
  edges: Edge[],
): Map<string, LabelOffset> {
  const blockersById = nodeRects(nodes);
  const blockers = [...blockersById.values()];
  const placedLabels: Rect[] = [];
  const offsets = new Map<string, LabelOffset>();

  edges.forEach((edge, edgeIndex) => {
    const label = String(edge.label ?? '').trim();
    if (!label) return;

    const source = blockersById.get(edge.source);
    const target = blockersById.get(edge.target);
    if (!source || !target) return;

    const sourceCenter = {
      x: source.x + source.width / 2,
      y: source.y + source.height / 2,
    };
    const targetCenter = {
      x: target.x + target.width / 2,
      y: target.y + target.height / 2,
    };
    const midpoint = {
      x: (sourceCenter.x + targetCenter.x) / 2,
      y: (sourceCenter.y + targetCenter.y) / 2,
    };
    const dx = targetCenter.x - sourceCenter.x;
    const dy = targetCenter.y - sourceCenter.y;
    const length = Math.hypot(dx, dy) || 1;
    const normal = { x: -dy / length, y: dx / length };
    const tangent = { x: dx / length, y: dy / length };
    const side = edgeIndex % 2 === 0 ? 1 : -1;
    const distances = [0, 30 * side, -30 * side, 58 * side, -58 * side, 88 * side, -88 * side];
    const candidates = distances.map(distance => ({
      x: normal.x * distance,
      y: normal.y * distance,
    }));
    candidates.push(
      { x: normal.x * 46 + tangent.x * 38, y: normal.y * 46 + tangent.y * 38 },
      { x: normal.x * -46 + tangent.x * -38, y: normal.y * -46 + tangent.y * -38 },
    );

    const size = estimateLabelSize(label);
    let best = candidates[0];
    let bestRect: Rect = {
      x: midpoint.x - size.width / 2,
      y: midpoint.y - size.height / 2,
      ...size,
    };
    let bestScore = Number.POSITIVE_INFINITY;

    for (const candidate of candidates) {
      const rect: Rect = {
        x: midpoint.x + candidate.x - size.width / 2,
        y: midpoint.y + candidate.y - size.height / 2,
        ...size,
      };
      const blockerPenalty = blockers.reduce(
        (total, blocker) => total + overlapArea(rect, blocker),
        0,
      );
      const labelPenalty = placedLabels.reduce(
        (total, placed) => total + overlapArea(rect, placed),
        0,
      );
      const distancePenalty = Math.hypot(candidate.x, candidate.y) * 0.35;
      const score = blockerPenalty * 5_000 + labelPenalty * 8_000 + distancePenalty;
      if (score < bestScore) {
        best = candidate;
        bestRect = rect;
        bestScore = score;
      }
    }

    offsets.set(edge.id, best);
    placedLabels.push(bestRect);
  });

  return offsets;
}

export function shouldRecalculateAutomaticEdgeLabels(changes: NodeChange[]): boolean {
  return changes.some((change) => (
    (change.type === 'position' && change.dragging !== true)
    || (change.type === 'dimensions' && change.resizing !== true)
  ));
}

export function applyAutomaticEdgeLabelOffsets(nodes: Node[], edges: Edge[]): Edge[] {
  const offsets = calculateAutomaticEdgeLabelOffsets(nodes, edges);
  return edges.map((edge) => {
    const data = (edge.data ?? {}) as Record<string, unknown>;
    if (data.labelOffsetAuto === false) return edge;
    const offset = offsets.get(edge.id);
    if (!offset) return edge;
    if (
      data.labelOffsetAuto === true
      && data.labelOffsetX === offset.x
      && data.labelOffsetY === offset.y
    ) {
      return edge;
    }
    return {
      ...edge,
      data: {
        ...data,
        labelOffsetX: offset.x,
        labelOffsetY: offset.y,
        labelOffsetAuto: true,
      },
    };
  });
}
