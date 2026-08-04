// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

import type { Node, XYPosition } from 'reactflow';

export const SERVICE_NODE_ESTIMATED_SIZE = {
  width: 180,
  height: 132,
} as const;

interface PlacementOptions {
  gap?: number;
  gridSize?: number;
  width?: number;
  height?: number;
  maxRings?: number;
}

function dimension(value: unknown, fallback: number): number {
  const numeric = Number(value);
  return Number.isFinite(numeric) && numeric > 0 ? numeric : fallback;
}

function snap(value: number, gridSize: number): number {
  return Math.round(value / gridSize) * gridSize;
}

function overlaps(
  candidate: { x: number; y: number; width: number; height: number },
  existing: { x: number; y: number; width: number; height: number },
  gap: number,
): boolean {
  return candidate.x < existing.x + existing.width + gap
    && candidate.x + candidate.width + gap > existing.x
    && candidate.y < existing.y + existing.height + gap
    && candidate.y + candidate.height + gap > existing.y;
}

function ringOffsets(ring: number): XYPosition[] {
  if (ring === 0) return [{ x: 0, y: 0 }];
  const offsets: XYPosition[] = [];
  for (let x = -ring; x <= ring; x += 1) offsets.push({ x, y: -ring });
  for (let y = -ring + 1; y <= ring; y += 1) offsets.push({ x: ring, y });
  for (let x = ring - 1; x >= -ring; x -= 1) offsets.push({ x, y: ring });
  for (let y = ring - 1; y > -ring; y -= 1) offsets.push({ x: -ring, y });
  return offsets;
}

export function findAvailableServicePosition(
  desired: XYPosition,
  existingNodes: Node[],
  options: PlacementOptions = {},
): XYPosition {
  const width = options.width ?? SERVICE_NODE_ESTIMATED_SIZE.width;
  const height = options.height ?? SERVICE_NODE_ESTIMATED_SIZE.height;
  const gap = options.gap ?? 28;
  const gridSize = options.gridSize ?? 20;
  const maxRings = options.maxRings ?? 12;
  const stepX = snap(width + gap, gridSize);
  const stepY = snap(height + gap, gridSize);
  const occupied = existingNodes.map((node) => ({
    x: node.position.x,
    y: node.position.y,
    width: dimension(node.width ?? node.style?.width, width),
    height: dimension(node.height ?? node.style?.height, height),
  }));

  for (let ring = 0; ring <= maxRings; ring += 1) {
    for (const offset of ringOffsets(ring)) {
      const position = {
        x: snap(desired.x + (offset.x * stepX), gridSize),
        y: snap(desired.y + (offset.y * stepY), gridSize),
      };
      const candidate = { ...position, width, height };
      if (!occupied.some((node) => overlaps(candidate, node, gap))) return position;
    }
  }

  return {
    x: snap(desired.x + ((maxRings + 1) * stepX), gridSize),
    y: snap(desired.y, gridSize),
  };
}
