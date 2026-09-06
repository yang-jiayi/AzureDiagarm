// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

import type { LayoutResult, PositionedEdge, PositionedNode } from './layoutEngine.js';

type Point = { x: number; y: number };
type Rect = Pick<PositionedNode, 'x' | 'y' | 'width' | 'height'>;

export function connectionKey(edge: Pick<PositionedEdge, 'key' | 'from' | 'to'>): string {
  return edge.key ?? `${edge.from}\u0000${edge.to}`;
}

function pairKey(from: string, to: string): string {
  return JSON.stringify(from < to ? [from, to] : [to, from]);
}

export function selfLoopPoints(node: Rect, ordinal = 0): Point[] {
  const right = node.x + node.width;
  const outer = right + Math.min(node.width, node.height) / 2 + 14 + ordinal * 16;
  const top = node.y + node.height * 0.3;
  const bottom = node.y + node.height * 0.7;
  return [
    { x: right, y: top }, { x: outer, y: top },
    { x: outer, y: bottom }, { x: right, y: bottom },
  ];
}

function parallelPoints(source: Rect, target: Rect, points: Point[], ordinal: number, count: number): Point[] {
  const start = points[0];
  const end = points[points.length - 1];
  if (!start || !end) return points;
  if (Math.abs(start.x - end.x) < 0.01 || Math.abs(start.y - end.y) < 0.01) return [start, end];
  const horizontal = Math.abs(target.x + target.width / 2 - source.x - source.width / 2)
    >= Math.abs(target.y + target.height / 2 - source.y - source.height / 2);
  const span = horizontal ? end.x - start.x : end.y - start.y;
  const step = Math.min(12, Math.abs(span) / (2 * Math.max(count, 1)));
  const offset = (ordinal - (count - 1) / 2) * step;
  const middle = horizontal ? (start.x + end.x) / 2 + offset : (start.y + end.y) / 2 + offset;
  return horizontal
    ? [start, { x: middle, y: start.y }, { x: middle, y: end.y }, end]
    : [start, { x: start.x, y: middle }, { x: end.x, y: middle }, end];
}

/** Only exceptional routes are fixed here; ordinary obstacle-aware routing is unchanged. */
export function preserveConnectionRoutes(
  nodes: PositionedNode[], edges: PositionedEdge[], reanchored = false,
): PositionedEdge[] {
  const byName = new Map(nodes.map(node => [node.name, node]));
  const counts = new Map<string, number>();
  const ordinals = new Map<string, number>();
  for (const edge of edges) {
    const pair = pairKey(edge.from, edge.to);
    counts.set(pair, (counts.get(pair) ?? 0) + 1);
  }
  return edges.map(edge => {
    const source = byName.get(edge.from);
    const target = byName.get(edge.to);
    if (!source || !target) return edge;
    const pair = pairKey(edge.from, edge.to);
    const ordinal = ordinals.get(pair) ?? 0;
    ordinals.set(pair, ordinal + 1);
    if (edge.from === edge.to) {
      const hasExcursion = edge.points.length > 2
        && edge.points.every(point => Number.isFinite(point.x) && Number.isFinite(point.y))
        && edge.points.some(point => point.x < source.x || point.x > source.x + source.width
          || point.y < source.y || point.y > source.y + source.height);
      return { ...edge, routeKind: 'self-loop', points: !reanchored && hasExcursion ? edge.points : selfLoopPoints(source, ordinal) };
    }
    const count = counts.get(pair)!;
    if (count > 1) {
      return {
        ...edge,
        routeKind: 'parallel',
        points: reanchored || edge.points.length < 3
          ? parallelPoints(source, target, edge.points, ordinal, count)
          : edge.points,
      };
    }
    return edge.routeKind ? { ...edge, routeKind: undefined } : edge;
  });
}

/** Include route excursions after both initial layout and presentation reanchoring. */
export function encloseConnectionRoutes(layout: LayoutResult): LayoutResult {
  let minX = 0, minY = 0, maxX = 0, maxY = 0;
  for (const edge of layout.edges) {
    for (const point of edge.points) {
      if (!Number.isFinite(point.x) || !Number.isFinite(point.y)) continue;
      minX = Math.min(minX, point.x);
      minY = Math.min(minY, point.y);
      maxX = Math.max(maxX, point.x);
      maxY = Math.max(maxY, point.y);
    }
  }
  const dx = minX < 0 ? 40 - minX : 0;
  const dy = minY < 0 ? 40 - minY : 0;
  const width = Math.max(layout.width, maxX + 40) + dx;
  const height = Math.max(layout.height, maxY + 40) + dy;
  if (width === layout.width && height === layout.height) return layout;
  if (dx === 0 && dy === 0) return { ...layout, width, height };
  return {
    ...layout, width, height,
    nodes: layout.nodes.map(node => ({ ...node, x: node.x + dx, y: node.y + dy })),
    groups: layout.groups.map(group => ({ ...group, x: group.x + dx, y: group.y + dy })),
    edges: layout.edges.map(edge => ({ ...edge, points: edge.points.map(point => ({ x: point.x + dx, y: point.y + dy })) })),
  };
}
