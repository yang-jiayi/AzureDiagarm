// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

/**
 * Shared, pure geometry helpers for the "native shape" exporters
 * (PowerPoint / Visio).
 *
 * The canvas stores React Flow coordinates in pixels with a top-left origin and
 * child nodes positioned relative to their parent group. Office formats need
 * absolute coordinates in inches, so every exporter needs the same three steps:
 *
 *   1. flatten node positions to an absolute pixel box,
 *   2. fit the drawing into the target frame (slide body / Visio page),
 *   3. route each edge as an orthogonal polyline between box anchors.
 *
 * Keeping this DOM-free makes it unit-testable and keeps the two exporters
 * visually consistent.
 */

import type { Edge, Node } from 'reactflow';

export const DEFAULT_SERVICE_W = 150;
export const DEFAULT_SERVICE_H = 75;
export const DEFAULT_GROUP_W = 400;
export const DEFAULT_GROUP_H = 300;

export interface ExportBox {
  id: string;
  kind: 'group' | 'service';
  label: string;
  iconPath?: string;
  /**
   * Azure icon-folder category (lower-cased, e.g. `databases`). Drives the
   * colour coding in every export. Resolved from the node's own `category`
   * first because AI-generated services keep their category even when no icon
   * file matched and `iconPath` is empty.
   */
  category: string;
  /** Canonical service name when it differs from the user-editable label. */
  serviceName?: string;
  /** Absolute pixel geometry (top-left origin). */
  x: number;
  y: number;
  w: number;
  h: number;
}

export interface Point {
  x: number;
  y: number;
}

export interface ExportRoute {
  id: string;
  sourceId: string;
  targetId: string;
  label: string;
  dashed: boolean;
  /** Orthogonal polyline in absolute pixels; first point starts at the source. */
  points: Point[];
  /** Preferred label anchor (centre of the middle segment). */
  labelAnchor: Point;
}

export interface Bounds {
  minX: number;
  minY: number;
  maxX: number;
  maxY: number;
}

export interface FitTransform {
  /** Multiply a pixel value by this to get inches. */
  scale: number;
  /** Inches added after scaling (already includes the frame origin). */
  offsetX: number;
  offsetY: number;
}

function readSize(node: Node): { w: number; h: number } {
  const anyNode = node as any;
  const styleW = typeof anyNode.style?.width === 'number' ? anyNode.style.width : undefined;
  const styleH = typeof anyNode.style?.height === 'number' ? anyNode.style.height : undefined;
  if (node.type === 'groupNode') {
    return {
      w: styleW ?? anyNode.width ?? DEFAULT_GROUP_W,
      h: styleH ?? anyNode.height ?? DEFAULT_GROUP_H,
    };
  }
  return {
    w: anyNode.width ?? styleW ?? DEFAULT_SERVICE_W,
    h: anyNode.height ?? styleH ?? DEFAULT_SERVICE_H,
  };
}

/**
 * Resolve every node's absolute canvas position by walking the full
 * `parentNode` chain. Groups can be nested (a zone inside a zone), and React
 * Flow stores each child relative to its immediate parent, so a single-level
 * lookup would drop nested content near the canvas origin. The `visiting` set
 * makes a corrupt cyclic parent chain fall back to the relative position
 * instead of recursing forever.
 */
function resolveAbsolutePositions(nodes: Node[]): Map<string, Point> {
  const nodesById = new Map(nodes.map((node) => [node.id, node]));
  const resolved = new Map<string, Point>();

  const resolve = (node: Node, visiting: Set<string>): Point => {
    const cached = resolved.get(node.id);
    if (cached) return cached;

    const own: Point = { x: node.position?.x ?? 0, y: node.position?.y ?? 0 };
    if (visiting.has(node.id)) return own;

    const parentId = (node as { parentNode?: string }).parentNode;
    const parent = parentId ? nodesById.get(parentId) : undefined;
    let position = own;
    if (parent) {
      const nextVisiting = new Set(visiting);
      nextVisiting.add(node.id);
      const parentPosition = resolve(parent, nextVisiting);
      position = { x: parentPosition.x + own.x, y: parentPosition.y + own.y };
    }
    resolved.set(node.id, position);
    return position;
  };

  for (const node of nodes) resolve(node, new Set());
  return resolved;
}

/**
 * Normalize an Azure icon-folder name into a category key.
 * Icon paths look like `/Azure_Public_Service_Icons/Icons/<folder>/<file>.svg`.
 */
export function categoryFromIconPath(iconPath?: string): string | undefined {
  const match = iconPath?.match(/\/Icons\/([^/]+)\//i);
  return match ? match[1].replace(/[-_]+/g, ' ').trim().toLowerCase() : undefined;
}

/**
 * Resolve the category used for export colour coding. The node's own category
 * wins: AI-generated nodes are stored with `iconPath: icon?.path || ''`, so a
 * service with no matching icon file would otherwise lose its colour and be
 * rendered as an unstyled grey tile.
 */
function resolveCategory(data: Record<string, unknown>, iconPath?: string): string {
  const explicit = typeof data.category === 'string' ? data.category.replace(/[-_]+/g, ' ').trim().toLowerCase() : '';
  return explicit || categoryFromIconPath(iconPath) || 'other';
}

/**
 * Flatten React Flow nodes into absolute pixel boxes. Children — including
 * nested groups — are offset by their whole ancestor chain.
 */
export function collectExportBoxes(nodes: Node[]): Map<string, ExportBox> {
  const positions = resolveAbsolutePositions(nodes);

  const boxes = new Map<string, ExportBox>();
  for (const node of nodes) {
    const { w, h } = readSize(node);
    const position = positions.get(node.id) ?? { x: node.position?.x ?? 0, y: node.position?.y ?? 0 };
    const data = (node.data ?? {}) as Record<string, unknown>;
    const iconPath = typeof data.iconPath === 'string' && data.iconPath ? data.iconPath : undefined;
    boxes.set(node.id, {
      id: node.id,
      kind: node.type === 'groupNode' ? 'group' : 'service',
      label: typeof data.label === 'string' && data.label.trim()
        ? data.label
        : node.type === 'groupNode' ? 'Group' : 'Service',
      iconPath,
      category: resolveCategory(data, iconPath),
      serviceName: typeof data.serviceName === 'string' && data.serviceName.trim()
        ? data.serviceName
        : undefined,
      x: position.x,
      y: position.y,
      w: Math.max(1, w),
      h: Math.max(1, h),
    });
  }
  return boxes;
}

export function computeBounds(boxes: Iterable<ExportBox>): Bounds {
  let minX = Number.POSITIVE_INFINITY;
  let minY = Number.POSITIVE_INFINITY;
  let maxX = Number.NEGATIVE_INFINITY;
  let maxY = Number.NEGATIVE_INFINITY;
  let seen = false;
  for (const box of boxes) {
    seen = true;
    minX = Math.min(minX, box.x);
    minY = Math.min(minY, box.y);
    maxX = Math.max(maxX, box.x + box.w);
    maxY = Math.max(maxY, box.y + box.h);
  }
  if (!seen) return { minX: 0, minY: 0, maxX: 1, maxY: 1 };
  return { minX, minY, maxX, maxY };
}

/**
 * Fit `bounds` (pixels) inside `frame` (inches) preserving the aspect ratio and
 * centring the result. `maxScale` prevents a two-node diagram from being blown
 * up to absurd tile sizes.
 */
export function computeFitTransform(
  bounds: Bounds,
  frame: { x: number; y: number; w: number; h: number },
  options: { maxScale?: number } = {},
): FitTransform {
  const contentW = Math.max(1, bounds.maxX - bounds.minX);
  const contentH = Math.max(1, bounds.maxY - bounds.minY);
  const maxScale = options.maxScale ?? 1 / 96; // never larger than 96 px per inch
  const scale = Math.min(frame.w / contentW, frame.h / contentH, maxScale);
  const offsetX = frame.x + (frame.w - contentW * scale) / 2 - bounds.minX * scale;
  const offsetY = frame.y + (frame.h - contentH * scale) / 2 - bounds.minY * scale;
  return { scale, offsetX, offsetY };
}

export function applyTransform(point: Point, transform: FitTransform): Point {
  return {
    x: point.x * transform.scale + transform.offsetX,
    y: point.y * transform.scale + transform.offsetY,
  };
}

function centre(box: ExportBox): Point {
  return { x: box.x + box.w / 2, y: box.y + box.h / 2 };
}

/**
 * Right-angle "Z" route between two boxes. The exit/entry anchors sit on the
 * facing edges so the arrow never starts inside a tile, and the elbow is placed
 * on the mid-line so labels stay clear of both shapes.
 */
export function routeOrthogonal(source: ExportBox, target: ExportBox): {
  points: Point[];
  labelAnchor: Point;
} {
  const sc = centre(source);
  const tc = centre(target);
  const dx = tc.x - sc.x;
  const dy = tc.y - sc.y;

  if (Math.abs(dx) >= Math.abs(dy)) {
    const startX = dx >= 0 ? source.x + source.w : source.x;
    const endX = dx >= 0 ? target.x : target.x + target.w;
    const startY = sc.y;
    const endY = tc.y;
    if (Math.abs(startY - endY) < 0.5) {
      return {
        points: [{ x: startX, y: startY }, { x: endX, y: endY }],
        labelAnchor: { x: (startX + endX) / 2, y: startY },
      };
    }
    const midX = (startX + endX) / 2;
    return {
      points: [
        { x: startX, y: startY },
        { x: midX, y: startY },
        { x: midX, y: endY },
        { x: endX, y: endY },
      ],
      labelAnchor: { x: midX, y: (startY + endY) / 2 },
    };
  }

  const startY = dy >= 0 ? source.y + source.h : source.y;
  const endY = dy >= 0 ? target.y : target.y + target.h;
  const startX = sc.x;
  const endX = tc.x;
  if (Math.abs(startX - endX) < 0.5) {
    return {
      points: [{ x: startX, y: startY }, { x: endX, y: endY }],
      labelAnchor: { x: startX, y: (startY + endY) / 2 },
    };
  }
  const midY = (startY + endY) / 2;
  return {
    points: [
      { x: startX, y: startY },
      { x: startX, y: midY },
      { x: endX, y: midY },
      { x: endX, y: endY },
    ],
    labelAnchor: { x: (startX + endX) / 2, y: midY },
  };
}

function readEdgeLabel(edge: Edge): string {
  const dataLabel = (edge.data as { label?: unknown } | undefined)?.label;
  if (typeof dataLabel === 'string' && dataLabel.trim()) return dataLabel.trim();
  if (typeof edge.label === 'string' && edge.label.trim()) return edge.label.trim();
  return '';
}

function isDashed(edge: Edge): boolean {
  if (edge.animated) return true;
  const dash = (edge.style as { strokeDasharray?: unknown } | undefined)?.strokeDasharray;
  return typeof dash === 'string' && dash.trim().length > 0 && dash !== 'none' && dash !== '0';
}

/** Build orthogonal routes for every edge whose endpoints exist on the canvas. */
export function buildExportRoutes(edges: Edge[], boxes: Map<string, ExportBox>): ExportRoute[] {
  const routes: ExportRoute[] = [];
  for (const edge of edges) {
    const source = boxes.get(edge.source);
    const target = boxes.get(edge.target);
    if (!source || !target || source.id === target.id) continue;
    const { points, labelAnchor } = routeOrthogonal(source, target);
    routes.push({
      id: edge.id,
      sourceId: edge.source,
      targetId: edge.target,
      label: readEdgeLabel(edge),
      dashed: isDashed(edge),
      points,
      labelAnchor,
    });
  }
  return routes;
}

/**
 * Split nodes into painting order: groups first (they sit behind), then
 * services. Both lists keep their canvas order so exports stay deterministic.
 */
export function partitionBoxes(boxes: Map<string, ExportBox>): {
  groups: ExportBox[];
  services: ExportBox[];
} {
  const groups: ExportBox[] = [];
  const services: ExportBox[] = [];
  for (const box of boxes.values()) {
    if (box.kind === 'group') groups.push(box);
    else services.push(box);
  }
  return { groups, services };
}
