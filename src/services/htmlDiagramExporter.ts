// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

/**
 * HTML Diagram Exporter
 *
 * Exports the current React Flow diagram as a self-contained interactive HTML
 * file. Uses dagre for layout computation and renders Azure-branded nodes,
 * edges, groups, tooltips, and pan/zoom — the same visual style produced by
 * the render_diagram MCP tool.
 */

import dagre from 'dagre';
import type { Node, Edge } from 'reactflow';
import { rasterizeIcons } from '../utils/exportIconRaster';
import {
  advanceWidthIn,
  buildExportRoutes,
  categoryStyle,
  collectExportBoxes,
  compactEmptyGutters,
  computeBounds,
  metaSubline,
  partitionBoxes,
  connectionLegendForRoutes,
  zoneStyleFor,
  GEOMETRY_FONT_STACK,
  type ExportBox,
  type ExportRoute,
  type Point,
  type ConnectionLegendEntry,
} from './diagramExportGeometry';

// ── Types ──────────────────────────────────────────────────────────────

interface PositionedNode {
  id: string;
  name: string;
  type: string;
  description: string;
  category: string;
  meta: string;
  x: number;
  y: number;
  width: number;
  height: number;
  color: string;
  bg: string;
  textColor: string;
  icon: string;
}

interface PositionedEdge {
  id: string;
  label: string;
  color: string;
  dashed: boolean;
  dashPattern: string;
  opacity: number;
  bidirectional: boolean;
  connectionType: string;
  /** Workflow step this arrow carries, drawn as a numbered callout. */
  stepNumber?: number;
  points: Array<{ x: number; y: number }>;
  labelAnchor: { x: number; y: number };
  labelPosition: Point;
  stepAnchor: Point;
  labelLeader?: Point[];
}

interface PositionedGroup {
  id: string;
  label: string;
  x: number;
  y: number;
  width: number;
  height: number;
  color: string;
  bg: string;
  textColor: string;
}

interface LegendEntry {
  type: string;
  label: string;
  color: string;
  dashed: boolean;
  dashPattern: string;
  opacity: number;
  hasMixedStyles: boolean;
}

interface LayoutResult {
  nodes: PositionedNode[];
  edges: PositionedEdge[];
  groups: PositionedGroup[];
  connectionLegend: LegendEntry[];
  width: number;
  height: number;
}

// ── Layout via the shared geometry layer ───────────────────────────────

const PADDING = 40;
const EDGE_LABEL_FONT_PX = 10;
const STEP_HALO_RADIUS_PX = 11;

interface LabelRect {
  x: number;
  y: number;
  width: number;
  height: number;
}

function labelRect(label: string, position: Point): LabelRect {
  const width = advanceWidthIn(label, EDGE_LABEL_FONT_PX * 72 / 96) * 96 + 6;
  return { x: position.x - width / 2, y: position.y - EDGE_LABEL_FONT_PX - 3, width,
    height: EDGE_LABEL_FONT_PX * 1.5 + 6 };
}

function overlaps(a: LabelRect, b: LabelRect, gap = 4): boolean {
  return a.x < b.x + b.width + gap && a.x + a.width + gap > b.x
    && a.y < b.y + b.height + gap && a.y + a.height + gap > b.y;
}

function positionAnnotations(routes: ExportRoute[], services: ExportBox[], edges: Edge[]) {
  const dataById = new Map(edges.map(edge => [edge.id, edge.data as Record<string, unknown> | undefined]));
  const finite = (value: unknown): value is number => typeof value === 'number' && Number.isFinite(value);
  const positioned = routes.map(route => {
    const data = dataById.get(route.id);
    const x = data?.labelOffsetX;
    const y = data?.labelOffsetY;
    const labelAnchor = { x: route.labelAnchor.x + (finite(x) ? x : 0), y: route.labelAnchor.y + (finite(y) ? y : 0) };
    return {
      ...route, labelAnchor,
      labelPosition: { x: labelAnchor.x, y: labelAnchor.y - (route.stepNumber ? 18 : 4) },
      stepAnchor: { ...route.labelAnchor },
      manual: data?.labelOffsetAuto !== true && (finite(x) || finite(y)),
      labelLeader: undefined as Point[] | undefined,
    };
  });
  const nodeRects = services.map(box => ({ x: box.x, y: box.y, width: box.w, height: box.h }));
  const badges = positioned.filter(route => route.stepNumber !== undefined && !route.isSelfLoop)
    .map(route => route.stepAnchor);
  for (const route of positioned) {
    if (route.stepNumber === undefined || !route.isSelfLoop) continue;
    const preferred = route.stepAnchor;
    const candidates = [preferred];
    for (let i = 1; i < route.points.length; i++) {
      const start = route.points[i - 1], end = route.points[i];
      const samples = Math.max(1, Math.min(32, Math.ceil(Math.hypot(end.x - start.x, end.y - start.y) / 5)));
      for (let sample = 0; sample <= samples; sample++) {
        candidates.push({ x: start.x + (end.x - start.x) * sample / samples,
          y: start.y + (end.y - start.y) * sample / samples });
      }
    }
    candidates.sort((a, b) => Math.hypot(a.x - preferred.x, a.y - preferred.y)
      - Math.hypot(b.x - preferred.x, b.y - preferred.y));
    // A badge may slide along its own loop, never into an unrelated label column.
    route.stepAnchor = candidates.find(point =>
      badges.every(other => Math.hypot(point.x - other.x, point.y - other.y) >= STEP_HALO_RADIUS_PX * 2 + 2)
      && nodeRects.every(box => {
        const x = Math.max(box.x, Math.min(point.x, box.x + box.width));
        const y = Math.max(box.y, Math.min(point.y, box.y + box.height));
        return Math.hypot(point.x - x, point.y - y) >= STEP_HALO_RADIUS_PX + 2;
      })) ?? preferred;
    badges.push(route.stepAnchor);
  }
  const badgeRects = positioned.filter(route => route.stepNumber !== undefined).map(route => ({
    id: route.id, x: route.stepAnchor.x - STEP_HALO_RADIUS_PX, y: route.stepAnchor.y - STEP_HALO_RADIUS_PX,
    width: STEP_HALO_RADIUS_PX * 2, height: STEP_HALO_RADIUS_PX * 2,
  }));
  const reserved = positioned.filter(route => route.label && (!route.isSelfLoop || route.manual))
    .map(route => labelRect(route.label, route.labelPosition));
  const loopRight = new Map<string, number>();
  for (const route of positioned.filter(route => route.isSelfLoop)) {
    for (const point of route.points) {
      loopRight.set(route.sourceId, Math.max(loopRight.get(route.sourceId) ?? -Infinity, point.x));
    }
  }
  for (const route of positioned) {
    if (!route.label || !route.isSelfLoop || route.manual) continue;
    const desired = route.labelPosition;
    const blockers = [...nodeRects, ...reserved, ...badgeRects.filter(badge => badge.id !== route.id)];
    let rect = labelRect(route.label, desired);
    if (blockers.some(box => overlaps(rect, box))) {
      const x = loopRight.get(route.sourceId)! + STEP_HALO_RADIUS_PX + 12 + rect.width / 2;
      let seat: Point | undefined;
      for (let row = 0; row < 32; row++) {
        const candidate = { x, y: desired.y + row * (rect.height + 6) };
        if (blockers.every(box => !overlaps(labelRect(route.label, candidate), box))) {
          seat = candidate;
          break;
        }
      }
      // A tall neighbouring card can occupy the whole local column. The outer
      // column is guaranteed clear without moving nodes or changing any route.
      route.labelPosition = seat ?? {
        x: blockers.reduce((right, box) => Math.max(right, box.x + box.width), x) + 12 + rect.width / 2,
        y: desired.y,
      };
      rect = labelRect(route.label, route.labelPosition);
      route.labelLeader = [route.stepAnchor, { x: rect.x - 5, y: rect.y + rect.height / 2 }];
    }
    reserved.push(rect);
  }
  return positioned;
}

/**
 * Are the node positions genuinely present, or is everything stacked at the
 * origin (in which case we fall back to an automatic dagre layout)?
 */
function positionsPresent(nodes: Node[], boxes: Map<string, ExportBox>): boolean {
  const services = nodes.filter((node) => node.type !== 'groupNode');
  if (services.length === 0) return false;
  if (!nodes.every(node => Number.isFinite(node.position?.x) && Number.isFinite(node.position?.y))) return false;
  // One placed service is a layout, even at the origin. Equal local offsets
  // in different parents are not a stack: compare the resolved canvas boxes.
  if (services.length === 1) return true;
  return new Set(services.map(node => {
    const box = boxes.get(node.id)!;
    return `${box.x},${box.y}`;
  })).size > 1;
}

/** Run dagre only when the user has no real layout, writing positions back. */
function assignDagrePositions(
  nodes: Node[],
  services: ExportBox[],
  groups: ExportBox[],
  edges: Edge[],
): void {
  const g = new dagre.graphlib.Graph({ compound: true, multigraph: true });
  g.setGraph({ rankdir: 'TB', nodesep: 60, ranksep: 80, edgesep: 30, marginx: PADDING, marginy: PADDING });
  g.setDefaultEdgeLabel(() => ({}));

  const groupIds = new Set(groups.map((group) => group.id));
  for (const group of groups) {
    g.setNode(`group:${group.id}`, { label: group.label, clusterLabelPos: 'top', width: group.w, height: group.h });
  }
  const parentOf = new Map(nodes.map((node) => [node.id, node.parentNode]));
  for (const group of groups) {
    const parent = parentOf.get(group.id);
    if (!parent || !groupIds.has(parent)) continue;
    const ancestors = new Set([group.id]);
    let ancestor: string | undefined = parent;
    while (ancestor && !ancestors.has(ancestor)) {
      ancestors.add(ancestor);
      ancestor = parentOf.get(ancestor);
    }
    if (!ancestor) g.setParent(`group:${group.id}`, `group:${parent}`);
  }
  for (const service of services) {
    g.setNode(`svc:${service.id}`, { width: service.w, height: service.h, label: service.label });
    const parent = parentOf.get(service.id);
    if (parent && groupIds.has(parent)) g.setParent(`svc:${service.id}`, `group:${parent}`);
  }
  const serviceIds = new Set(services.map((service) => service.id));
  for (const edge of edges) {
    if (serviceIds.has(edge.source) && serviceIds.has(edge.target)) {
      g.setEdge(`svc:${edge.source}`, `svc:${edge.target}`, {}, edge.id);
    }
  }

  dagre.layout(g);

  for (const service of services) {
    const laid = g.node(`svc:${service.id}`);
    if (laid) { service.x = laid.x - service.w / 2; service.y = laid.y - service.h / 2; }
  }
  for (const group of groups) {
    const laid = g.node(`group:${group.id}`);
    if (laid && laid.width && laid.height) {
      group.x = laid.x - laid.width / 2;
      group.y = laid.y - laid.height / 2;
      group.w = laid.width;
      group.h = laid.height;
    }
  }
}

/**
 * Flatten React Flow nodes/edges into a positioned layout, honouring the user's
 * real coordinates (fix 1), the shared category palette + real icons (fix 2),
 * per-connection colour (fix 4), zone colours (fix 6) and metadata (fix 10).
 */
function buildLayout(nodes: Node[], edges: Edge[], icons: Map<string, string>): LayoutResult {
  // Empty bands are closed first, the same way the PPTX and Visio exporters do
  // it, so the PNG and the deck are the same drawing rather than one being a
  // fiftieth-scale version of the other.
  const boxes = compactEmptyGutters(collectExportBoxes(nodes));
  const { groups, services } = partitionBoxes(boxes);
  const dataById = new Map(nodes.map((node) => [node.id, (node.data ?? {}) as Record<string, unknown>]));

  if (!positionsPresent(nodes, boxes)) {
    assignDagrePositions(nodes, services, groups, edges);
    groups.sort((a, b) => b.w * b.h - a.w * a.h);
  }

  const routes = positionAnnotations(buildExportRoutes(edges, boxes), services, edges);
  const bounds = computeBounds(boxes.values());
  const include = (x: number, y: number): void => {
    bounds.minX = Math.min(bounds.minX, x);
    bounds.minY = Math.min(bounds.minY, y);
    bounds.maxX = Math.max(bounds.maxX, x);
    bounds.maxY = Math.max(bounds.maxY, y);
  };
  for (const route of routes) {
    for (const point of route.points) {
      include(point.x - 6, point.y - 6);
      include(point.x + 6, point.y + 6);
    }
    const at = route.stepAnchor;
    if (route.stepNumber !== undefined) {
      include(at.x - STEP_HALO_RADIUS_PX, at.y - STEP_HALO_RADIUS_PX);
      include(at.x + STEP_HALO_RADIUS_PX, at.y + STEP_HALO_RADIUS_PX);
    }
    if (route.label) {
      const rect = labelRect(route.label, route.labelPosition);
      include(rect.x, rect.y);
      include(rect.x + rect.width, rect.y + rect.height);
    }
    for (const point of route.labelLeader ?? []) {
      include(point.x - 1, point.y - 1);
      include(point.x + 1, point.y + 1);
    }
  }
  const dx = PADDING - bounds.minX;
  const dy = PADDING - bounds.minY;

  const positionedNodes: PositionedNode[] = services.map((box) => {
    const style = categoryStyle(box.category);
    const data = dataById.get(box.id) ?? {};
    return {
      id: box.id,
      name: box.label,
      type: box.serviceName ?? box.category,
      description: typeof data.description === 'string' ? data.description : '',
      category: box.category,
      meta: metaSubline(box),
      x: box.x + dx,
      y: box.y + dy,
      width: box.w,
      height: box.h,
      color: style.border,
      bg: style.bg,
      textColor: style.text,
      icon: (box.iconPath && icons.get(box.iconPath)) || '',
    };
  });

  const positionedGroups: PositionedGroup[] = groups.map((box) => {
    const style = zoneStyleFor(box);
    return {
      id: box.id,
      label: box.label,
      x: box.x + dx,
      y: box.y + dy,
      width: box.w,
      height: box.h,
      color: style.border,
      bg: style.bg,
      textColor: style.text,
    };
  });

  const positionedEdges: PositionedEdge[] = routes.map(route => ({
    id: route.id,
    label: route.label,
    color: route.color,
    dashed: route.dashed,
    dashPattern: route.dashPattern ?? '',
    opacity: route.opacity,
    bidirectional: route.bidirectional,
    connectionType: route.connectionType,
    ...(route.stepNumber !== undefined ? { stepNumber: route.stepNumber } : {}),
    points: route.points.map((point) => ({ x: point.x + dx, y: point.y + dy })),
    labelAnchor: { x: route.labelAnchor.x + dx, y: route.labelAnchor.y + dy },
    labelPosition: { x: route.labelPosition.x + dx, y: route.labelPosition.y + dy },
    stepAnchor: { x: route.stepAnchor.x + dx, y: route.stepAnchor.y + dy },
    ...(route.labelLeader ? { labelLeader: route.labelLeader.map(point => ({ x: point.x + dx, y: point.y + dy })) } : {}),
  }));

  const connectionLegend: LegendEntry[] = connectionLegendForRoutes(routes).map((entry: ConnectionLegendEntry) => ({
    type: entry.type,
    label: entry.label,
    color: entry.color,
    dashed: entry.dashed,
    dashPattern: entry.dashPattern ?? '',
    opacity: entry.opacity,
    hasMixedStyles: entry.hasMixedStyles === true,
  }));

  const width = Math.max(1, bounds.maxX - bounds.minX) + PADDING * 2;
  const height = Math.max(1, bounds.maxY - bounds.minY) + PADDING * 2;

  return { nodes: positionedNodes, edges: positionedEdges, groups: positionedGroups, connectionLegend, width, height };
}

// ── HTML generation ────────────────────────────────────────────────────

function esc(s: string): string {
  return s.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
}

function generateHtml(layout: LayoutResult, title: string): string {
  const layoutJson = JSON.stringify(layout)
    .replace(/</g, '\\u003c')
    .replace(/\u2028/g, '\\u2028')
    .replace(/\u2029/g, '\\u2029');

  return `<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>${esc(title)}</title>
<style>
  * { margin: 0; padding: 0; box-sizing: border-box; }
  body { font-family: ${GEOMETRY_FONT_STACK}; background: #f8f9fa; overflow: hidden; }
  .header {
    background: linear-gradient(135deg, #0078D4, #005A9E);
    color: white; padding: 12px 24px; display: flex; align-items: center; gap: 16px;
    box-shadow: 0 2px 8px rgba(0,0,0,0.15); z-index: 100; position: relative;
  }
  .header h1 { font-size: 18px; font-weight: 600; }
  .header .meta { font-size: 12px; opacity: 0.8; margin-left: auto; }
  .header .controls { display: flex; gap: 6px; }
  .header button {
    background: rgba(255,255,255,0.2); border: 1px solid rgba(255,255,255,0.3);
    color: white; padding: 4px 12px; border-radius: 4px; cursor: pointer; font-size: 12px;
  }
  .header button:hover { background: rgba(255,255,255,0.3); }
  .canvas-container { width: 100vw; height: calc(100vh - 52px); overflow: hidden; position: relative; cursor: grab; }
  .canvas-container.dragging { cursor: grabbing; }
  .canvas { position: absolute; transform-origin: 0 0; }
  .node {
    position: absolute; background: white;
    border-radius: 8px; border: 2px solid #ccc; box-shadow: 0 2px 8px rgba(0,0,0,0.08);
    cursor: pointer; transition: box-shadow 0.2s, transform 0.2s;
    display: flex; flex-direction: row; align-items: center; gap: 8px; padding: 6px 10px;
    overflow: hidden;
  }
  .node:hover { box-shadow: 0 4px 16px rgba(0,0,0,0.15); transform: translateY(-1px); z-index: 10; }
  .node.highlighted { box-shadow: 0 0 0 3px rgba(0,120,212,0.4), 0 4px 16px rgba(0,0,0,0.15); }
  .node .node-icon { width: 34px; height: 34px; object-fit: contain; flex-shrink: 0; }
  .node .node-mono {
    width: 34px; height: 34px; border-radius: 6px; flex-shrink: 0; color: #fff;
    display: flex; align-items: center; justify-content: center; font-weight: 700; font-size: 15px;
  }
  .node .node-body { display: flex; flex-direction: column; overflow: hidden; min-width: 0; }
  .node .name { font-size: 13px; font-weight: 600; color: #1B1B1B; white-space: nowrap; overflow: hidden; text-overflow: ellipsis; }
  .node .type { font-size: 11px; white-space: nowrap; overflow: hidden; text-overflow: ellipsis; }
  .node .meta { font-size: 10px; color: #64748b; white-space: nowrap; overflow: hidden; text-overflow: ellipsis; }
  .edges-layer { position: absolute; top: 0; left: 0; pointer-events: none; }
  .edge-path { fill: none; stroke-width: 1.5; }
  .edge-step {
    font-family: ${GEOMETRY_FONT_STACK};
    font-size: 11px; font-weight: 700;
  }
  .edge-label {
    font-family: ${GEOMETRY_FONT_STACK}; font-size: ${EDGE_LABEL_FONT_PX}px;
    paint-order: stroke; stroke: white; stroke-width: 3px;
  }
  .group {
    position: absolute; border-radius: 12px; border: 1.5px dashed;
  }
  .group .group-label {
    position: absolute; top: -18px; left: 50%; transform: translateX(-50%);
    font-size: 12px; font-weight: 600; white-space: nowrap;
  }
  .tooltip {
    position: fixed; display: none; background: #1B1B1B; color: white;
    padding: 8px 12px; border-radius: 6px; font-size: 12px; max-width: 280px;
    box-shadow: 0 4px 12px rgba(0,0,0,0.3); z-index: 1000; pointer-events: none;
  }
  .tooltip .tt-name { font-weight: 600; margin-bottom: 4px; }
  .tooltip .tt-type { opacity: 0.7; font-size: 11px; }
  .tooltip .tt-desc { margin-top: 4px; font-size: 11px; opacity: 0.85; }
  .legend {
    position: fixed; bottom: 12px; left: 12px; background: white;
    border-radius: 8px; padding: 10px 14px; box-shadow: 0 2px 8px rgba(0,0,0,0.1);
    font-size: 11px; display: flex; gap: 12px; flex-wrap: wrap; max-width: 600px; z-index: 50;
  }
  .legend-item { display: flex; align-items: center; gap: 4px; }
  .legend-dot { width: 10px; height: 10px; border-radius: 50%; }
  .legend-sep { width: 100%; height: 0; border-top: 1px solid #e5e7eb; margin: 2px 0; }
  .legend-line { width: 22px; height: 8px; flex-shrink: 0; }
</style>
</head>
<body>

<div class="header">
  <h1>${esc(title)}</h1>
  <div class="controls">
    <button onclick="zoomIn()">+</button>
    <button onclick="zoomOut()">\\u2212</button>
    <button onclick="resetView()">Reset</button>
    <button onclick="fitView()">Fit</button>
  </div>
  <div class="meta">Generated by Microsoft Product Architecture Diagram Builder &middot; Swarm Data SE, Jiayi Yang</div>
</div>

<div class="canvas-container" id="container">
  <div class="canvas" id="canvas"></div>
</div>

<div class="tooltip" id="tooltip">
  <div class="tt-name"></div>
  <div class="tt-type"></div>
  <div class="tt-desc"></div>
</div>

<div class="legend" id="legend"></div>

<script>
const layout = ${layoutJson};

let scale = 1, offsetX = 0, offsetY = 0, isDragging = false, dragStartX = 0, dragStartY = 0;
const container = document.getElementById('container');
const canvas = document.getElementById('canvas');
const tooltip = document.getElementById('tooltip');

function monogram(name) {
  const s = (name || '?').trim();
  return s ? s.charAt(0).toUpperCase() : '?';
}

function render() {
  canvas.innerHTML = '';
  canvas.style.width = layout.width + 'px';
  canvas.style.height = layout.height + 'px';

  layout.groups.forEach((g) => {
    const el = document.createElement('div');
    el.className = 'group';
    el.style.left = (g.x - 12) + 'px';
    el.style.top = (g.y - 32) + 'px';
    el.style.width = (g.width + 24) + 'px';
    el.style.height = (g.height + 44) + 'px';
    el.style.borderColor = g.color;
    // g.bg is already the accent composited onto the page by zoneStyleFor.
    // Appending an alpha byte applied that 8-10% tint a second time, so the
    // panel came out at well under 1% of the accent: a green zone rendered
    // #f6f9f9 where the canvas shows #e1f3ee, which is to say the zone tint
    // the reader picked was invisible in the file.
    el.style.background = g.bg;
    // The AA-guaranteed ink, not the raw accent. This label is not the canvas
    // header: the export drops the group-node-header bar and floats the title
    // above the panel on the bare page, so the accent has no tinted bar behind
    // it to read as a header treatment, and amber landed at 2.04:1 and green
    // at 2.41:1. Every other exporter already draws the title in style.text.
    el.innerHTML = '<div class="group-label" style="color:' + g.textColor + '">' + esc(g.label) + '</div>';
    canvas.appendChild(el);
  });

  const svgNs = 'http://www.w3.org/2000/svg';
  const svg = document.createElementNS(svgNs, 'svg');
  svg.classList.add('edges-layer');
  svg.setAttribute('width', layout.width);
  svg.setAttribute('height', layout.height);
  svg.style.width = layout.width + 'px';
  svg.style.height = layout.height + 'px';

  // One arrow marker per distinct connection colour so heads match their line.
  const defs = document.createElementNS(svgNs, 'defs');
  const markerByColor = {};
  const startMarkerByColor = {};
  let markerSeq = 0;
  const arrowMarker = (color, id, atStart) => {
    const marker = document.createElementNS(svgNs, 'marker');
    marker.setAttribute('id', id);
    marker.setAttribute('viewBox', '0 0 10 10');
    marker.setAttribute('refX', atStart ? '0' : '10'); marker.setAttribute('refY', '5');
    marker.setAttribute('markerWidth', '8'); marker.setAttribute('markerHeight', '8');
    marker.setAttribute('orient', 'auto');
    const poly = document.createElementNS(svgNs, 'polygon');
    poly.setAttribute('points', atStart ? '10,0 0,5 10,10' : '0,0 10,5 0,10');
    poly.setAttribute('fill', color);
    marker.appendChild(poly);
    defs.appendChild(marker);
  };
  layout.edges.forEach(e => {
    if (!markerByColor[e.color]) {
      const id = 'arrow-' + (markerSeq++);
      markerByColor[e.color] = id;
      arrowMarker(e.color, id, false);
    }
    if (e.bidirectional && !startMarkerByColor[e.color]) {
      const id = 'arrow-start-' + (markerSeq++);
      startMarkerByColor[e.color] = id;
      arrowMarker(e.color, id, true);
    }
  });
  svg.appendChild(defs);

  layout.edges.forEach(e => {
    if (e.points.length < 2) return;
    const color = e.color || '#64748b';
    const d = e.points.map((p, i) => (i === 0 ? 'M' : 'L') + p.x + ' ' + p.y).join(' ');
    const path = document.createElementNS(svgNs, 'path');
    path.setAttribute('d', d);
    path.setAttribute('stroke', color);
    path.setAttribute('opacity', String(e.opacity));
    path.setAttribute('data-edge-id', e.id);
    path.classList.add('edge-path');
    path.setAttribute('marker-end', 'url(#' + markerByColor[e.color] + ')');
    if (e.bidirectional) path.setAttribute('marker-start', 'url(#' + startMarkerByColor[e.color] + ')');
    if (e.dashed) path.setAttribute('stroke-dasharray', e.dashPattern || '6,4');
    svg.appendChild(path);
  });

  layout.edges.forEach(e => {
    if (!e.labelLeader) return;
    const leader = document.createElementNS(svgNs, 'path');
    leader.setAttribute('d', e.labelLeader.map((p, i) => (i ? 'L' : 'M') + p.x + ' ' + p.y).join(' '));
    leader.setAttribute('data-edge-id', e.id);
    leader.setAttribute('stroke', e.color);
    leader.setAttribute('stroke-width', '1');
    leader.setAttribute('stroke-dasharray', '2,3');
    leader.setAttribute('fill', 'none');
    leader.classList.add('edge-label-leader');
    svg.appendChild(leader);
  });

  layout.edges.forEach(e => {
    const color = e.color || '#64748b';
    if (e.label) {
      const position = e.labelPosition;
      const text = document.createElementNS(svgNs, 'text');
      text.setAttribute('x', position.x);
      text.setAttribute('y', position.y);
      text.setAttribute('data-edge-id', e.id);
      text.setAttribute('text-anchor', 'middle');
      text.setAttribute('fill', color);
      text.classList.add('edge-label');
      text.textContent = e.label;
      svg.appendChild(text);
    }

    // Numbered callout, matching the workflow list: the Azure Architecture
    // Center convention that ties each arrow to the step describing it.
    if (e.stepNumber) {
      const mid = e.stepAnchor;
      const halo = document.createElementNS(svgNs, 'circle');
      halo.setAttribute('cx', mid.x);
      halo.setAttribute('cy', mid.y);
      halo.setAttribute('r', '${STEP_HALO_RADIUS_PX}');
      halo.setAttribute('fill', '#ffffff');
      svg.appendChild(halo);
      const disc = document.createElementNS(svgNs, 'circle');
      disc.setAttribute('cx', mid.x);
      disc.setAttribute('cy', mid.y);
      disc.setAttribute('r', '9');
      disc.setAttribute('fill', color);
      svg.appendChild(disc);
      const num = document.createElementNS(svgNs, 'text');
      num.setAttribute('x', mid.x);
      num.setAttribute('y', mid.y + 4);
      num.setAttribute('data-edge-id', e.id);
      num.setAttribute('text-anchor', 'middle');
      num.setAttribute('fill', '#ffffff');
      num.classList.add('edge-step');
      num.textContent = String(e.stepNumber);
      svg.appendChild(num);
    }
  });
  canvas.appendChild(svg);

  layout.nodes.forEach(n => {
    const el = document.createElement('div');
    el.className = 'node';
    el.style.left = n.x + 'px';
    el.style.top = n.y + 'px';
    el.style.width = n.width + 'px';
    el.style.height = n.height + 'px';
    el.style.borderColor = n.color;
    el.style.background = n.bg || '#ffffff';
    const iconHtml = n.icon
      ? '<img class="node-icon" src="' + n.icon + '" alt="" />'
      : '<div class="node-mono" style="background:' + n.color + '">' + esc(monogram(n.name)) + '</div>';
    el.innerHTML =
      iconHtml +
      '<div class="node-body">' +
        '<div class="name">' + esc(n.name) + '</div>' +
        '<div class="type" style="color:' + n.textColor + '">' + esc(n.type) + '</div>' +
        (n.meta ? '<div class="meta">' + esc(n.meta) + '</div>' : '') +
      '</div>';

    el.addEventListener('mouseenter', ev => showTooltip(ev, n));
    el.addEventListener('mouseleave', hideTooltip);
    el.addEventListener('click', () => {
      document.querySelectorAll('.node').forEach(nd => nd.classList.remove('highlighted'));
      el.classList.toggle('highlighted');
    });
    canvas.appendChild(el);
  });

  // Legend: category colour dots plus the connection-type colour key so the
  // interactive view agrees with the PNG/PPTX/VSDX legends.
  const catColors = {};
  layout.nodes.forEach(n => { catColors[n.category] = n.color; });
  const cats = Object.keys(catColors).sort();
  const legendEl = document.getElementById('legend');
  let legendHtml = cats.map(c =>
    '<div class="legend-item"><div class="legend-dot" style="background:' + catColors[c] + '"></div>' + esc(c) + '</div>'
  ).join('');
  const conn = layout.connectionLegend || [];
  if (conn.length) {
    legendHtml += '<div class="legend-sep"></div>' + conn.map(c => {
      const swatch = c.hasMixedStyles ? '' :
        '<svg class="legend-line" viewBox="0 0 22 8" aria-hidden="true"><path d="M0 4H22" fill="none" stroke-width="1.5" stroke="' +
        esc(c.color) + '" opacity="' + c.opacity + '"' +
        (c.dashed ? ' stroke-dasharray="' + esc(c.dashPattern || '6,4') + '"' : '') + '/></svg>';
      return '<div class="legend-item">' + swatch + esc(c.label) + '</div>';
    }).join('');
  }
  legendEl.innerHTML = legendHtml;

  applyTransform();
}

function applyTransform() {
  canvas.style.transform = 'translate(' + offsetX + 'px,' + offsetY + 'px) scale(' + scale + ')';
}

function showTooltip(ev, n) {
  tooltip.style.display = 'block';
  tooltip.style.left = (ev.clientX + 12) + 'px';
  tooltip.style.top = (ev.clientY + 12) + 'px';
  tooltip.querySelector('.tt-name').textContent = n.name;
  tooltip.querySelector('.tt-type').textContent = n.type + ' (' + n.category + ')';
  tooltip.querySelector('.tt-desc').textContent = n.description || '';
}
function hideTooltip() { tooltip.style.display = 'none'; }

container.addEventListener('mousedown', e => {
  if (e.target.closest('.node')) return;
  isDragging = true; dragStartX = e.clientX - offsetX; dragStartY = e.clientY - offsetY;
  container.classList.add('dragging');
});
window.addEventListener('mousemove', e => {
  if (!isDragging) return;
  offsetX = e.clientX - dragStartX; offsetY = e.clientY - dragStartY;
  applyTransform();
});
window.addEventListener('mouseup', () => { isDragging = false; container.classList.remove('dragging'); });

container.addEventListener('wheel', e => {
  e.preventDefault();
  const delta = e.deltaY > 0 ? 0.9 : 1.1;
  const rect = container.getBoundingClientRect();
  const mx = e.clientX - rect.left;
  const my = e.clientY - rect.top;
  const newScale = Math.max(0.1, Math.min(5, scale * delta));
  offsetX = mx - (mx - offsetX) * (newScale / scale);
  offsetY = my - (my - offsetY) * (newScale / scale);
  scale = newScale;
  applyTransform();
}, { passive: false });

function zoomIn() { scale = Math.min(5, scale * 1.2); applyTransform(); }
function zoomOut() { scale = Math.max(0.1, scale * 0.8); applyTransform(); }
function resetView() { scale = 1; offsetX = 0; offsetY = 0; applyTransform(); }
function fitView() {
  const cw = container.clientWidth;
  const ch = container.clientHeight;
  scale = Math.min(cw / layout.width, ch / layout.height) * 0.9;
  offsetX = (cw - layout.width * scale) / 2;
  offsetY = (ch - layout.height * scale) / 2;
  applyTransform();
}

function esc(s) { const d = document.createElement('div'); d.textContent = s; return d.innerHTML; }

render();
fitView();
<\/script>
</body>
</html>`;
}

// ── Public API ─────────────────────────────────────────────────────────

/**
 * Rasterise the Azure icons referenced by the service nodes into base64 PNG
 * data URIs so the exported HTML is fully self-contained (fix 2). In a
 * non-browser context this resolves to an empty map and the renderer falls
 * back to a coloured monogram.
 */
async function buildIconDataUrls(nodes: Node[]): Promise<Map<string, string>> {
  const paths = new Set<string>();
  for (const node of nodes) {
    const iconPath = (node.data as { iconPath?: unknown } | undefined)?.iconPath;
    if (typeof iconPath === 'string' && iconPath) paths.add(iconPath);
  }
  const result = new Map<string, string>();
  if (paths.size === 0) return result;
  try {
    const rastered = await rasterizeIcons(Array.from(paths), 64);
    for (const [path, raster] of rastered) {
      if (raster?.dataUrl) result.set(path, raster.dataUrl);
    }
  } catch {
    // Non-browser or fetch failure — monogram fallback is used.
  }
  return result;
}

export async function exportDiagramAsHtml(
  nodes: Node[],
  edges: Edge[],
  title?: string,
): Promise<void> {
  const diagramTitle = title || 'Azure Architecture Diagram';
  const html = await buildInteractiveDiagramHtml(nodes, edges, diagramTitle);

  if (!html) {
    alert('No services to export. Add Azure services to the diagram first.');
    return;
  }

  // Trigger download
  const blob = new Blob([html], { type: 'text/html' });
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url;
  a.download = `${diagramTitle.replace(/[^a-zA-Z0-9-_ ]/g, '').replace(/\s+/g, '-').toLowerCase()}.html`;
  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);
  URL.revokeObjectURL(url);
}

export async function buildInteractiveDiagramHtml(
  nodes: Node[],
  edges: Edge[],
  title?: string,
): Promise<string | null> {
  const diagramTitle = title || 'Azure Architecture Diagram';
  const hasServices = nodes.some((node) => node.type !== 'groupNode');
  if (!hasServices) {
    return null;
  }

  const icons = await buildIconDataUrls(nodes);
  const layout = buildLayout(nodes, edges, icons);
  return generateHtml(layout, diagramTitle);
}
