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
  buildExportRoutes,
  categoryStyle,
  collectExportBoxes,
  computeBounds,
  metaSubline,
  partitionBoxes,
  usedConnectionLegend,
  zoneStyleFor,
  type ExportBox,
  type ExportRoute,
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
  bidirectional: boolean;
  connectionType: string;
  points: Array<{ x: number; y: number }>;
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
}

interface LegendEntry {
  type: string;
  label: string;
  color: string;
  dashed: boolean;
  dashPattern: string;
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

/**
 * Are the node positions genuinely present, or is everything stacked at the
 * origin (in which case we fall back to an automatic dagre layout)?
 */
function positionsPresent(nodes: Node[]): boolean {
  const services = nodes.filter((node) => node.type !== 'groupNode');
  if (services.length === 0) return false;
  const seen = new Set<string>();
  let anyNonZero = false;
  for (const node of services) {
    const x = node.position?.x ?? 0;
    const y = node.position?.y ?? 0;
    if (x !== 0 || y !== 0) anyNonZero = true;
    seen.add(`${x},${y}`);
  }
  // Real layouts have distinct, non-zero coordinates; a fresh AI import often
  // has none, so only then do we synthesise positions.
  return anyNonZero && seen.size > 1;
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
    g.setNode(`group:${group.id}`, { label: group.label, clusterLabelPos: 'top' });
  }
  const parentOf = new Map(nodes.map((node) => [node.id, node.parentNode]));
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
  const boxes = collectExportBoxes(nodes);
  const { groups, services } = partitionBoxes(boxes);
  const dataById = new Map(nodes.map((node) => [node.id, (node.data ?? {}) as Record<string, unknown>]));

  if (!positionsPresent(nodes)) {
    assignDagrePositions(nodes, services, groups, edges);
  }

  const routes = buildExportRoutes(edges, boxes);
  const bounds = computeBounds(boxes.values());
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

  const positionedGroups: PositionedGroup[] = groups.map((box, index) => {
    const style = zoneStyleFor(box, index);
    return {
      id: box.id,
      label: box.label,
      x: box.x + dx,
      y: box.y + dy,
      width: box.w,
      height: box.h,
      color: style.border,
      bg: style.bg,
    };
  });

  const positionedEdges: PositionedEdge[] = routes.map((route: ExportRoute) => ({
    id: route.id,
    label: route.label,
    color: route.color,
    dashed: route.dashed,
    dashPattern: route.dashPattern ?? '',
    bidirectional: route.bidirectional,
    connectionType: route.connectionType,
    points: route.points.map((point) => ({ x: point.x + dx, y: point.y + dy })),
  }));

  const connectionLegend: LegendEntry[] = usedConnectionLegend(edges).map((entry: ConnectionLegendEntry) => ({
    type: entry.type,
    label: entry.label,
    color: entry.color,
    dashed: entry.dashed,
    dashPattern: entry.dashPattern ?? '',
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
  body { font-family: 'Yu Gothic UI', 'Segoe UI', system-ui, -apple-system, sans-serif; background: #f8f9fa; overflow: hidden; }
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
  .edge-label {
    font-family: 'Yu Gothic UI', 'Segoe UI', system-ui, sans-serif; font-size: 10px;
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
  .legend-line { width: 22px; height: 0; border-top-width: 2px; border-top-style: solid; }
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
    el.style.background = g.bg + '14';
    el.innerHTML = '<div class="group-label" style="color:' + g.color + '">' + esc(g.label) + '</div>';
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
    path.classList.add('edge-path');
    path.setAttribute('marker-end', 'url(#' + markerByColor[e.color] + ')');
    if (e.bidirectional) path.setAttribute('marker-start', 'url(#' + startMarkerByColor[e.color] + ')');
    if (e.dashed) path.setAttribute('stroke-dasharray', (e.dashPattern || '6,4').replace(/\\s+/g, ''));
    svg.appendChild(path);

    if (e.label) {
      const mid = e.points[Math.floor(e.points.length / 2)];
      const text = document.createElementNS(svgNs, 'text');
      text.setAttribute('x', mid.x);
      text.setAttribute('y', mid.y - 4);
      text.setAttribute('text-anchor', 'middle');
      text.setAttribute('fill', color);
      text.classList.add('edge-label');
      text.textContent = e.label;
      svg.appendChild(text);
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
    legendHtml += '<div class="legend-sep"></div>' + conn.map(c =>
      '<div class="legend-item"><div class="legend-line" style="border-top-color:' + c.color + ';border-top-style:' + (c.dashed ? 'dashed' : 'solid') + '"></div>' + esc(c.label) + '</div>'
    ).join('');
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
