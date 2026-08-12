// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

/**
 * Draw.io (diagrams.net) Export Service
 * 
 * Converts React Flow diagrams to draw.io compatible XML format (.drawio)
 * enabling users to import and edit diagrams in draw.io/diagrams.net
 * 
 * Icons are embedded as Base64-encoded data URIs for full offline support
 * 
 * @author Arturo Quiroga
 * @date January 2026
 */

import { Node, Edge } from 'reactflow';
import { generateModelFilename } from '../utils/modelNaming';
import {
  buildExportRoutes,
  categoryStyle,
  collectExportBoxes,
  compactEmptyGutters,
  computeBounds,
  metaSubline,
  partitionBoxes,
  zoneStyleFor,
  type ExportBox,
  type ExportRoute,
  type Point,
} from './diagramExportGeometry';

/** Trim coordinates to 2dp so the XML stays compact and deterministic. */
const f = (n: number): number => +n.toFixed(2);

// Load SVG icon and convert to Base64 data URI for Draw.io
async function loadIconAsBase64(iconPath: string): Promise<string | null> {
  if (!iconPath) {
    console.log('[Draw.io Export] No icon path provided');
    return null;
  }
  
  try {
    console.log('[Draw.io Export] Loading icon:', iconPath);
    
    // Fetch the SVG file
    const response = await fetch(iconPath);
    if (!response.ok) {
      console.warn(`[Draw.io Export] Failed to load icon (${response.status}): ${iconPath}`);
      return null;
    }
    
    const svgText = await response.text();
    console.log('[Draw.io Export] Icon loaded, size:', svgText.length, 'bytes');
    
    // Convert to Base64 using modern approach
    const encoder = new TextEncoder();
    const data = encoder.encode(svgText);
    let binary = '';
    const len = data.byteLength;
    for (let i = 0; i < len; i++) {
      binary += String.fromCharCode(data[i]);
    }
    const base64 = btoa(binary);
    
    console.log('[Draw.io Export] Base64 encoded, size:', base64.length, 'chars');
    
    // Return as data URI - Draw.io format: data:image/svg+xml,BASE64
    // Note: Draw.io uses comma separator, not semicolon+base64
    return `data:image/svg+xml,${base64}`;
  } catch (error) {
    console.error(`[Draw.io Export] Error loading icon ${iconPath}:`, error);
    return null;
  }
}

// Draw.io XML escape helper
function escapeXml(str: string): string {
  return str
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&apos;');
}

// Generate unique IDs for draw.io cells
let cellIdCounter = 2; // 0 and 1 are reserved for root cells
function generateCellId(): string {
  return `cell-${cellIdCounter++}`;
}

// Reset cell ID counter (call before each export)
function resetCellIdCounter(): void {
  cellIdCounter = 2;
}

// Connection type → draw.io dash + colour, derived once in the shared layer so
// Draw.io agrees with PPTX/VSDX/HTML and the PNG legend (fixes 4 & 11).
function drawioEdgeStyle(route: ExportRoute): string {
  let dash = 'dashed=0;';
  if (route.dashed) {
    switch (route.connectionType) {
      case 'async':
        dash = 'dashed=1;dashPattern=8 8;';
        break;
      case 'optional':
      case 'security':
        dash = 'dashed=1;dashPattern=1 4;';
        break;
      case 'telemetry':
        dash = 'dashed=1;dashPattern=8 4 2 4;';
        break;
      default:
        dash = 'dashed=1;dashPattern=6 6;';
    }
  }
  const opacity = route.opacity < 1 ? `opacity=${Math.round(route.opacity * 100)};` : '';
  return `${dash}strokeColor=${route.color};${opacity}`;
}

// Create draw.io mxCell for a group node — absolute geometry, parent="1".
// Zones and services are siblings on the root layer so a nested zone's absolute
// position (resolved in the shared layer) is honoured verbatim (fixes 8 & 16).
function createGroupCell(box: ExportBox, cellId: string, index: number): string {
  const style = zoneStyleFor(box, index);
  const label = escapeXml(box.label);
  const cellStyle = `swimlane;whiteSpace=wrap;html=1;fillColor=${style.bg};strokeColor=${style.border};fontColor=${style.text};fontStyle=1;startSize=30;rounded=1;arcSize=6;`;
  return `
      <mxCell id="${cellId}" value="${label}" style="${cellStyle}" vertex="1" parent="1">
        <mxGeometry x="${f(box.x)}" y="${f(box.y)}" width="${f(box.w)}" height="${f(box.h)}" as="geometry" />
      </mxCell>`;
}

// Create draw.io mxCell for an Azure service node with an embedded icon.
// The tile uses absolute geometry on the root layer, and the icon is a CHILD of
// the tile (parent=cellId) so moving the service in draw.io moves its icon with
// it (fix 17). Colours come from the shared category palette (fix 19) and the
// SKU · region · cost sub-line mirrors the canvas tile (fix 10).
async function createServiceCell(
  box: ExportBox,
  cellId: string,
): Promise<{ containerCell: string; iconCell: string | null }> {
  const width = box.w;
  const height = box.h;
  const style = categoryStyle(box.category);
  const label = escapeXml(box.label);
  const meta = metaSubline(box);
  const value = meta ? `${label}&#xa;${escapeXml(meta)}` : label;

  const cellStyle = `rounded=1;whiteSpace=wrap;html=1;fillColor=${style.bg};strokeColor=${style.border};fontColor=${style.text};strokeWidth=2;fontStyle=0;verticalAlign=bottom;spacingBottom=8;`;

  const containerCell = `
      <mxCell id="${cellId}" value="${value}" style="${cellStyle}" vertex="1" parent="1">
        <mxGeometry x="${f(box.x)}" y="${f(box.y)}" width="${f(width)}" height="${f(height)}" as="geometry" />
      </mxCell>`;

  let iconCell: string | null = null;
  if (box.iconPath) {
    const iconDataUri = await loadIconAsBase64(box.iconPath);
    if (iconDataUri) {
      const iconCellId = generateCellId();
      const iconSize = 48;
      // Icon geometry is RELATIVE to the parent tile so the two never drift.
      const iconX = (width - iconSize) / 2;
      const iconY = 8;
      const iconStyle = `shape=image;imageAspect=0;aspect=fixed;verticalLabelPosition=bottom;labelBackgroundColor=default;verticalAlign=top;html=1;image=${iconDataUri};`;
      iconCell = `
      <mxCell id="${iconCellId}" value="" style="${iconStyle}" vertex="1" parent="${cellId}">
        <mxGeometry x="${f(iconX)}" y="${iconY}" width="${iconSize}" height="${iconSize}" as="geometry" />
      </mxCell>`;
    }
  }

  return { containerCell, iconCell };
}

/** Waypoints (absolute) that de-collide parallel edges and draw self-loops. */
function edgeWaypoints(route: ExportRoute): Point[] {
  if (route.isSelfLoop) return route.points.slice(1, -1);
  if (route.ordinal > 0 && route.points.length >= 2) return [route.labelAnchor];
  return [];
}

// Create draw.io mxCell for an edge/connection from a shared-layer route.
function createEdgeCell(
  route: ExportRoute,
  edgeCellId: string,
  nodeIdToCellId: Map<string, string>,
): string {
  const sourceCellId = nodeIdToCellId.get(route.sourceId);
  const targetCellId = nodeIdToCellId.get(route.targetId);

  if (!sourceCellId || !targetCellId) {
    console.warn(`Edge ${route.id} references missing nodes: ${route.sourceId} -> ${route.targetId}`);
    return '';
  }

  const label = route.label;

  // Force ORTHOGONAL (right-angle) routing on export; colour + dash come from
  // the canonical connection style so every format agrees.
  const style = `edgeStyle=orthogonalEdgeStyle;curved=0;rounded=0;orthogonalLoop=1;jettySize=auto;html=1;${drawioEdgeStyle(route)}strokeWidth=2;endArrow=classic;endFill=1;${route.bidirectional ? 'startArrow=classic;startFill=1;' : ''}`;

  const waypoints = edgeWaypoints(route);
  const pointsXml = waypoints.length
    ? `
          <Array as="points">${waypoints
            .map((point) => `<mxPoint x="${f(point.x)}" y="${f(point.y)}" />`)
            .join('')}</Array>`
    : '';

  let cells = `
      <mxCell id="${edgeCellId}" value="" style="${style}" edge="1" parent="1" source="${sourceCellId}" target="${targetCellId}">
        <mxGeometry relative="1" as="geometry">${pointsXml}
        </mxGeometry>
      </mxCell>`;

  // Sized here rather than inside the label block so the numbered callout can
  // clear the chip it belongs to — a fixed offset put the badge on top of the
  // text as soon as the label wrapped to a second line.
  const labelText = label ? String(label) : '';
  const wrapWidth = 170;
  const charsPerLine = 26;
  const labelLines = labelText ? Math.max(1, Math.ceil(labelText.length / charsPerLine)) : 0;
  const labelBoxHeight = 12 + labelLines * 16;
  // Stagger parallel-edge chips along the segment so they don't stack.
  const offsetX = route.ordinal === 0 ? 0 : (route.ordinal % 2 === 1 ? 1 : -1) * Math.ceil(route.ordinal / 2) * 24;

  if (label) {
    const labelStyle = 'edgeLabel;html=1;align=center;verticalAlign=middle;whiteSpace=wrap;rounded=1;fillColor=#fef9c3;strokeColor=#374151;fontColor=#1f2937;fontStyle=1;fontSize=12;spacingLeft=4;spacingRight=4;spacingTop=2;spacingBottom=2;';
    cells += `
      <mxCell id="${edgeCellId}-lbl" value="${escapeXml(labelText)}" style="${labelStyle}" vertex="1" connectable="0" parent="${edgeCellId}">
        <mxGeometry x="0" y="0" width="${wrapWidth}" height="${labelBoxHeight}" relative="1" as="geometry">
          <mxPoint x="${offsetX}" as="offset" />
        </mxGeometry>
      </mxCell>`;
  }

  // Numbered callout matching the workflow list — the Azure Architecture
  // Center convention. Drawn as an edge-anchored ellipse so it stays glued to
  // the arrow when the reader reroutes it in Draw.io.
  if (route.stepNumber !== undefined) {
    const badge = 24;
    const stepStyle = 'ellipse;html=1;align=center;verticalAlign=middle;fillColor=#1f2937;strokeColor=#ffffff;strokeWidth=2;fontColor=#ffffff;fontStyle=1;fontSize=12;';
    // Both boxes are centred on their offset point, so clearing the chip needs
    // half of each plus a gap. Without a chip the badge sits on the line.
    const badgeY = label ? Math.round(labelBoxHeight / 2 + badge / 2 + 4) : 0;
    cells += `
      <mxCell id="${edgeCellId}-step" value="${escapeXml(String(route.stepNumber))}" style="${stepStyle}" vertex="1" connectable="0" parent="${edgeCellId}">
        <mxGeometry x="0" y="0" width="${badge}" height="${badge}" relative="1" as="geometry">
          <mxPoint x="${label ? offsetX : 0}" y="${badgeY}" as="offset" />
        </mxGeometry>
      </mxCell>`;
  }

  return cells;
}

// Main export function (async to support icon embedding)
export async function exportToDrawio(
  nodes: Node[],
  edges: Edge[],
  diagramName: string = 'Azure Architecture'
): Promise<string> {
  resetCellIdCounter();

  // Resolve every node into an absolute pixel box via the shared layer so the
  // Draw.io output matches PNG/PPTX/VSDX exactly (fixes 8, 12, 16). Empty bands
  // are closed here too, or the same drawing arrives in Draw.io with fifty
  // inches of blank canvas the other three exporters removed.
  const boxes = compactEmptyGutters(collectExportBoxes(nodes));
  const { groups, services } = partitionBoxes(boxes);
  const routes = buildExportRoutes(edges, boxes);

  const nodeIdToCellId = new Map<string, string>();

  // Groups first so they paint behind the services that sit on top of them.
  const groupCells: string[] = [];
  groups.forEach((group, index) => {
    const cellId = generateCellId();
    nodeIdToCellId.set(group.id, cellId);
    groupCells.push(createGroupCell(group, cellId, index));
  });

  // Services (with async icon embedding) — all on the root layer.
  const serviceCellPromises = services.map(async (service) => {
    const cellId = generateCellId();
    nodeIdToCellId.set(service.id, cellId);
    return createServiceCell(service, cellId);
  });
  const serviceCellResults = await Promise.all(serviceCellPromises);

  const serviceCells: string[] = [];
  const iconCells: string[] = [];
  for (const result of serviceCellResults) {
    serviceCells.push(result.containerCell);
    if (result.iconCell) {
      iconCells.push(result.iconCell);
    }
  }

  // Edges from the shared routes (canonical colour/dash + parallel de-collision).
  const edgeCells: string[] = [];
  for (const route of routes) {
    const edgeCellId = generateCellId();
    const edgeCell = createEdgeCell(route, edgeCellId, nodeIdToCellId);
    if (edgeCell) {
      edgeCells.push(edgeCell);
    }
  }

  // Combine all cells - put icons after services so they appear on top
  const allCells = [...groupCells, ...serviceCells, ...iconCells, ...edgeCells].join('');

  // Page bounds from the real, resolved boxes (fix 12): no more 400x300 guesses.
  const bounds = computeBounds(boxes.values());
  const pageWidth = Math.max(f(bounds.maxX) + 100, 1200);
  const pageHeight = Math.max(f(bounds.maxY) + 100, 800);
  
  // Build the full draw.io XML
  const drawioXml = `<?xml version="1.0" encoding="UTF-8"?>
<mxfile host="app.diagrams.net" modified="${new Date().toISOString()}" agent="Microsoft Product Architecture Diagram Builder" version="1.0" type="device">
  <diagram id="azure-architecture" name="${escapeXml(diagramName)}">
    <mxGraphModel dx="0" dy="0" grid="1" gridSize="10" guides="1" tooltips="1" connect="1" arrows="1" fold="1" page="1" pageScale="1" pageWidth="${pageWidth}" pageHeight="${pageHeight}" math="0" shadow="0">
      <root>
        <mxCell id="0" />
        <mxCell id="1" parent="0" />${allCells}
      </root>
    </mxGraphModel>
  </diagram>
</mxfile>`;
  
  return drawioXml;
}

// Download helper function
export function downloadDrawioFile(xml: string, fileName: string = 'azure-architecture.drawio'): void {
  const blob = new Blob([xml], { type: 'application/xml;charset=utf-8' });
  const url = URL.createObjectURL(blob);
  const link = document.createElement('a');
  link.href = url;
  link.download = fileName;
  document.body.appendChild(link);
  link.click();
  document.body.removeChild(link);
  URL.revokeObjectURL(url);
}

// Combined export and download (async)
export async function exportAndDownloadDrawio(
  nodes: Node[], 
  edges: Edge[], 
  diagramName?: string
): Promise<string> {
  const xml = await exportToDrawio(nodes, edges, diagramName);
  const fileName = generateModelFilename('azure-diagram', 'drawio');
  downloadDrawioFile(xml, fileName);
  return fileName;
}
