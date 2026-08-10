// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

import type { Edge, Node } from 'reactflow';
import { fitGroupToContent } from './groupUtils';
import { buildAbsolutePositionMap } from './preserveManualLayout';
import { resolveServiceIconLoose } from './serviceIconFuzzy';

export type DiagramQualityCategory =
  | 'crossing'
  | 'overlap'
  | 'orphan'
  | 'label'
  | 'density'
  | 'group-padding'
  | 'contrast'
  | 'unrecognized-service'
  | 'generic-edge'
  | 'dangling-edge'
  | 'self-loop'
  | 'duplicate-edge'
  | 'under-specified'
  | 'over-crowded'
  | 'empty-group';

export type DiagramQualitySeverity = 'high' | 'medium' | 'low';
export type DiagramQualityFixKind =
  | 'layout'
  | 'expand-label'
  | 'fit-group'
  | 'reset-contrast';

export interface DiagramQualityFinding {
  id: string;
  category: DiagramQualityCategory;
  severity: DiagramQualitySeverity;
  title: string;
  detail: string;
  nodeIds: string[];
  edgeIds: string[];
  fixKind?: DiagramQualityFixKind;
}

export interface DiagramQualityReport {
  findings: DiagramQualityFinding[];
  score: number;
  counts: Record<DiagramQualitySeverity, number>;
}

export interface DiagramQualityFixResult {
  nodes: Node[];
  edges: Edge[];
  fixedFindingIds: string[];
  requiresLayout: boolean;
}

interface Rect {
  left: number;
  top: number;
  right: number;
  bottom: number;
  width: number;
  height: number;
}

interface Point {
  x: number;
  y: number;
}

const MAX_FINDINGS = 100;
const GROUP_PADDING = 24;
const GROUP_HEADER = 54;

function isRecord(value: unknown): value is Record<string, unknown> {
  return Boolean(value) && typeof value === 'object' && !Array.isArray(value);
}

function numericDimension(value: unknown): number | undefined {
  if (typeof value === 'number' && Number.isFinite(value) && value > 0) return value;
  if (typeof value === 'string') {
    const parsed = Number.parseFloat(value);
    if (Number.isFinite(parsed) && parsed > 0) return parsed;
  }
  return undefined;
}

function dimensions(node: Node): { width: number; height: number } {
  const style = isRecord(node.style) ? node.style : {};
  return {
    width: numericDimension(node.width)
      ?? numericDimension(style.width)
      ?? (node.type === 'groupNode' ? 400 : 160),
    height: numericDimension(node.height)
      ?? numericDimension(style.height)
      ?? (node.type === 'groupNode' ? 300 : 128),
  };
}

function nodeLabel(node: Node): string {
  const data = isRecord(node.data) ? node.data : {};
  const label = data.label ?? data.serviceName;
  return typeof label === 'string' && label.trim() ? label.trim() : node.id;
}

/** Service identity used for icon resolution: prefers serviceName, then label. */
function serviceIdentityName(node: Node): string {
  const data = isRecord(node.data) ? node.data : {};
  const identity = data.serviceName ?? data.label;
  return typeof identity === 'string' ? identity.trim() : '';
}

function iconPathValue(node: Node): string {
  const data = isRecord(node.data) ? node.data : {};
  return typeof data.iconPath === 'string' ? data.iconPath.trim() : '';
}

/** Resolve an edge's label from either `edge.label` or `edge.data.label`. */
function edgeLabelText(edge: Edge): string {
  if (typeof edge.label === 'string' && edge.label.trim()) return edge.label.trim();
  const data = isRecord(edge.data) ? edge.data : {};
  return typeof data.label === 'string' ? data.label.trim() : '';
}

// Labels that are technically present but carry no real meaning — every
// generator prompt demands a specific payload/protocol label instead.
const GENERIC_EDGE_LABELS = new Set([
  '',
  'connects to',
  'connect',
  'data',
  'request',
  'response',
  'uses',
  'link',
  'flow',
  'api',
]);

/** True when `ancestorId` is a (transitive) parent of `node`. */
function isAncestorGroup(
  ancestorId: string,
  node: Node,
  nodesById: ReadonlyMap<string, Node>,
): boolean {
  const seen = new Set<string>();
  let current: Node | undefined = node;
  while (current?.parentNode && !seen.has(current.id)) {
    if (current.parentNode === ancestorId) return true;
    seen.add(current.id);
    current = nodesById.get(current.parentNode);
  }
  return false;
}

function rectForNode(node: Node, absolutePositions: ReadonlyMap<string, Point>): Rect {
  const position = absolutePositions.get(node.id) || node.position;
  const size = dimensions(node);
  return {
    left: position.x,
    top: position.y,
    right: position.x + size.width,
    bottom: position.y + size.height,
    width: size.width,
    height: size.height,
  };
}

function intersectionArea(left: Rect, right: Rect): number {
  const width = Math.min(left.right, right.right) - Math.max(left.left, right.left);
  const height = Math.min(left.bottom, right.bottom) - Math.max(left.top, right.top);
  return width > 0 && height > 0 ? width * height : 0;
}

function center(rect: Rect): Point {
  return {
    x: rect.left + rect.width / 2,
    y: rect.top + rect.height / 2,
  };
}

function orientation(a: Point, b: Point, c: Point): number {
  return (b.y - a.y) * (c.x - b.x) - (b.x - a.x) * (c.y - b.y);
}

function segmentsCross(a: Point, b: Point, c: Point, d: Point): boolean {
  const first = orientation(a, b, c);
  const second = orientation(a, b, d);
  const third = orientation(c, d, a);
  const fourth = orientation(c, d, b);
  return (
    ((first > 0 && second < 0) || (first < 0 && second > 0))
    && ((third > 0 && fourth < 0) || (third < 0 && fourth > 0))
  );
}

type Rgb = [number, number, number];

function parseColor(value: unknown): Rgb | null {
  if (typeof value !== 'string') return null;
  const color = value.trim().toLowerCase();
  const hex = color.match(/^#([0-9a-f]{3}|[0-9a-f]{6})$/i);
  if (hex) {
    const digits = hex[1].length === 3
      ? hex[1].split('').map(character => character.repeat(2)).join('')
      : hex[1];
    return [
      Number.parseInt(digits.slice(0, 2), 16),
      Number.parseInt(digits.slice(2, 4), 16),
      Number.parseInt(digits.slice(4, 6), 16),
    ];
  }
  const rgb = color.match(/^rgba?\(\s*(\d+)[,\s]+(\d+)[,\s]+(\d+)/);
  if (!rgb) return null;
  return [Number(rgb[1]), Number(rgb[2]), Number(rgb[3])];
}

function luminance([red, green, blue]: Rgb): number {
  const channels = [red, green, blue].map((channel) => {
    const normalized = channel / 255;
    return normalized <= 0.03928
      ? normalized / 12.92
      : ((normalized + 0.055) / 1.055) ** 2.4;
  });
  return channels[0] * 0.2126 + channels[1] * 0.7152 + channels[2] * 0.0722;
}

function contrastRatio(foreground: Rgb, background: Rgb): number {
  const first = luminance(foreground);
  const second = luminance(background);
  const lighter = Math.max(first, second);
  const darker = Math.min(first, second);
  return (lighter + 0.05) / (darker + 0.05);
}

function pushFinding(
  findings: DiagramQualityFinding[],
  finding: DiagramQualityFinding,
): void {
  if (findings.length < MAX_FINDINGS) findings.push(finding);
}

export function analyzeDiagramQuality(nodes: Node[], edges: Edge[]): DiagramQualityReport {
  const findings: DiagramQualityFinding[] = [];
  const absolutePositions = buildAbsolutePositionMap(nodes);
  const nodesById = new Map(nodes.map(node => [node.id, node]));
  const rectangles = new Map(nodes.map(node => [
    node.id,
    rectForNode(node, absolutePositions),
  ]));

  for (let leftIndex = 0; leftIndex < nodes.length; leftIndex += 1) {
    const leftNode = nodes[leftIndex];
    for (let rightIndex = leftIndex + 1; rightIndex < nodes.length; rightIndex += 1) {
      const rightNode = nodes[rightIndex];
      if ((leftNode.parentNode || '') !== (rightNode.parentNode || '')) continue;
      const leftRect = rectangles.get(leftNode.id);
      const rightRect = rectangles.get(rightNode.id);
      if (!leftRect || !rightRect) continue;
      const overlap = intersectionArea(leftRect, rightRect);
      if (overlap < 400) continue;
      pushFinding(findings, {
        id: `overlap:${leftNode.id}:${rightNode.id}`,
        category: 'overlap',
        severity: 'high',
        title: 'Overlapping diagram elements',
        detail: `${nodeLabel(leftNode)} overlaps ${nodeLabel(rightNode)}.`,
        nodeIds: [leftNode.id, rightNode.id],
        edgeIds: [],
        fixKind: 'layout',
      });
    }
  }

  for (let leftIndex = 0; leftIndex < edges.length; leftIndex += 1) {
    const leftEdge = edges[leftIndex];
    const leftSource = rectangles.get(leftEdge.source);
    const leftTarget = rectangles.get(leftEdge.target);
    if (!leftSource || !leftTarget) continue;
    for (let rightIndex = leftIndex + 1; rightIndex < edges.length; rightIndex += 1) {
      const rightEdge = edges[rightIndex];
      if (
        leftEdge.source === rightEdge.source
        || leftEdge.source === rightEdge.target
        || leftEdge.target === rightEdge.source
        || leftEdge.target === rightEdge.target
      ) continue;
      const rightSource = rectangles.get(rightEdge.source);
      const rightTarget = rectangles.get(rightEdge.target);
      if (!rightSource || !rightTarget) continue;
      if (!segmentsCross(
        center(leftSource),
        center(leftTarget),
        center(rightSource),
        center(rightTarget),
      )) continue;
      pushFinding(findings, {
        id: `crossing:${leftEdge.id}:${rightEdge.id}`,
        category: 'crossing',
        severity: 'medium',
        title: 'Crossing connections',
        detail: 'Two connections cross and may be difficult to follow.',
        nodeIds: [],
        edgeIds: [leftEdge.id, rightEdge.id],
        fixKind: 'layout',
      });
    }
  }

  const serviceNodes = nodes.filter(node => node.type !== 'groupNode');
  const connectedNodeIds = new Set(edges.flatMap(edge => [edge.source, edge.target]));
  if (serviceNodes.length > 1) {
    for (const node of serviceNodes) {
      if (connectedNodeIds.has(node.id)) continue;
      pushFinding(findings, {
        id: `orphan:${node.id}`,
        category: 'orphan',
        severity: 'medium',
        title: 'Unconnected service',
        detail: `${nodeLabel(node)} has no incoming or outgoing connection.`,
        nodeIds: [node.id],
        edgeIds: [],
      });
    }
  }

  for (const node of serviceNodes) {
    const label = nodeLabel(node);
    const data = isRecord(node.data) ? node.data : {};
    const configuredWidth = numericDimension(data.labelMaxWidth) ?? 140;
    const longestToken = Math.max(...label.split(/\s+/).map(part => part.length), 0);
    if (label.length <= 34 && longestToken <= 20) continue;
    const recommendedWidth = Math.min(260, Math.max(160, Math.ceil(label.length * 4.6)));
    if (configuredWidth >= recommendedWidth) continue;
    pushFinding(findings, {
      id: `label:${node.id}`,
      category: 'label',
      severity: 'low',
      title: 'Crowded service label',
      detail: `${label} may wrap into a dense or clipped label block.`,
      nodeIds: [node.id],
      edgeIds: [],
      fixKind: 'expand-label',
    });
  }

  if (serviceNodes.length >= 12) {
    const serviceRects = serviceNodes
      .map(node => rectangles.get(node.id))
      .filter((rect): rect is Rect => Boolean(rect));
    if (serviceRects.length > 0) {
      const left = Math.min(...serviceRects.map(rect => rect.left));
      const top = Math.min(...serviceRects.map(rect => rect.top));
      const right = Math.max(...serviceRects.map(rect => rect.right));
      const bottom = Math.max(...serviceRects.map(rect => rect.bottom));
      const areaPerService = ((right - left) * (bottom - top)) / serviceRects.length;
      if (areaPerService < 22_000) {
        pushFinding(findings, {
          id: 'density:canvas',
          category: 'density',
          severity: 'medium',
          title: 'Excessive diagram density',
          detail: `${serviceNodes.length} services are packed into a small canvas area.`,
          nodeIds: serviceNodes.map(node => node.id),
          edgeIds: [],
          fixKind: 'layout',
        });
      }
    }
  }

  for (const group of nodes.filter(node => node.type === 'groupNode')) {
    const groupSize = dimensions(group);
    const children = nodes.filter(node => node.parentNode === group.id);
    if (children.length === 0) {
      pushFinding(findings, {
        id: `empty-group:${group.id}`,
        category: 'empty-group',
        severity: 'low',
        title: 'Empty group',
        detail: `${nodeLabel(group)} contains no services and renders as an empty labelled box.`,
        nodeIds: [group.id],
        edgeIds: [],
      });
      continue;
    }
    const crowded = children.some((child) => {
      const childSize = dimensions(child);
      return (
        child.position.x < GROUP_PADDING
        || child.position.y < GROUP_HEADER
        || child.position.x + childSize.width > groupSize.width - GROUP_PADDING
        || child.position.y + childSize.height > groupSize.height - GROUP_PADDING
      );
    });
    if (!crowded) continue;
    pushFinding(findings, {
      id: `group-padding:${group.id}`,
      category: 'group-padding',
      severity: 'medium',
      title: 'Insufficient group padding',
      detail: `${nodeLabel(group)} has content too close to or outside its boundary.`,
      nodeIds: [group.id, ...children.map(child => child.id)],
      edgeIds: [],
      fixKind: 'fit-group',
    });
  }

  for (const node of nodes) {
    const style = isRecord(node.style) ? node.style : {};
    const foreground = parseColor(style.color);
    const background = parseColor(style.backgroundColor ?? style.background);
    if (!foreground || !background || contrastRatio(foreground, background) >= 4.5) continue;
    pushFinding(findings, {
      id: `contrast:node:${node.id}`,
      category: 'contrast',
      severity: 'high',
      title: 'Low node text contrast',
      detail: `${nodeLabel(node)} uses custom colors below the WCAG AA contrast target.`,
      nodeIds: [node.id],
      edgeIds: [],
      fixKind: 'reset-contrast',
    });
  }

  for (const edge of edges) {
    const labelStyle = isRecord(edge.labelStyle) ? edge.labelStyle : {};
    const backgroundStyle = isRecord(edge.labelBgStyle) ? edge.labelBgStyle : {};
    const foreground = parseColor(labelStyle.fill);
    const background = parseColor(backgroundStyle.fill);
    if (!foreground || !background || contrastRatio(foreground, background) >= 4.5) continue;
    pushFinding(findings, {
      id: `contrast:edge:${edge.id}`,
      category: 'contrast',
      severity: 'high',
      title: 'Low connection-label contrast',
      detail: 'A connection label uses custom colors below the WCAG AA contrast target.',
      nodeIds: [],
      edgeIds: [edge.id],
      fixKind: 'reset-contrast',
    });
  }

  // --- Unrecognized service icons (issue 6). The most visible AI defect:
  // a node whose name cannot be resolved to an Azure icon renders icon-less.
  for (const node of serviceNodes) {
    const identity = serviceIdentityName(node);
    const resolved = identity ? resolveServiceIconLoose(identity) : null;
    if (resolved && iconPathValue(node)) continue;
    const shownName = identity || nodeLabel(node);
    pushFinding(findings, {
      id: `unrecognized-service:${node.id}`,
      category: 'unrecognized-service',
      severity: 'high',
      title: 'Unrecognized service',
      detail: `"${shownName}" does not resolve to an Azure icon — rename it to a known service so it renders correctly.`,
      nodeIds: [node.id],
      edgeIds: [],
    });
  }

  // --- Generic / missing edge labels (issue 7). Every prompt demands a
  // specific payload or protocol label on each connection.
  for (const edge of edges) {
    if (!GENERIC_EDGE_LABELS.has(edgeLabelText(edge).toLowerCase())) continue;
    pushFinding(findings, {
      id: `generic-edge:${edge.id}`,
      category: 'generic-edge',
      severity: 'medium',
      title: 'Generic connection label',
      detail: 'A connection has an empty or generic label; name the payload or protocol it carries.',
      nodeIds: [],
      edgeIds: [edge.id],
    });
  }

  // --- Extended overlap (issue 10). The same-parent overlap loop above misses
  // (a) a service overlapping a group it does not belong to and (b) two group
  // rectangles that do not share a parent. Distinct id namespaces avoid clashes.
  const groupNodes = nodes.filter(node => node.type === 'groupNode');
  for (const service of serviceNodes) {
    const serviceRect = rectangles.get(service.id);
    if (!serviceRect) continue;
    for (const group of groupNodes) {
      if ((service.parentNode || '') === (group.parentNode || '')) continue;
      if (isAncestorGroup(group.id, service, nodesById)) continue;
      const groupRect = rectangles.get(group.id);
      if (!groupRect) continue;
      if (intersectionArea(serviceRect, groupRect) <= 400) continue;
      pushFinding(findings, {
        id: `overlap:node-group:${service.id}:${group.id}`,
        category: 'overlap',
        severity: 'medium',
        title: 'Service overlaps unrelated group',
        detail: `${nodeLabel(service)} overlaps ${nodeLabel(group)}, a group it does not belong to.`,
        nodeIds: [service.id, group.id],
        edgeIds: [],
        fixKind: 'layout',
      });
    }
  }
  for (let leftIndex = 0; leftIndex < groupNodes.length; leftIndex += 1) {
    const leftGroup = groupNodes[leftIndex];
    const leftRect = rectangles.get(leftGroup.id);
    if (!leftRect) continue;
    for (let rightIndex = leftIndex + 1; rightIndex < groupNodes.length; rightIndex += 1) {
      const rightGroup = groupNodes[rightIndex];
      if ((leftGroup.parentNode || '') === (rightGroup.parentNode || '')) continue;
      if (
        isAncestorGroup(leftGroup.id, rightGroup, nodesById)
        || isAncestorGroup(rightGroup.id, leftGroup, nodesById)
      ) continue;
      const rightRect = rectangles.get(rightGroup.id);
      if (!rightRect) continue;
      if (intersectionArea(leftRect, rightRect) <= 400) continue;
      pushFinding(findings, {
        id: `overlap:groups:${leftGroup.id}:${rightGroup.id}`,
        category: 'overlap',
        severity: 'medium',
        title: 'Overlapping groups',
        detail: `${nodeLabel(leftGroup)} overlaps ${nodeLabel(rightGroup)}.`,
        nodeIds: [leftGroup.id, rightGroup.id],
        edgeIds: [],
        fixKind: 'layout',
      });
    }
  }

  // --- Edge integrity (issue 11): dangling endpoints, self-loops, duplicates.
  const seenEdgePairs = new Set<string>();
  for (const edge of edges) {
    const sourceMissing = !nodesById.has(edge.source);
    const targetMissing = !nodesById.has(edge.target);
    if (sourceMissing || targetMissing) {
      pushFinding(findings, {
        id: `dangling-edge:${edge.id}`,
        category: 'dangling-edge',
        severity: 'high',
        title: 'Dangling connection',
        detail: `A connection references a missing ${sourceMissing ? 'source' : 'target'} node and cannot render.`,
        nodeIds: [],
        edgeIds: [edge.id],
      });
    }
    if (edge.source === edge.target) {
      const loopNode = nodesById.get(edge.source);
      pushFinding(findings, {
        id: `self-loop:${edge.id}`,
        category: 'self-loop',
        severity: 'medium',
        title: 'Self-referencing connection',
        detail: `${loopNode ? nodeLabel(loopNode) : edge.source} connects to itself.`,
        nodeIds: loopNode ? [edge.source] : [],
        edgeIds: [edge.id],
      });
    }
    const pairKey = `${edge.source}->${edge.target}`;
    if (seenEdgePairs.has(pairKey)) {
      pushFinding(findings, {
        id: `duplicate-edge:${edge.id}`,
        category: 'duplicate-edge',
        severity: 'low',
        title: 'Duplicate connection',
        detail: 'The same source-to-target connection is drawn more than once.',
        nodeIds: [],
        edgeIds: [edge.id],
      });
    } else {
      seenEdgePairs.add(pairKey);
    }
  }

  // --- Diagram scale (issue 13), independent of the density heuristic above.
  if (serviceNodes.length > 0 && serviceNodes.length < 3) {
    pushFinding(findings, {
      id: 'under-specified',
      category: 'under-specified',
      severity: 'low',
      title: 'Under-specified architecture',
      detail: `Only ${serviceNodes.length} service${serviceNodes.length === 1 ? '' : 's'} — a production architecture usually needs more detail.`,
      nodeIds: serviceNodes.map(node => node.id),
      edgeIds: [],
    });
  }
  if (serviceNodes.length > 25) {
    pushFinding(findings, {
      id: 'over-crowded',
      category: 'over-crowded',
      severity: 'medium',
      title: 'Over-crowded diagram',
      detail: `${serviceNodes.length} services make the diagram hard to read — split it into focused views.`,
      nodeIds: [],
      edgeIds: [],
    });
  }

  const counts = {
    high: findings.filter(finding => finding.severity === 'high').length,
    medium: findings.filter(finding => finding.severity === 'medium').length,
    low: findings.filter(finding => finding.severity === 'low').length,
  };
  const score = Math.max(0, 100 - counts.high * 12 - counts.medium * 7 - counts.low * 3);
  return { findings, score, counts };
}

export function applyDiagramQualityFixes(
  nodes: Node[],
  edges: Edge[],
  findings: DiagramQualityFinding[],
  selectedFindingIds: Iterable<string>,
): DiagramQualityFixResult {
  const selected = new Set(selectedFindingIds);
  const selectedFindings = findings.filter(finding => (
    selected.has(finding.id) && finding.fixKind
  ));
  let nextNodes: Node[] = nodes.map(node => ({
    ...node,
    position: { ...node.position },
    style: node.style ? { ...node.style } : undefined,
    data: isRecord(node.data) ? { ...node.data } : node.data,
  }));
  let nextEdges: Edge[] = edges.map(edge => ({
    ...edge,
    style: edge.style ? { ...edge.style } : undefined,
    labelStyle: edge.labelStyle ? { ...edge.labelStyle } : undefined,
    labelBgStyle: edge.labelBgStyle ? { ...edge.labelBgStyle } : undefined,
    data: isRecord(edge.data) ? { ...edge.data } : edge.data,
  }));
  let requiresLayout = false;
  const fixedFindingIds: string[] = [];

  for (const finding of selectedFindings) {
    switch (finding.fixKind) {
      case 'layout':
        requiresLayout = true;
        fixedFindingIds.push(finding.id);
        break;
      case 'expand-label':
        nextNodes = nextNodes.map((node) => {
          if (!finding.nodeIds.includes(node.id)) return node;
          const label = nodeLabel(node);
          const labelMaxWidth = Math.min(260, Math.max(160, Math.ceil(label.length * 4.6)));
          return {
            ...node,
            style: {
              ...node.style,
              width: Math.max(dimensions(node).width, labelMaxWidth + 24),
            },
            data: {
              ...node.data,
              labelMaxWidth,
            },
          };
        });
        fixedFindingIds.push(finding.id);
        break;
      case 'fit-group': {
        const groupId = finding.nodeIds[0];
        const fitted = fitGroupToContent(nextNodes, groupId);
        if (fitted) {
          nextNodes = fitted;
          fixedFindingIds.push(finding.id);
        }
        break;
      }
      case 'reset-contrast':
        nextNodes = nextNodes.map((node) => {
          if (!finding.nodeIds.includes(node.id)) return node;
          const style = { ...node.style } as Record<string, unknown>;
          delete style.color;
          delete style.background;
          delete style.backgroundColor;
          return { ...node, style };
        });
        nextEdges = nextEdges.map((edge) => {
          if (!finding.edgeIds.includes(edge.id)) return edge;
          return {
            ...edge,
            labelStyle: {
              ...edge.labelStyle,
              fill: '#1f2937',
            },
            labelBgStyle: {
              ...edge.labelBgStyle,
              fill: '#ffffff',
              fillOpacity: 0.96,
            },
          };
        });
        fixedFindingIds.push(finding.id);
        break;
    }
  }

  return {
    nodes: nextNodes,
    edges: nextEdges,
    fixedFindingIds: [...new Set(fixedFindingIds)],
    requiresLayout,
  };
}
