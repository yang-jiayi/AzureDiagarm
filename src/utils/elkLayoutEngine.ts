// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

/**
 * ELK Layout Engine - Automatic graph layout using ELK.js
 * Alternative to Dagre for hierarchical layout with compound node support.
 * 
 * Exports the same interface as layoutEngine.ts (Dagre) so the two engines
 * are drop-in replacements for each other.
 */

import ELK, { ElkNode, ElkExtendedEdge } from 'elkjs/lib/elk.bundled.js';
import {
  buildNestedHierarchyLayout,
  layoutNodeDimensions,
} from './layoutHierarchy';
import { wrapPositionedLayout } from './serpentineWrap';
import { fitGroupsToMembers, GROUP_INNER_PAD_PX, GROUP_HEADER_PAD_PX } from './groupFit';

export interface LayoutOptions {
  direction: 'LR' | 'TB' | 'RL' | 'BT';
  nodeSpacing: number;
  rankSpacing: number;
  groupPadding: number;
}

interface LayoutService {
  id: string;
  name: string;
  groupId?: string;
  width?: number;
  height?: number;
  [key: string]: any;
}

interface LayoutConnection {
  from: string;
  to: string;
  [key: string]: any;
}

interface LayoutGroup {
  id: string;
  label: string;
  [key: string]: any;
}

interface PositionedService extends LayoutService {
  position: { x: number; y: number };
}

interface PositionedGroup extends LayoutGroup {
  position: { x: number; y: number };
  width: number;
  height: number;
}

const DEFAULT_OPTIONS: LayoutOptions = {
  direction: 'LR',
  nodeSpacing: 180,
  rankSpacing: 250,
  groupPadding: 100,
};

const NODE_WIDTH = 180;
const NODE_HEIGHT = 100;
const GROUP_GAP = 40;

/** Map our direction codes to ELK's direction enum */
function elkDirection(dir: string): string {
  switch (dir) {
    case 'LR': return 'RIGHT';
    case 'RL': return 'LEFT';
    case 'TB': return 'DOWN';
    case 'BT': return 'UP';
    default: return 'RIGHT';
  }
}

/**
 * Detect and resolve overlapping groups by pushing them apart.
 * Identical to the Dagre engine's implementation for consistency.
 */
function resolveGroupOverlaps(
  groups: PositionedGroup[],
  services: PositionedService[]
): { groups: PositionedGroup[]; services: PositionedService[] } {
  if (groups.length < 2) return { groups, services };

  const resolved = groups.map(g => ({ ...g, position: { ...g.position } }));
  let moved = true;
  let passes = 0;

  while (moved && passes < 10) {
    moved = false;
    passes++;

    for (let i = 0; i < resolved.length; i++) {
      for (let j = i + 1; j < resolved.length; j++) {
        const a = resolved[i];
        const b = resolved[j];

        const overlapX =
          Math.min(a.position.x + a.width, b.position.x + b.width) -
          Math.max(a.position.x, b.position.x);
        const overlapY =
          Math.min(a.position.y + a.height, b.position.y + b.height) -
          Math.max(a.position.y, b.position.y);

        if (overlapX > 0 && overlapY > 0) {
          moved = true;

          if (overlapX < overlapY) {
            const push = (overlapX + GROUP_GAP) / 2;
            const aCx = a.position.x + a.width / 2;
            const bCx = b.position.x + b.width / 2;
            if (aCx <= bCx) {
              a.position.x -= push;
              b.position.x += push;
            } else {
              a.position.x += push;
              b.position.x -= push;
            }
          } else {
            const push = (overlapY + GROUP_GAP) / 2;
            const aCy = a.position.y + a.height / 2;
            const bCy = b.position.y + b.height / 2;
            if (aCy <= bCy) {
              a.position.y -= push;
              b.position.y += push;
            } else {
              a.position.y += push;
              b.position.y -= push;
            }
          }
        }
      }
    }
  }

  if (passes > 1) {
    console.log(`  🔧 Resolved group overlaps in ${passes} passes`);
  }

  return { groups: resolved, services };
}

// Singleton ELK instance (re-used across calls)
const elk = new ELK();

export const ELK_QUALITY_LAYOUT_OPTIONS = {
  'elk.edgeRouting': 'ORTHOGONAL',
  'elk.layered.nodePlacement.strategy': 'NETWORK_SIMPLEX',
  'elk.layered.unnecessaryBendpoints': 'true',
  'elk.spacing.edgeLabelNode': '20',
} as const;

/**
 * Calculate optimal layout for Azure architecture diagram using ELK.js.
 * Signature-compatible with the Dagre `layoutArchitecture`.
 */
export async function layoutArchitecture(
  services: LayoutService[],
  connections: LayoutConnection[],
  groups: LayoutGroup[] = [],
  options: Partial<LayoutOptions> = {}
): Promise<{ services: PositionedService[]; groups: PositionedGroup[] }> {
  const opts = { ...DEFAULT_OPTIONS, ...options };

  console.log('📐 [ELK] Calculating layout for', services.length, 'services and', groups.length, 'groups');

  // Build a lookup: groupId → services[]
  const groupMembers = new Map<string, LayoutService[]>();
  const ungrouped: LayoutService[] = [];

  for (const s of services) {
    if (s.groupId) {
      if (!groupMembers.has(s.groupId)) groupMembers.set(s.groupId, []);
      groupMembers.get(s.groupId)!.push(s);
    } else {
      ungrouped.push(s);
    }
  }

  const usedElkIds = new Set(services.map(service => service.id));
  const groupIdMap = new Map<string, string>();
  for (const group of groups) {
    let elkId = group.id;
    while (usedElkIds.has(elkId)) elkId = `__group__${elkId}`;
    usedElkIds.add(elkId);
    groupIdMap.set(group.id, elkId);
  }

  // Build ELK children: groups become compound nodes containing their members.
  // The container is refitted to its contents after layout, so the padding here
  // only has to keep members clear of the zone's title bar.
  const groupElkNodes: ElkNode[] = groups.map(group => {
    const members = groupMembers.get(group.id) || [];
    return {
      id: groupIdMap.get(group.id) ?? group.id,
      layoutOptions: {
        'elk.padding': `[top=${GROUP_HEADER_PAD_PX}, left=${GROUP_INNER_PAD_PX}, bottom=${GROUP_INNER_PAD_PX}, right=${GROUP_INNER_PAD_PX}]`,
        'elk.spacing.nodeNode': String(Math.max(opts.nodeSpacing * 0.6, 80)),
        'elk.layered.spacing.nodeNodeBetweenLayers': String(Math.max(opts.rankSpacing * 0.6, 120)),
        ...ELK_QUALITY_LAYOUT_OPTIONS,
      },
      children: members.map(m => ({
        id: m.id,
        width: m.width ?? NODE_WIDTH,
        height: m.height ?? NODE_HEIGHT,
      })),
      edges: connections
        .filter(c => {
          const fromIn = members.some(m => m.id === c.from);
          const toIn = members.some(m => m.id === c.to);
          return fromIn && toIn;
        })
        .map(c => ({
          id: `e-${c.from}-${c.to}`,
          sources: [c.from],
          targets: [c.to],
        })),
    };
  });

  const ungroupedElkNodes: ElkNode[] = ungrouped.map(s => ({
    id: s.id,
    width: s.width ?? NODE_WIDTH,
    height: s.height ?? NODE_HEIGHT,
  }));

  // Build top-level edges (cross-group + from/to ungrouped)
  const topEdges: ElkExtendedEdge[] = connections
    .filter(c => {
      // Keep if at least one endpoint is top-level (ungrouped) OR endpoints are in different groups
      const fromGroup = services.find(s => s.id === c.from)?.groupId;
      const toGroup = services.find(s => s.id === c.to)?.groupId;
      if (!fromGroup || !toGroup) return true; // one is ungrouped
      return fromGroup !== toGroup; // cross-group
    })
    .map(c => ({
      id: `e-${c.from}-${c.to}`,
      sources: [c.from],
      targets: [c.to],
    }));

  const root: ElkNode = {
    id: 'root',
    layoutOptions: {
      'elk.algorithm': 'layered',
      'elk.direction': elkDirection(opts.direction),
      'elk.spacing.nodeNode': String(opts.nodeSpacing),
      'elk.layered.spacing.nodeNodeBetweenLayers': String(opts.rankSpacing),
      'elk.spacing.edgeNode': '80',
      'elk.spacing.edgeEdge': '30',
      ...ELK_QUALITY_LAYOUT_OPTIONS,
      'elk.layered.considerModelOrder.strategy': 'NODES_AND_EDGES',
      'elk.hierarchyHandling': 'INCLUDE_CHILDREN',
      'elk.layered.crossingMinimization.strategy': 'LAYER_SWEEP',
    },
    children: [...groupElkNodes, ...ungroupedElkNodes],
    edges: topEdges,
  };

  console.log('  ⚡ Running ELK layout algorithm...');
  const layoutResult = await elk.layout(root);

  // ── Extract positions ────────────────────────────────────────────
  const positionedServices: PositionedService[] = [];
  const positionedGroups: PositionedGroup[] = [];

  // Helper to find a node recursively in the ELK result tree
  function findElkNode(parent: ElkNode, id: string): ElkNode | undefined {
    if (parent.id === id) return parent;
    for (const child of parent.children || []) {
      const found = findElkNode(child, id);
      if (found) return found;
    }
    return undefined;
  }

  // Groups
  for (const group of groups) {
    const elkNode = findElkNode(layoutResult, groupIdMap.get(group.id) ?? group.id);
    if (!elkNode) {
      console.warn(`  ⚠️ [ELK] Group ${group.id} not found in layout result`);
      continue;
    }
    // Guard against NaN values (empty compound nodes can produce NaN)
    const gx = elkNode.x ?? 0;
    const gy = elkNode.y ?? 0;
    const gw = elkNode.width ?? 300;
    const gh = elkNode.height ?? 200;
    if (isNaN(gx) || isNaN(gy)) {
      console.warn(`  ⚠️ [ELK] Group "${group.id}" has NaN position — using fallback`);
    }
    positionedGroups.push({
      ...group,
      position: { x: isNaN(gx) ? 0 : gx, y: isNaN(gy) ? 0 : gy },
      width: isNaN(gw) || gw <= 0 ? 300 : gw,
      height: isNaN(gh) || gh <= 0 ? 200 : gh,
    });
  }

  // Services — absolute positions first
  for (const service of services) {
    const elkNode = findElkNode(layoutResult, service.id);
    if (!elkNode) {
      console.warn(`  ⚠️ [ELK] Service ${service.id} not found in layout result`);
      positionedServices.push({ ...service, position: { x: 0, y: 0 } });
      continue;
    }

    // Guard against NaN
    const sx = elkNode.x ?? 0;
    const sy = elkNode.y ?? 0;
    if (isNaN(sx) || isNaN(sy)) {
      console.warn(`  ⚠️ [ELK] Service "${service.id}" has NaN position — using fallback`);
    }
    positionedServices.push({
      ...service,
      position: { x: isNaN(sx) ? 0 : sx, y: isNaN(sy) ? 0 : sy },
    });
  }

  // ELK reports child coordinates relative to their parent, which is the form
  // the fit expects. Hug the contents before anything downstream measures the
  // drawing.
  const fitted = fitGroupsToMembers(positionedGroups, positionedServices, {
    nodeWidth: NODE_WIDTH,
    nodeHeight: NODE_HEIGHT,
  });

  // Post-process: resolve any overlapping groups
  const { groups: finalGroups } = resolveGroupOverlaps(fitted.groups, fitted.services);

  // Same reason as the Dagre path: a one-service-per-layer flow is a strip, and
  // a strip cannot be exported cleanly onto any page.
  const wrapped = wrapPositionedLayout(fitted.services, finalGroups, {
    direction: opts.direction,
    bandGap: opts.rankSpacing,
    nodeWidth: NODE_WIDTH,
    nodeHeight: NODE_HEIGHT,
  });
  if (wrapped.bands > 1) {
    console.log(`  ✅ [ELK] Wrapped an over-wide layout into ${wrapped.bands} bands`);
  }

  console.log('  ✅ [ELK] Services positioned');
  console.log('  ✅ [ELK] Groups positioned, overlaps resolved');
  console.log('📐 [ELK] Layout complete!');

  return {
    services: wrapped.services,
    groups: wrapped.groups,
  };
}

/**
 * Re-layout existing diagram with new options (ELK version).
 * Signature-compatible with the Dagre `relayoutDiagram`.
 */
export async function relayoutDiagram(
  nodes: any[],
  edges: any[],
  options: Partial<LayoutOptions> = {}
): Promise<any[]> {
  const {
    protectedNodeIds,
    units,
    unitByNodeId,
  } = buildNestedHierarchyLayout(nodes);
  const layoutNodes = nodes.filter(node => !protectedNodeIds.has(node.id));

  const services: LayoutService[] = [
    ...layoutNodes.filter(n => n.type === 'azureNode').map(n => ({
      id: n.id,
      name: n.data.label,
      groupId: n.parentNode,
      ...layoutNodeDimensions(n),
    })),
    ...units.map(unit => ({
      id: unit.surrogateId,
      name: unit.rootId,
      width: unit.width,
      height: unit.height,
    })),
  ];
  const serviceIds = new Set(services.map(service => service.id));

  const connections = edges
    .map(edge => ({
      from: unitByNodeId.get(edge.source)?.surrogateId ?? edge.source,
      to: unitByNodeId.get(edge.target)?.surrogateId ?? edge.target,
    }))
    .filter(connection => (
      connection.from !== connection.to
      && serviceIds.has(connection.from)
      && serviceIds.has(connection.to)
    ));

  const groups = layoutNodes
    .filter(n => n.type === 'groupNode')
    .map(n => ({
      id: n.id,
      label: n.data.label,
    }));

  const { services: positioned, groups: positionedGroups } = await layoutArchitecture(
    services,
    connections,
    groups,
    options
  );

  const updatedNodes = nodes.map(node => {
    const hierarchyUnit = units.find(unit => unit.rootId === node.id);
    if (hierarchyUnit) {
      const pos = positioned.find(service => service.id === hierarchyUnit.surrogateId);
      return pos
        ? {
            ...node,
            position: {
              x: pos.position.x - hierarchyUnit.offsetX,
              y: pos.position.y - hierarchyUnit.offsetY,
            },
          }
        : node;
    }
    if (protectedNodeIds.has(node.id)) return node;
    if (node.type === 'azureNode') {
      const pos = positioned.find(s => s.id === node.id);
      if (pos) {
        return { ...node, position: pos.position };
      }
    } else if (node.type === 'groupNode') {
      const pos = positionedGroups.find(g => g.id === node.id);
      if (pos) {
        return {
          ...node,
          position: pos.position,
          style: {
            ...node.style,
            width: pos.width,
            height: pos.height,
          },
        };
      }
    }
    return node;
  });

  return updatedNodes;
}

/**
 * Calculate layout direction based on architecture type.
 */
export function suggestLayoutDirection(services: LayoutService[]): 'LR' | 'TB' {
  if (services.length > 8) return 'LR';
  return 'LR';
}
