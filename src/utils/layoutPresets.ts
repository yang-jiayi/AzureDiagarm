// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

import type { Edge, Node } from 'reactflow';
import { collectNestedHierarchyNodeIds } from './layoutHierarchy';
import { buildAbsolutePositionMap, selectHorizontalConnectionHandles } from './preserveManualLayout';

export type LayoutEngineType = 'dagre' | 'elk';

export type LayoutPreset = 'flow-lr' | 'flow-tb' | 'swimlanes' | 'radial';
export type LayoutSpacing = 'compact' | 'comfortable';
export type LayoutEdgeStyle = 'straight' | 'smooth' | 'orthogonal';

export interface ApplyLayoutOptions {
  preset: LayoutPreset;
  spacing: LayoutSpacing;
  edgeStyle: LayoutEdgeStyle;
  emphasizePrimaryPath: boolean;
  selectedNodeId?: string;
  layoutEngine?: LayoutEngineType;
}

const NODE_WIDTH = 180;
const NODE_HEIGHT = 100;

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
      ?? (node.type === 'groupNode' ? 420 : NODE_WIDTH),
    height: numericDimension(node.height)
      ?? numericDimension(style?.height)
      ?? (node.type === 'groupNode' ? 260 : NODE_HEIGHT),
  };
}

function getSpacing(spacing: LayoutSpacing) {
  if (spacing === 'compact') {
    return {
      nodeSpacing: 110,
      rankSpacing: 140,
      groupPadding: 60,
      laneGap: 80,
      radialBaseRadius: 170,
      radialRingStep: 160,
    };
  }

  return {
    nodeSpacing: 150,
    rankSpacing: 200,
    groupPadding: 80,
    laneGap: 120,
    radialBaseRadius: 220,
    radialRingStep: 200,
  };
}

function withEdgeStyle(edges: Edge[], edgeStyle: LayoutEdgeStyle): Edge[] {
  return edges.map((e) => ({
    ...e,
    data: { ...(e.data ?? {}), pathStyle: edgeStyle },
  }));
}

type Dir = 'forward' | 'reverse' | 'bidirectional';

function normalizeDirectedAdjacency(nodes: Node[], edges: Edge[]) {
  const azureIds = new Set(nodes.filter((n) => n.type === 'azureNode').map((n) => n.id));

  const out = new Map<string, Set<string>>();
  const indeg = new Map<string, number>();
  const outdeg = new Map<string, number>();

  for (const id of azureIds) {
    out.set(id, new Set());
    indeg.set(id, 0);
    outdeg.set(id, 0);
  }

  const addEdge = (a: string, b: string) => {
    if (!azureIds.has(a) || !azureIds.has(b)) return;
    const s = out.get(a);
    if (!s) return;
    if (!s.has(b)) {
      s.add(b);
      outdeg.set(a, (outdeg.get(a) ?? 0) + 1);
      indeg.set(b, (indeg.get(b) ?? 0) + 1);
    }
  };

  for (const e of edges) {
    const dir = ((e.data as any)?.direction ?? 'forward') as Dir;
    if (dir === 'reverse') {
      addEdge(e.target, e.source);
    } else if (dir === 'bidirectional') {
      addEdge(e.source, e.target);
      addEdge(e.target, e.source);
    } else {
      addEdge(e.source, e.target);
    }
  }

  return { azureIds, out, indeg, outdeg };
}

function computePrimaryChain(nodes: Node[], edges: Edge[]): string[] {
  const { azureIds, out, indeg, outdeg } = normalizeDirectedAdjacency(nodes, edges);

  const entries = [...azureIds].filter((id) => (indeg.get(id) ?? 0) === 0);
  const startCandidates = entries.length > 0 ? entries : [...azureIds];

  let start: string | undefined;
  for (const id of startCandidates) {
    const score = outdeg.get(id) ?? 0;
    if (!start || score > (outdeg.get(start) ?? 0)) start = id;
  }

  if (!start) return [];

  const chain: string[] = [start];
  const visited = new Set(chain);

  while (true) {
    const cur = chain[chain.length - 1];
    const nextCandidates = [...(out.get(cur) ?? new Set())].filter((n) => !visited.has(n));
    if (nextCandidates.length === 0) break;

    nextCandidates.sort((a, b) => {
      const da = outdeg.get(a) ?? 0;
      const db = outdeg.get(b) ?? 0;
      return db - da;
    });

    const next = nextCandidates[0];
    chain.push(next);
    visited.add(next);

    if (chain.length > 64) break; // guardrail
  }

  return chain;
}

function clamp(v: number, min: number, max: number) {
  return Math.max(min, Math.min(max, v));
}

export function straightenPrimaryPath(nodes: Node[], edges: Edge[], direction: 'LR' | 'TB') {
  const chain = computePrimaryChain(nodes, edges);
  if (chain.length < 3) return { nodes, edges };

  const byId = new Map(nodes.map((n) => [n.id, n] as const));
  const chainNodes = chain.map((id) => byId.get(id)).filter(Boolean) as Node[];

  if (chainNodes.length < 3) return { nodes, edges };

  const axis = direction === 'LR' ? 'y' : 'x';
  const absolutePositions = buildAbsolutePositionMap(nodes);

  // The serpentine wrap folds a long flow into bands and reverses every other
  // one, so the chain doubles back on itself. Snapping the whole chain to one
  // median would stack every band on top of the others. A band is exactly a
  // maximal run of the chain that travels in one direction, so split on the
  // direction changes and straighten within each run. An unwrapped chain never
  // reverses, giving one run and the original behaviour.
  const majorAxis = direction === 'LR' ? 'x' : 'y';
  const majorOf = (node: Node): number =>
    absolutePositions.get(node.id)?.[majorAxis] ?? node.position[majorAxis];
  const minorOf = (node: Node): number =>
    absolutePositions.get(node.id)?.[axis] ?? node.position[axis];

  // What a seam looks like depends on the direction and on how the bands
  // happen to line up, so no threshold on either axis identifies one. Reading
  // the minor axis alone misses every TB seam, where the band gap (a tile
  // width plus the rank spacing, about 290px) is smaller than a node is wide.
  // Reading the major axis alone misses an LR seam between bands of different
  // widths, where the hop is (extent_next - extent_prev)/2 and can be large
  // and forward-pointing.
  //
  // What is always true is that the chain reverses at a seam, and that the
  // seam hop is whichever of the two steps around that reversal crossed the
  // band gap. So find the reversal and attribute it, comparing those two steps
  // against each other rather than against any constant.
  const steps = chainNodes.slice(1).map((node, i) => ({
    step: majorOf(node) - majorOf(chainNodes[i]),
    drift: minorOf(node) - minorOf(chainNodes[i]),
  }));

  const seamAfter = new Set<number>();
  let heading = 0;
  for (let i = 0; i < steps.length; i += 1) {
    const stepSign = Math.sign(steps[i].step);
    if (stepSign === 0) continue;
    if (heading === 0) {
      heading = stepSign;
      continue;
    }
    if (stepSign === heading) continue;
    // The reversal shows up at step i, but the band was left at whichever of
    // steps i-1 and i actually travelled across the minor axis.
    if (i > 0 && Math.abs(steps[i - 1].drift) > Math.abs(steps[i].drift)) {
      seamAfter.add(i - 1);
      heading = stepSign;
    } else {
      seamAfter.add(i);
      heading = 0;
    }
  }

  const runs: Node[][] = [[chainNodes[0]]];
  for (let i = 1; i < chainNodes.length; i += 1) {
    if (seamAfter.has(i - 1)) runs.push([]);
    runs[runs.length - 1].push(chainNodes[i]);
  }

  const medianById = new Map<string, number>();
  for (const run of runs) {
    const sorted = run.map(minorOf).sort((a, b) => a - b);
    const runMedian = sorted[Math.floor(sorted.length / 2)];
    for (const node of run) medianById.set(node.id, runMedian);
  }

  const parentById = new Map<string, Node>();
  for (const n of nodes) {
    if (n.type === 'groupNode') parentById.set(n.id, n);
  }

  const updatedNodes = nodes.map((n) => {
    if (!chain.includes(n.id) || n.type !== 'azureNode') return n;

    const median = medianById.get(n.id);
    if (median === undefined) return n;

    // Keep node inside group bounds if it's a child.
    const parentId = (n as any).parentNode as string | undefined;
    if (parentId) {
      const parent = parentById.get(parentId);
      const parentPosition = absolutePositions.get(parentId);
      const { width, height } = parent ? nodeDimensions(parent) : { width: 0, height: 0 };
      if (!parentPosition) return n;

      if (axis === 'y') {
        const relativeMedian = median - parentPosition.y;
        const newY = clamp(relativeMedian, 0, Math.max(0, height - NODE_HEIGHT));
        return { ...n, position: { ...n.position, y: newY } };
      }
      if (axis === 'x') {
        const relativeMedian = median - parentPosition.x;
        const newX = clamp(relativeMedian, 0, Math.max(0, width - NODE_WIDTH));
        return { ...n, position: { ...n.position, x: newX } };
      }
    }

    return {
      ...n,
      position: {
        ...n.position,
        [axis]: median,
      },
    };
  });

  // Mark chain edges as primary for optional styling.
  const chainPairs = new Set<string>();
  for (let i = 0; i < chain.length - 1; i++) {
    chainPairs.add(`${chain[i]}->${chain[i + 1]}`);
  }

  const updatedEdges = edges.map((e) => {
    const isPrimary = chainPairs.has(`${e.source}->${e.target}`) || chainPairs.has(`${e.target}->${e.source}`);
    return { ...e, data: { ...(e.data ?? {}), primaryPath: isPrimary } };
  });

  return { nodes: updatedNodes, edges: updatedEdges };
}

async function applySwimlanesByGroup(nodes: Node[], edges: Edge[], spacing: ReturnType<typeof getSpacing>, doRelayout: RelayoutFn) {
  const groupNodes = nodes.filter((n) => n.type === 'groupNode');
  const serviceNodes = nodes.filter((n) => n.type === 'azureNode');
  const protectedNestedNodeIds = collectNestedHierarchyNodeIds(nodes);

  // Map groupId -> members
  const membersByGroup = new Map<string, Node[]>();
  for (const g of groupNodes) membersByGroup.set(g.id, []);
  const ungrouped: Node[] = [];

  for (const n of serviceNodes) {
    const parentId = (n as any).parentNode as string | undefined;
    if (parentId && membersByGroup.has(parentId)) {
      membersByGroup.get(parentId)!.push(n);
    } else {
      ungrouped.push(n);
    }
  }

  const groupOrder = groupNodes.filter(group => !group.parentNode).sort((a, b) => {
    const la = String((a.data as any)?.label ?? a.id);
    const lb = String((b.data as any)?.label ?? b.id);
    return la.localeCompare(lb);
  });

  let yCursor = 80;
  const updatedNodes = new Map<string, Node>();

  const edgesFor = (setIds: Set<string>) =>
    edges.filter((e) => setIds.has(e.source) && setIds.has(e.target));

  // Layout each group internally (L→R) and stack groups vertically.
  for (const g of groupOrder) {
    const members = membersByGroup.get(g.id) ?? [];
    if (protectedNestedNodeIds.has(g.id)) {
      const { width, height } = nodeDimensions(g);
      updatedNodes.set(g.id, {
        ...g,
        position: { x: 80, y: yCursor },
        style: { ...(g.style ?? {}), width, height },
      });
      yCursor += height + spacing.laneGap;
      continue;
    }

    if (members.length === 0) {
      // Keep size but position lane.
      const { width, height } = nodeDimensions(g);
      updatedNodes.set(g.id, { ...g, position: { x: 80, y: yCursor }, style: { ...(g.style ?? {}), width, height } });
      yCursor += height + spacing.laneGap;
      continue;
    }

    const ids = new Set(members.map((m) => m.id));
    const subEdges = edgesFor(ids);

    // Temporarily strip parentNode so relayoutDiagram treats them as top-level,
    // then we re-apply as relative positions.
    const tempNodes: Node[] = [
      ...members.map((m) => ({ ...m, parentNode: undefined, extent: undefined } as any)),
    ];

    const laidOut = await doRelayout(tempNodes, subEdges, {
      direction: 'LR',
      nodeSpacing: spacing.nodeSpacing,
      rankSpacing: spacing.rankSpacing,
      groupPadding: spacing.groupPadding,
    });

    // Compute bounding box for members.
    let minX = Infinity;
    let minY = Infinity;
    let maxX = -Infinity;
    let maxY = -Infinity;

    for (const n of laidOut) {
      minX = Math.min(minX, n.position.x);
      minY = Math.min(minY, n.position.y);
      maxX = Math.max(maxX, n.position.x + NODE_WIDTH);
      maxY = Math.max(maxY, n.position.y + NODE_HEIGHT);
    }

    const contentW = Math.max(220, maxX - minX);
    const contentH = Math.max(160, maxY - minY);

    const groupWidth = contentW + spacing.groupPadding * 2;
    const groupHeight = contentH + spacing.groupPadding * 2;

    const groupX = 80;
    const groupY = yCursor;

    updatedNodes.set(g.id, {
      ...g,
      position: { x: groupX, y: groupY },
      style: { ...(g.style ?? {}), width: groupWidth, height: groupHeight },
    });

    for (const n of laidOut) {
      const relativeX = n.position.x - minX + spacing.groupPadding;
      const relativeY = n.position.y - minY + spacing.groupPadding;

      updatedNodes.set(n.id, {
        ...(members.find((m) => m.id === n.id) ?? n),
        position: { x: relativeX, y: relativeY },
        parentNode: g.id,
        extent: 'parent',
      } as any);
    }

    yCursor += groupHeight + spacing.laneGap;
  }

  // Layout ungrouped nodes in their own lane (no container).
  if (ungrouped.length > 0) {
    const ids = new Set(ungrouped.map((u) => u.id));
    const subEdges = edgesFor(ids);
    const laidOut = await doRelayout(
      ungrouped.map((m) => ({ ...m, parentNode: undefined, extent: undefined } as any)),
      subEdges,
      {
        direction: 'LR',
        nodeSpacing: spacing.nodeSpacing,
        rankSpacing: spacing.rankSpacing,
        groupPadding: spacing.groupPadding,
      }
    );

    // Place them below groups.
    const xBase = 80;
    const yBase = yCursor;
    for (const n of laidOut) {
      updatedNodes.set(n.id, {
        ...(ungrouped.find((u) => u.id === n.id) ?? n),
        position: { x: xBase + n.position.x, y: yBase + n.position.y },
        parentNode: undefined,
        extent: undefined,
      } as any);
    }
  }

  // Carry over any other nodes unchanged.
  const finalNodes = nodes.map((n) => updatedNodes.get(n.id) ?? n);
  return { nodes: finalNodes, edges };
}

function applyRadial(
  nodes: Node[],
  edges: Edge[],
  spacing: ReturnType<typeof getSpacing>,
  selectedNodeId?: string
) {
  const serviceNodes = nodes.filter((n) => n.type === 'azureNode');
  if (serviceNodes.length === 0) return { nodes, edges };

  const nodesById = new Map(nodes.map(node => [node.id, node]));
  const serviceEntityIds = new Map<string, string>();
  const entityNodes = new Map<string, Node>();

  const topLevelEntity = (node: Node): Node => {
    let current = node;
    let entity = node;
    const visiting = new Set<string>();
    while (current.parentNode && !visiting.has(current.id)) {
      visiting.add(current.id);
      const parent = nodesById.get(current.parentNode);
      if (!parent) break;
      if (parent.type === 'groupNode') entity = parent;
      current = parent;
    }
    return entity;
  };

  for (const service of serviceNodes) {
    const entity = topLevelEntity(service);
    serviceEntityIds.set(service.id, entity.id);
    entityNodes.set(entity.id, entity);
  }

  const adj = new Map<string, Set<string>>();
  for (const entityId of entityNodes.keys()) adj.set(entityId, new Set());
  for (const edge of edges) {
    const sourceEntity = serviceEntityIds.get(edge.source);
    const targetEntity = serviceEntityIds.get(edge.target);
    if (!sourceEntity || !targetEntity || sourceEntity === targetEntity) continue;
    adj.get(sourceEntity)?.add(targetEntity);
    adj.get(targetEntity)?.add(sourceEntity);
  }

  let centerId: string | undefined;
  if (selectedNodeId) {
    const selectedNode = nodesById.get(selectedNodeId);
    if (selectedNode) {
      const selectedEntity = topLevelEntity(selectedNode);
      if (entityNodes.has(selectedEntity.id)) centerId = selectedEntity.id;
    }
  }
  if (!centerId) {
    for (const [entityId, neighbors] of adj) {
      if (!centerId || neighbors.size > (adj.get(centerId)?.size ?? -1)) {
        centerId = entityId;
      }
    }
  }
  centerId ??= entityNodes.keys().next().value;
  if (!centerId) return { nodes, edges };

  const layer = new Map<string, number>();
  const q: string[] = [centerId];
  layer.set(centerId, 0);

  while (q.length) {
    const cur = q.shift()!;
    const d = layer.get(cur)!;
    for (const nb of adj.get(cur) ?? []) {
      if (layer.has(nb)) continue;
      layer.set(nb, d + 1);
      q.push(nb);
    }
  }

  const maxLayer = Math.max(...[...layer.values()], 0);
  for (const entityId of entityNodes.keys()) {
    if (!layer.has(entityId)) layer.set(entityId, maxLayer + 1);
  }

  const rings = new Map<number, string[]>();
  for (const [id, d] of layer.entries()) {
    rings.set(d, [...(rings.get(d) ?? []), id]);
  }

  const center = { x: 520, y: 360 };
  const entityExtents = new Map(
    [...entityNodes].map(([entityId, entity]) => {
      const { width, height } = nodeDimensions(entity);
      return [entityId, Math.hypot(width, height) / 2] as const;
    }),
  );
  const ringRadii = new Map<number, number>();
  let previousRadius = 0;
  let previousExtent = entityExtents.get(centerId) ?? 0;
  const ringGap = spacing.nodeSpacing;
  const ringIndexes = [...rings.keys()].filter(index => index > 0).sort((a, b) => a - b);
  for (const ringIndex of ringIndexes) {
    const ids = rings.get(ringIndex) ?? [];
    const maxExtent = Math.max(...ids.map(id => entityExtents.get(id) ?? 0), 0);
    const requiredCenterDistance = (maxExtent * 2) + ringGap;
    const chordRadius = ids.length > 1
      ? requiredCenterDistance / (2 * Math.sin(Math.PI / ids.length))
      : 0;
    const configuredRadius = spacing.radialBaseRadius
      + (ringIndex - 1) * spacing.radialRingStep;
    const separatedRadius = previousRadius + previousExtent + maxExtent + ringGap;
    const radius = Math.max(configuredRadius, chordRadius, separatedRadius);
    ringRadii.set(ringIndex, radius);
    previousRadius = radius;
    previousExtent = maxExtent;
  }

  const positions = new Map<string, { x: number; y: number }>();

  for (const [entityId, entity] of entityNodes) {
    let x = center.x;
    let y = center.y;
    if (entityId !== centerId) {
      const d = layer.get(entityId) ?? 1;
      const ids = rings.get(d) ?? [];
      const idx = ids.indexOf(entityId);
      const radius = ringRadii.get(d) ?? spacing.radialBaseRadius;
      const angle = (idx / Math.max(1, ids.length)) * Math.PI * 2;
      x += radius * Math.cos(angle);
      y += radius * Math.sin(angle);
    }

    const { width, height } = nodeDimensions(entity);
    positions.set(entityId, {
      x: x - (width / 2),
      y: y - (height / 2),
    });
  }

  const updatedNodes = nodes.map((node) => {
    const position = positions.get(node.id);
    return position ? { ...node, position } : node;
  });

  return { nodes: updatedNodes, edges };
}

/**
 * Unified relayout function type — works for both sync (Dagre) and async (ELK).
 */
type RelayoutFn = (nodes: any[], edges: any[], options?: any) => any[] | Promise<any[]>;

async function getRelayoutFn(engine: LayoutEngineType = 'dagre'): Promise<RelayoutFn> {
  if (engine === 'elk') {
    const { relayoutDiagram } = await import('./elkLayoutEngine');
    return relayoutDiagram;
  }
  const { relayoutDiagram } = await import('./layoutEngine');
  return relayoutDiagram;
}

/**
 * React Flow takes the direction an edge leaves and enters a tile from its
 * handles, not from where the tiles ended up. Only the generation path chose
 * them, so after a re-arrange — and above all after the serpentine wrap
 * reverses every other band — an edge running right-to-left still left the
 * right face and looped back into a target behind it, contradicting the
 * numbered badge on the same arrow. Re-derive them from the geometry the
 * layout actually produced.
 *
 * Only the two horizontal pairs the selector itself emits are rewritten, so a
 * hand-attached top/bottom handle is left alone.
 */
const HORIZONTAL_HANDLE_PAIRS = new Set(['right|left', 'left-source|right-target']);

function realignHorizontalConnectionHandles(nodes: Node[], edges: Edge[]): Edge[] {
  if (edges.length === 0) return edges;
  const absolutePositions = buildAbsolutePositionMap(nodes);
  let changed = false;
  const next = edges.map((edge) => {
    const pair = `${edge.sourceHandle ?? ''}|${edge.targetHandle ?? ''}`;
    if (!HORIZONTAL_HANDLE_PAIRS.has(pair)) return edge;
    const handles = selectHorizontalConnectionHandles(absolutePositions, edge.source, edge.target);
    if (handles.sourceHandle === edge.sourceHandle && handles.targetHandle === edge.targetHandle) return edge;
    changed = true;
    return { ...edge, ...handles };
  });
  return changed ? next : edges;
}

export async function applyLayoutPreset(nodes: Node[], edges: Edge[], opts: ApplyLayoutOptions): Promise<{ nodes: Node[]; edges: Edge[] }> {
  const result = await layoutWithPreset(nodes, edges, opts);
  return { nodes: result.nodes, edges: realignHorizontalConnectionHandles(result.nodes, result.edges) };
}

async function layoutWithPreset(nodes: Node[], edges: Edge[], opts: ApplyLayoutOptions): Promise<{ nodes: Node[]; edges: Edge[] }> {
  const spacing = getSpacing(opts.spacing);
  const doRelayout = await getRelayoutFn(opts.layoutEngine);

  // Always apply edge path style.
  const styledEdges = withEdgeStyle(edges, opts.edgeStyle);

  if (opts.preset === 'flow-lr') {
    const laidOutNodes = await doRelayout(nodes, styledEdges, {
      direction: 'LR',
      nodeSpacing: spacing.nodeSpacing,
      rankSpacing: spacing.rankSpacing,
      groupPadding: spacing.groupPadding,
    });

    if (opts.emphasizePrimaryPath) {
      return straightenPrimaryPath(laidOutNodes, styledEdges, 'LR');
    }

    return { nodes: laidOutNodes, edges: styledEdges };
  }

  if (opts.preset === 'flow-tb') {
    const laidOutNodes = await doRelayout(nodes, styledEdges, {
      direction: 'TB',
      nodeSpacing: spacing.nodeSpacing,
      rankSpacing: spacing.rankSpacing,
      groupPadding: spacing.groupPadding,
    });

    if (opts.emphasizePrimaryPath) {
      return straightenPrimaryPath(laidOutNodes, styledEdges, 'TB');
    }

    return { nodes: laidOutNodes, edges: styledEdges };
  }

  if (opts.preset === 'swimlanes') {
    const result = await applySwimlanesByGroup(nodes, styledEdges, spacing, doRelayout);
    return result;
  }

  // radial
  return applyRadial(nodes, styledEdges, spacing, opts.selectedNodeId);
}
