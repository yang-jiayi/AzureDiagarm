// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

import type { Edge, Node } from 'reactflow';
import { MAX_PRICING_AMOUNT, MAX_PRICING_QUANTITY } from './pricingConfiguration';

export interface DiagramGraph {
  nodes: Node[];
  edges: Edge[];
}

export interface DiagramChange {
  id: string;
  entity: 'node' | 'edge';
  kind: 'add' | 'delete' | 'change';
  entityId: string;
  label: string;
  before?: Node | Edge;
  after?: Node | Edge;
  fields: string[];
  /** Monthly USD delta, only when both relevant estimates are known. */
  costDelta?: number;
}

export interface DiagramChangeSet {
  before: DiagramGraph;
  proposed: DiagramGraph;
  changes: DiagramChange[];
}

export interface DiagramProposalOptions {
  /** Imported graph identities and metadata must not be merged with the current document. */
  reconcile?: boolean;
}

export class DiagramSelectionError extends Error {
  constructor(readonly issues: string[]) {
    super(issues.join('\n'));
    this.name = 'DiagramSelectionError';
  }
}

const transientFields = new Set([
  'selected', 'dragging', 'resizing', 'positionAbsolute', 'width', 'height',
  'measured', '__rf', 'internals',
]);
const record = (value: unknown): value is Record<string, any> =>
  !!value && typeof value === 'object' && !Array.isArray(value);

/** Unlike JSON/structuredClone, this keeps editor callbacks without comparing them. */
function copy<T>(value: T): T {
  if (Array.isArray(value)) return value.map(copy) as T;
  if (record(value)) return Object.fromEntries(Object.entries(value).map(([key, item]) => [key, copy(item)])) as T;
  return value;
}

function content(value: unknown, topLevel = false): unknown {
  if (typeof value === 'function' || value === undefined) return undefined;
  if (Array.isArray(value)) return value.map(item => content(item));
  if (record(value)) {
    return Object.fromEntries(Object.keys(value).sort()
      .filter(key => !(topLevel && transientFields.has(key)))
      .map(key => [key, content(value[key])])
      .filter(([key, item]) => item !== undefined
        && !(topLevel && key === 'data' && record(item) && Object.keys(item).length === 0)));
  }
  return value;
}

function equal(left: unknown, right: unknown, topLevel = false): boolean {
  return JSON.stringify(content(left, topLevel)) === JSON.stringify(content(right, topLevel));
}

function merge(previous: any, proposed: any): any {
  if (typeof previous === 'function') return previous;
  if (!record(previous) || !record(proposed)) return copy(proposed);
  const result = copy(previous);
  for (const [key, value] of Object.entries(proposed)) {
    result[key] = merge(previous[key], value);
  }
  return result;
}

const identity = (value: unknown): string =>
  typeof value === 'string' ? value.normalize('NFKC').trim().toLowerCase().replace(/\s+/g, ' ') : '';
const nodeLabel = (node: Node): string => String(node.data?.label || node.data?.serviceName || node.id);
const serviceIdentity = (node: Node): string => identity(node.data?.serviceName || node.data?.label);
const parentId = (node: Node): string | undefined => node.parentNode ?? (node as any).parentId;

function uniqueMatches<T extends { id: string }>(
  previous: T[], proposed: T[], matched: Map<string, T>, key: (value: T) => string,
): void {
  const used = new Set([...matched.values()].map(value => value.id));
  // Count across the whole graph, not just unmatched entities: two instances
  // of one Azure service must not become "unique" after an arbitrary match.
  const left = new Map<string, T[]>();
  const right = new Map<string, T[]>();
  for (const item of previous) {
    const id = key(item);
    if (id) left.set(id, [...(left.get(id) || []), item]);
  }
  for (const item of proposed) {
    const id = key(item);
    if (id) right.set(id, [...(right.get(id) || []), item]);
  }
  for (const [id, values] of right) {
    const candidates = left.get(id);
    if (values.length !== 1 || candidates?.length !== 1) continue;
    if (matched.has(values[0].id) || used.has(candidates[0].id)) continue;
    matched.set(values[0].id, candidates[0]);
    used.add(candidates[0].id);
  }
}

function validate(graph: DiagramGraph): void {
  const issues: string[] = [];
  const nodes = new Map<string, Node>();
  const edgeIds = new Set<string>();
  for (const node of graph.nodes) {
    if (!node.id || nodes.has(node.id)) issues.push(`Duplicate or empty node ID: ${node.id}`);
    nodes.set(node.id, node);
  }
  for (const edge of graph.edges) {
    if (!edge.id || edgeIds.has(edge.id)) issues.push(`Duplicate or empty edge ID: ${edge.id}`);
    edgeIds.add(edge.id);
    for (const id of [edge.source, edge.target]) {
      if (!nodes.has(id)) issues.push(`Connection "${String(edge.label || edge.id)}" requires node "${id}".`);
    }
  }
  for (const node of graph.nodes) {
    if (node.parentNode && (node as any).parentId && node.parentNode !== (node as any).parentId) {
      issues.push(`Node "${nodeLabel(node)}" has conflicting parents.`);
    }
    const seen = new Set([node.id]);
    let parent = parentId(node);
    while (parent) {
      if (seen.has(parent)) {
        issues.push(`Parent cycle involving "${nodeLabel(node)}".`);
        break;
      }
      seen.add(parent);
      const ancestor = nodes.get(parent);
      if (!ancestor) {
        issues.push(`Node "${nodeLabel(node)}" requires parent "${parent}".`);
        break;
      }
      parent = parentId(ancestor);
    }
  }
  if (issues.length) throw new DiagramSelectionError([...new Set(issues)]);
}

function normalizeProposal(before: DiagramGraph, proposed: DiagramGraph): DiagramGraph {
  const matches = new Map<string, Node>();
  const byId = new Map(before.nodes.map(node => [node.id, node]));
  for (const node of proposed.nodes) {
    const previous = byId.get(node.id);
    if (previous) matches.set(node.id, previous);
  }
  uniqueMatches(before.nodes, proposed.nodes, matches, node =>
    serviceIdentity(node) ? `${node.type || ''}\0${serviceIdentity(node)}` : '');
  uniqueMatches(before.nodes, proposed.nodes, matches, node =>
    identity(node.data?.label) ? `${node.type || ''}\0${identity(node.data?.label)}` : '');
  const ids = new Map(proposed.nodes.map(node => [node.id, matches.get(node.id)?.id || node.id]));
  const nodes = proposed.nodes.map(node => {
    const previous = matches.get(node.id);
    const next = previous ? merge(previous, node) as Node : copy(node);
    if (previous && serviceIdentity(previous) !== serviceIdentity(node) && !node.data?.pricing) {
      delete next.data.pricing;
    }
    next.id = ids.get(node.id)!;
    // Parent omission means explicitly ungrouped, not "keep the old parent".
    delete next.parentNode;
    delete (next as any).parentId;
    delete next.extent;
    const parent = parentId(node);
    if (parent) {
      next.parentNode = ids.get(parent) || parent;
      next.extent = node.extent;
    }
    if (next.data && 'parentNode' in next.data) next.data.parentNode = next.parentNode;
    // Generated service names must not erase a user's descriptive instance label.
    if (previous && identity(node.data?.label) === identity(node.data?.serviceName)
      && identity(previous.data?.label) !== identity(previous.data?.serviceName)
      && serviceIdentity(previous) === serviceIdentity(node)) {
      next.data.label = previous.data.label;
    }
    if (previous && node.type === 'azureNode' && node.data?.serviceName && previous.data?.stylePreset !== undefined) {
      next.data.stylePreset = previous.data.stylePreset;
    }
    if (previous && node.id !== previous.id && node.type === 'azureNode' && node.data?.serviceName
      && parentId(previous) === parentId(next)) {
      next.position = copy(previous.position);
    }
    return next;
  });
  const edges = proposed.edges.map(edge => ({
    ...copy(edge), source: ids.get(edge.source) || edge.source, target: ids.get(edge.target) || edge.target,
  }));
  const edgeMatches = new Map<string, Edge>();
  const previousEdges = new Map(before.edges.map(edge => [edge.id, edge]));
  const generatedId = (id: string) => /^edge-\d+$/.test(id);
  for (const edge of edges) {
    const previous = previousEdges.get(edge.id);
    // edge-N is an array index, not identity. Never match reordered outputs by it.
    if (previous && !generatedId(edge.id)) edgeMatches.set(edge.id, previous);
  }
  const endpoints = (edge: Edge) => JSON.stringify([edge.source, edge.target]);
  const signature = (edge: Edge) => JSON.stringify([
    edge.source, edge.target, content(edge.label), edge.data?.connectionType || 'sync',
  ]);
  // Match parallel connections as a multiset. Identical labels still represent
  // separate edges; each old edge (and its routing) can be consumed only once.
  const used = new Set([...edgeMatches.values()].map(edge => edge.id));
  for (const edge of edges) {
    if (edgeMatches.has(edge.id)) continue;
    const candidates = before.edges.filter(old => !used.has(old.id) && signature(old) === signature(edge));
    if (candidates.length) {
      const old = candidates.find(old => old.id === edge.id) || candidates[0];
      edgeMatches.set(edge.id, old);
      used.add(old.id);
    }
  }
  uniqueMatches(before.edges, edges, edgeMatches, endpoints);
  const reserved = new Set([...before.edges, ...edges].map(edge => edge.id));
  return {
    nodes,
    edges: edges.map(edge => {
      const previous = edgeMatches.get(edge.id);
      if (previous) {
        const next: Edge = { ...merge(previous, edge), id: previous.id };
        if (generatedId(edge.id) && endpoints(previous) === endpoints(edge)) {
          // Parent-generated edge-N records contain fresh routing defaults,
          // not a user request to erase hand-routed handles or label positions.
          for (const key of ['sourceHandle', 'targetHandle', 'labelStyle', 'labelBgStyle', 'labelBgPadding', 'labelBgBorderRadius']) {
            if (key in previous) (next as any)[key] = copy((previous as any)[key]);
          }
          for (const key of ['labelOffsetX', 'labelOffsetY', 'pathStyle', 'controlPoints', 'waypoints', 'bendPoints']) {
            if (previous.data && key in previous.data) {
              next.data = { ...next.data, [key]: copy(previous.data[key]) };
            }
          }
          if ((previous.data?.connectionType || 'sync') === (edge.data?.connectionType || 'sync')) {
            next.style = copy(previous.style ?? next.style);
          }
        }
        return next;
      }
      let id = edge.id;
      if (previousEdges.has(id)) {
        let suffix = 1;
        while (reserved.has(`${id}-ai-${suffix}`)) suffix++;
        id = `${id}-ai-${suffix}`;
        reserved.add(id);
      }
      return { ...edge, id };
    }),
  };
}

function monthlyCost(node: Node | Edge | undefined): number | undefined {
  if (!node) return 0;
  const pricing = node.data?.pricing;
  const value = pricing ? pricing.estimatedCost : node.data?.monthlyCost ?? node.data?.estimatedCost;
  if (typeof value !== 'number' || !Number.isFinite(value) || value < 0 || value > MAX_PRICING_AMOUNT) return undefined;
  const quantity = pricing?.quantity ?? 1;
  const total = value * quantity;
  return typeof quantity === 'number' && Number.isInteger(quantity) && quantity >= 1
    && quantity <= MAX_PRICING_QUANTITY && Number.isFinite(total) ? total : undefined;
}

export function buildDiagramChanges(
  before: DiagramGraph, proposed: DiagramGraph, options: DiagramProposalOptions = {},
): DiagramChangeSet {
  validate(before);
  validate(proposed);
  const baseline = copy(before);
  const normalized = options.reconcile === false ? copy(proposed) : normalizeProposal(baseline, proposed);
  validate(normalized);
  const changes: DiagramChange[] = [];
  const allNodes = new Map([...baseline.nodes, ...normalized.nodes].map(node => [node.id, node]));
  for (const entity of ['node', 'edge'] as const) {
    const previous: (Node | Edge)[] = entity === 'node' ? baseline.nodes : baseline.edges;
    const next: (Node | Edge)[] = entity === 'node' ? normalized.nodes : normalized.edges;
    const left = new Map(previous.map(item => [item.id, item]));
    const right = new Map(next.map(item => [item.id, item]));
    for (const id of new Set([...left.keys(), ...right.keys()])) {
      const old = left.get(id);
      const value = right.get(id);
      if (old && value && equal(old, value, true)) continue;
      const kind = !old ? 'add' : !value ? 'delete' : 'change';
      const item = (value || old)!;
      const edge = item as Edge;
      const label = entity === 'node' ? nodeLabel(item as Node)
        : `${allNodes.has(edge.source) ? nodeLabel(allNodes.get(edge.source)!) : edge.source} → ${allNodes.has(edge.target) ? nodeLabel(allNodes.get(edge.target)!) : edge.target}${edge.label ? ` · ${String(edge.label)}` : ''}`;
      const oldCost = monthlyCost(old);
      const nextCost = monthlyCost(value);
      changes.push({
        id: `${entity}:${kind}:${id}`, entity, kind, entityId: id, label,
        before: old, after: value,
        fields: old && value ? [...new Set([...Object.keys(old), ...Object.keys(value)])]
          .filter(key => !transientFields.has(key) && !equal((old as any)[key], (value as any)[key])) : [],
        ...(entity === 'node' && oldCost !== undefined && nextCost !== undefined
          ? { costDelta: nextCost - oldCost } : {}),
      });
    }
  }
  return { before: baseline, proposed: normalized, changes };
}

/** Reject invalid subsets rather than silently adding/dropping unselected work. */
export function applyDiagramChanges(changeSet: DiagramChangeSet, selectedIds: ReadonlySet<string>): DiagramGraph {
  const known = new Set(changeSet.changes.map(change => change.id));
  const unknown = [...selectedIds].filter(id => !known.has(id));
  if (unknown.length) throw new DiagramSelectionError(unknown.map(id => `Unknown change: ${id}`));
  const nodes = new Map(changeSet.before.nodes.map(node => [node.id, copy(node)]));
  const edges = new Map(changeSet.before.edges.map(edge => [edge.id, copy(edge)]));
  for (const change of changeSet.changes) {
    if (!selectedIds.has(change.id)) continue;
    const values = change.entity === 'node' ? nodes : edges;
    if (change.kind === 'delete') values.delete(change.entityId);
    else if (change.after) (values as Map<string, Node | Edge>).set(change.entityId, copy(change.after));
  }
  const graph = { nodes: [...nodes.values()], edges: [...edges.values()] };
  validate(graph);
  // React Flow requires parents before their descendants.
  const sorted: Node[] = [];
  const visited = new Set<string>();
  const visit = (node: Node) => {
    if (visited.has(node.id)) return;
    const parent = parentId(node);
    if (parent) visit(nodes.get(parent)!);
    visited.add(node.id);
    sorted.push(node);
  };
  graph.nodes.forEach(visit);
  graph.nodes = sorted;
  return graph;
}

/**
 * Compare an accepted graph with the complete, normalized proposal. Collection
 * order, callback references and React Flow measurements are not diagram edits.
 * This also works when the review has no changes or the baseline is empty.
 */
export function isCompleteDiagramChangeSet(changeSet: DiagramChangeSet, graph: DiagramGraph): boolean {
  try {
    validate(graph);
  } catch (error) {
    if (error instanceof DiagramSelectionError) return false;
    throw error;
  }
  const complete = applyDiagramChanges(changeSet, new Set(changeSet.changes.map(change => change.id)));
  const sameEntities = <T extends Node | Edge>(expected: T[], actual: T[]) => {
    if (expected.length !== actual.length) return false;
    const byId = new Map(actual.map(entity => [entity.id, entity]));
    return expected.every(entity => byId.has(entity.id) && equal(entity, byId.get(entity.id), true));
  };
  return sameEntities(complete.nodes, graph.nodes) && sameEntities(complete.edges, graph.edges);
}
