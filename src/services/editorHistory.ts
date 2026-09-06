import type { Edge, Node } from 'reactflow';
import { buildAbsolutePositionMap } from '../utils/preserveManualLayout';
import { canonicalStringify } from '../utils/canonicalJson';
import type { NodePricingConfig, PricingScenario } from '../types/pricing';
import type { IaCBaseline } from './iacRoundTrip';

export interface EditorDocument {
  nodes: Node[];
  edges: Edge[];
  titleBlockData: {
    architectureName: string;
    author: string;
    version: string;
    date: string;
  };
  workflow: unknown[];
  architecturePrompt: string;
  originalPrompt: string;
  settings: {
    pricingMode: 'payg' | 'reserved1yr';
    stylePreset: 'detailed' | 'presentation';
    edgeStyle: 'straight' | 'smooth' | 'orthogonal';
    pricingRegion?: string;
    animateConnections?: boolean;
    showCostBadges?: boolean;
    layoutPreset?: string;
    layoutSpacing?: string;
    layoutEngine?: 'elk' | 'dagre';
    emphasizePrimaryPath?: boolean;
  };
  reviewHistory?: unknown[];
  validationScore?: number;
  validationSourceFingerprint?: string | null;
  pricingScenarios?: PricingScenario[];
  iacBaseline?: IaCBaseline | null;
  lineageId?: string;
  viewport?: { x: number; y: number; zoom: number };
}

export function cloneEditorDocument(document: EditorDocument): EditorDocument {
  const copy: EditorDocument = JSON.parse(JSON.stringify(document));
  copy.nodes = copy.nodes.map(({ dragging, resizing, positionAbsolute, ...node }) => node);
  return copy;
}

export function editorFingerprint(document: EditorDocument, options: { includeReviewHistory?: boolean } = {}): string {
  const stored = cloneEditorDocument(document);
  const nodes = stored.nodes.map(node => {
    const {
      selected, dragging, resizing, positionAbsolute, width, height, measured, __rf, internals,
      ...content
    } = node as Node & Record<string, unknown>;
    return content;
  });
  const edges = stored.edges.map(({ selected, ...edge }) => edge);
  const { reviewHistory, validationScore, validationSourceFingerprint, viewport, lineageId, ...content } = stored;
  return canonicalStringify({
    ...content, nodes, edges,
    ...(options.includeReviewHistory ? {
      reviewHistory: reviewHistory ?? [],
      validationScore,
      validationSourceFingerprint,
      lineageId,
      viewport,
    } : {}),
  });
}

export function nodeServiceIdentity(node: Node): string {
  return JSON.stringify([node.type, node.data?.serviceName || node.data?.label || '']);
}

function enrichments(before: EditorDocument, after: EditorDocument): Map<string, NodePricingConfig> {
  const previous = new Map(before.nodes.map(node => [node.id, node]));
  const pricing = new Map<string, NodePricingConfig>();
  for (const node of after.nodes) {
    const old = previous.get(node.id);
    if (old && old.data.pricing == null && node.data.pricing?.isCustom === false
      && nodeServiceIdentity(old) === nodeServiceIdentity(node)
      && before.settings.pricingRegion === after.settings.pricingRegion
      && (!after.settings.pricingRegion || node.data.pricing.region === after.settings.pricingRegion)) {
      pricing.set(node.id, node.data.pricing);
    }
  }
  return pricing;
}

interface HistoryEntry {
  document: EditorDocument;
  fingerprint: string;
}

export class EditorHistory {
  private past: HistoryEntry[] = [];
  private future: HistoryEntry[] = [];
  private present: HistoryEntry;
  private gestureStart: HistoryEntry | null = null;
  revision = 0;

  constructor(document: EditorDocument, private limit = 75, private byteLimit = 16 * 1024 * 1024) {
    if (!Number.isInteger(limit) || limit < 1) throw new Error('History limit must be positive');
    this.present = this.entry(document);
  }

  private entry(document: EditorDocument): HistoryEntry {
    return { document: cloneEditorDocument(document), fingerprint: editorFingerprint(document) };
  }

  private trim(): void {
    let bytes = this.past.reduce((sum, entry) => sum + entry.fingerprint.length * 4, 0);
    while (this.past.length > this.limit || (this.past.length > 1 && bytes > this.byteLimit)) {
      bytes -= this.past.shift()!.fingerprint.length * 4;
    }
  }

  get canUndo(): boolean { return this.past.length > 0 || this.gestureStart !== null; }
  get canRedo(): boolean { return this.future.length > 0; }
  get current(): EditorDocument { return cloneEditorDocument(this.present.document); }

  enrichPricing(node: Node, pricing: NodePricingConfig): boolean {
    if (pricing.isCustom) return false;
    const identity = nodeServiceIdentity(node);
    let changed = false;
    const enrich = (entry: HistoryEntry): HistoryEntry => {
      if (entry.document.settings.pricingRegion && entry.document.settings.pricingRegion !== pricing.region) return entry;
      let enriched = false;
      const nodes = entry.document.nodes.map(candidate => {
        if (candidate.id !== node.id || candidate.data.pricing != null || nodeServiceIdentity(candidate) !== identity) return candidate;
        enriched = true;
        return { ...candidate, data: { ...candidate.data, pricing } };
      });
      if (!enriched) return entry;
      changed = true;
      return this.entry({ ...entry.document, nodes });
    };
    this.past = this.past.map(enrich);
    this.future = this.future.map(enrich);
    if (this.gestureStart) this.gestureStart = enrich(this.gestureStart);
    this.present = enrich(this.present);
    this.trim();
    return changed;
  }

  reset(document: EditorDocument): void {
    this.present = this.entry(document);
    this.past = [];
    this.future = [];
    this.gestureStart = null;
    this.revision += 1;
  }

  record(document: EditorDocument, transient = false): boolean {
    const next = this.entry(document);
    // Automatic lookup results belong to the edit that introduced the service,
    // even if a drag/rename was committed in the same React update.
    for (const [id, pricing] of enrichments(this.present.document, next.document)) {
      this.enrichPricing(next.document.nodes.find(node => node.id === id)!, pricing);
    }
    const changed = next.fingerprint !== this.present.fingerprint;
    if (!changed) {
      this.present = next;
      return !transient && this.finishGesture();
    }
    if (transient) {
      this.gestureStart ??= this.present;
    } else {
      this.past.push(this.gestureStart ?? this.present);
      this.gestureStart = null;
      this.trim();
    }
    this.present = next;
    this.future = [];
    this.revision += 1;
    return true;
  }

  finishGesture(): boolean {
    if (!this.gestureStart) return false;
    if (this.gestureStart.fingerprint !== this.present.fingerprint) this.past.push(this.gestureStart);
    this.gestureStart = null;
    this.trim();
    return true;
  }

  undo(): EditorDocument | null {
    this.finishGesture();
    const previous = this.past.pop();
    if (!previous) return null;
    this.future.push(this.present);
    this.present = this.withCurrentReview(previous);
    this.revision += 1;
    return this.current;
  }

  redo(): EditorDocument | null {
    const next = this.future.pop();
    if (!next) return null;
    this.past.push(this.present);
    this.trim();
    this.present = this.withCurrentReview(next);
    this.revision += 1;
    return this.current;
  }

  private withCurrentReview(entry: HistoryEntry): HistoryEntry {
    const document = { ...entry.document } as EditorDocument & Record<string, unknown>;
    const present = this.present.document as EditorDocument & Record<string, unknown>;
    for (const key of ['reviewHistory', 'validationScore', 'validationSourceFingerprint', 'lineageId', 'viewport']) {
      if (key in present) document[key] = present[key];
      else delete document[key];
    }
    return this.entry(document);
  }
}

export function deleteSelectedElements(nodes: Node[], edges: Edge[]): { nodes: Node[]; edges: Edge[] } {
  const removed = new Set(nodes.filter(node => node.selected).map(node => node.id));
  const byId = new Map(nodes.map(node => [node.id, node]));
  const absolute = buildAbsolutePositionMap(nodes);
  const remaining = nodes.filter(node => !removed.has(node.id)).map(node => {
    if (!node.parentNode || !removed.has(node.parentNode)) return node;
    let parent = node.parentNode;
    while (parent && removed.has(parent)) parent = byId.get(parent)?.parentNode ?? '';
    const position = absolute.get(node.id)!;
    const parentPosition = parent ? absolute.get(parent)! : { x: 0, y: 0 };
    return {
      ...node, parentNode: parent || undefined, extent: parent ? node.extent : undefined,
      position: { x: position.x - parentPosition.x, y: position.y - parentPosition.y },
      data: { ...node.data, parentNode: parent || undefined },
    };
  });
  return {
    nodes: remaining,
    edges: edges.filter(edge => !edge.selected && !removed.has(edge.source) && !removed.has(edge.target)),
  };
}

export function duplicateSelectedElements(
  nodes: Node[], edges: Edge[], createId: () => string = () => crypto.randomUUID(),
): { nodes: Node[]; edges: Edge[] } {
  const selected = new Set(nodes.filter(node => node.selected).map(node => node.id));
  let count = -1;
  while (count !== selected.size) {
    count = selected.size;
    for (const node of nodes) if (node.parentNode && selected.has(node.parentNode)) selected.add(node.id);
  }
  const ids = new Map([...selected].map(id => [id, createId()]));
  const duplicates = nodes.filter(node => selected.has(node.id)).map(node => {
    const parentNode = ids.get(node.parentNode ?? '') ?? node.parentNode;
    return {
      ...node, id: ids.get(node.id)!, parentNode,
      data: { ...node.data, parentNode },
      position: ids.has(node.parentNode ?? '') ? { ...node.position } : { x: node.position.x + 50, y: node.position.y + 50 },
      selected: true,
    };
  });
  return {
    nodes: [...nodes.map(node => ({ ...node, selected: false })), ...duplicates],
    edges: [
      ...edges.map(edge => ({ ...edge, selected: false })),
      ...edges.filter(edge => selected.has(edge.source) && selected.has(edge.target)).map(edge => ({
        ...edge, id: createId(), source: ids.get(edge.source)!, target: ids.get(edge.target)!,
        data: { ...edge.data }, selected: false,
      })),
    ],
  };
}
