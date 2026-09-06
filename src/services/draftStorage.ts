import type { Edge, Node } from 'reactflow';
import { cloneEditorDocument, editorFingerprint, type EditorDocument } from './editorHistory';
import { isPricingProvenance, isPricingUsage, validatePricingAmount, validatePricingQuantity } from './pricingConfiguration';
import { parseValidationReview } from './validationReview';

export const DRAFT_DATABASE = 'AzureDiagramDrafts';
export const DRAFT_STORE = 'drafts';

export interface DiagramDraft {
  id: string;
  schemaVersion: 1;
  revision: number;
  updatedAt: number;
  document: EditorDocument;
}

export class DraftConflictError extends Error {
  constructor() {
    super('Another tab has updated this draft. Download your current diagram before reloading.');
    this.name = 'DraftConflictError';
  }
}

function record(value: unknown): value is Record<string, unknown> {
  return value !== null && typeof value === 'object' && !Array.isArray(value);
}

export function validateRestoredNodes(values: unknown[]): Node[] {
  if (values.length > 2000) throw new Error('The diagram contains too many nodes.');
  const ids = new Set<string>();
  const nodes = values.map((value, index) => {
    if (!record(value) || typeof value.id !== 'string' || !value.id.trim()) {
      throw new Error(`Invalid graph node at index ${index}`);
    }
    if (ids.has(value.id)) throw new Error(`The diagram contains duplicate node id: ${value.id}`);
    if (!record(value.position) || !Number.isFinite(value.position.x) || !Number.isFinite(value.position.y)
      || !record(value.data) || (value.type !== undefined && typeof value.type !== 'string')
      || (value.parentNode !== undefined && typeof value.parentNode !== 'string')) {
      throw new Error(`Node ${value.id} contains invalid graph data`);
    }
    for (const field of ['width', 'height'] as const) {
      const size = value[field];
      if (size !== undefined && size !== null
        && (typeof size !== 'number' || !Number.isFinite(size) || size <= 0 || size > 100_000)) {
        throw new Error(`Node ${value.id} has an invalid ${field}`);
      }
      const styledSize = record(value.style) ? value.style[field] : undefined;
      if (styledSize !== undefined) {
        const numericSize = typeof styledSize === 'string' && /^\d+(?:\.\d+)?(?:px)?$/.test(styledSize)
          ? Number.parseFloat(styledSize) : styledSize;
        if (typeof numericSize !== 'number' || !Number.isFinite(numericSize) || numericSize <= 0 || numericSize > 100_000) {
          throw new Error(`Node ${value.id} has an invalid style ${field}`);
        }
      }
    }
    const data = { ...value.data };
    for (const field of ['label', 'serviceName', 'category', 'iconPath', 'description', 'stylePreset', 'groupId', 'groupLabel']) {
      if (data[field] !== undefined && typeof data[field] !== 'string') {
        throw new Error(`Node ${value.id} has an invalid ${field}`);
      }
    }
    if (data.tags !== undefined && (!Array.isArray(data.tags) || data.tags.length > 12
      || data.tags.some(tag => typeof tag !== 'string' || tag.length > 40))) {
      throw new Error(`Node ${value.id} has invalid tags`);
    }
    if (data.pricing !== undefined && data.pricing !== null) {
      if (!record(data.pricing)) throw new Error(`Node ${value.id} has invalid pricing data`);
      const pricing = { ...data.pricing };
      for (const field of ['estimatedCost', 'customPrice', 'reserved1yrCost']) {
        const amount = pricing[field];
        if (amount !== undefined && amount !== null) {
          if (typeof amount !== 'number') throw new Error(`Node ${value.id} has invalid pricing ${field}`);
          validatePricingAmount(amount);
        }
      }
      if (pricing.estimatedCost !== null && typeof pricing.estimatedCost !== 'number') {
        throw new Error(`Node ${value.id} has invalid pricing estimatedCost`);
      }
      const quantity = pricing.quantity ?? 1;
      if (typeof quantity !== 'number') throw new Error(`Node ${value.id} has invalid pricing quantity`);
      validatePricingQuantity(quantity);
      pricing.quantity = quantity;
      if (typeof pricing.estimatedCost === 'number' && !Number.isFinite(pricing.estimatedCost * quantity)) {
        throw new Error(`Node ${value.id} has an invalid total price`);
      }
      for (const field of ['isCustom', 'isUsageBased', 'reservedIsSavingsPlan']) {
        if (pricing[field] !== undefined && pricing[field] !== null && typeof pricing[field] !== 'boolean') {
          throw new Error(`Node ${value.id} has invalid pricing ${field}`);
        }
      }
      for (const field of ['tier', 'tierId', 'skuName', 'unit', 'lastUpdated', 'meterAsOf']) {
        if (pricing[field] !== undefined && pricing[field] !== null && typeof pricing[field] !== 'string') {
          throw new Error(`Node ${value.id} has invalid pricing ${field}`);
        }
      }
      if (pricing.usage !== undefined && !isPricingUsage(pricing.usage)) {
        throw new Error(`Node ${value.id} has invalid pricing usage`);
      }
      if (pricing.provenance !== undefined && !isPricingProvenance(pricing.provenance)) {
        throw new Error(`Node ${value.id} has invalid pricing provenance`);
      }
      if (pricing.region === undefined || pricing.region === null || (typeof pricing.region === 'string' && !pricing.region.trim())) {
        pricing.region = 'Unknown';
      } else if (typeof pricing.region !== 'string' || !/^[A-Za-z0-9][A-Za-z0-9 ._()-]{0,127}$/.test(pricing.region.trim())) {
        throw new Error(`Node ${value.id} has invalid pricing region`);
      } else pricing.region = pricing.region.trim();
      data.pricing = pricing;
    }
    if (data.customColor !== undefined && data.customColor !== null
      && (!record(data.customColor) || ['bg', 'border', 'header'].some(field => typeof (data.customColor as Record<string, unknown>)[field] !== 'string'))) {
      throw new Error(`Node ${value.id} has invalid custom colors`);
    }
    ids.add(value.id);
    return { ...value, position: { ...value.position }, data } as unknown as Node;
  });
  const byId = new Map(nodes.map(node => [node.id, node]));
  for (const node of nodes) {
    const visited = new Set([node.id]);
    let parent = node.parentNode;
    while (parent !== undefined) {
      if (visited.has(parent)) throw new Error('The diagram contains cyclic groups.');
      if (!byId.has(parent)) throw new Error(`Node ${node.id} has a dangling parent reference`);
      visited.add(parent);
      parent = byId.get(parent)!.parentNode;
    }
  }
  return nodes;
}

export function validateRestoredEdges(values: unknown[], nodeIds: Set<string>): Edge[] {
  if (values.length > 5000) throw new Error('The diagram contains too many edges.');
  const ids = new Set<string>();
  return values.map((value, index) => {
    if (!record(value) || typeof value.id !== 'string' || !value.id.trim()) {
      throw new Error(`Invalid edge at index ${index}`);
    }
    if (ids.has(value.id)) throw new Error(`The diagram contains duplicate edge id: ${value.id}`);
    if (typeof value.source !== 'string' || typeof value.target !== 'string'
      || !nodeIds.has(value.source) || !nodeIds.has(value.target)) {
      throw new Error(`Edge ${value.id} has dangling node references`);
    }
    if ((value.data !== undefined && value.data !== null && !record(value.data))
      || (value.type !== undefined && typeof value.type !== 'string')) {
      throw new Error(`Edge ${value.id} has invalid data`);
    }
    for (const field of ['sourceHandle', 'targetHandle']) {
      if (value[field] !== undefined && value[field] !== null && typeof value[field] !== 'string') {
        throw new Error(`Edge ${value.id} has an invalid ${field}`);
      }
    }
    ids.add(value.id);
    return { ...value } as Edge;
  });
}

export function validateRestoredWorkflow(value: unknown): Array<Record<string, unknown> & { step: number; description: string; services: string[] }> {
  if (value === undefined || value === null) return [];
  if (!Array.isArray(value) || value.length > 2000) throw new Error('Invalid workflow in diagram payload');
  return value.map((step, index) => {
    if (!record(step) || !Number.isFinite(step.step) || typeof step.description !== 'string'
      || !Array.isArray(step.services) || step.services.some(service => typeof service !== 'string')) {
      throw new Error(`Invalid workflow step at index ${index}`);
    }
    return { ...step, services: [...step.services] } as Record<string, unknown> & { step: number; description: string; services: string[] };
  });
}

export function validateEditorSettings(value: unknown, defaults: EditorDocument['settings']): EditorDocument['settings'] {
  if (value === undefined) return { ...defaults };
  if (!record(value)
    || !['payg', 'reserved1yr'].includes(String(value.pricingMode))
    || !['detailed', 'presentation'].includes(String(value.stylePreset))
    || !['straight', 'smooth', 'orthogonal'].includes(String(value.edgeStyle))) {
    throw new Error('Invalid draft display settings.');
  }
  if (value.pricingRegion !== undefined && (typeof value.pricingRegion !== 'string' || !/^[a-z][a-z0-9]{1,40}$/.test(value.pricingRegion))) {
    throw new Error('Invalid draft pricing region.');
  }
  for (const key of ['animateConnections', 'showCostBadges', 'emphasizePrimaryPath']) {
    if (value[key] !== undefined && typeof value[key] !== 'boolean') throw new Error(`Invalid editor setting ${key}`);
  }
  if (value.layoutPreset !== undefined && !['flow-lr', 'flow-tb', 'swimlanes', 'radial'].includes(String(value.layoutPreset))) throw new Error('Invalid editor layout preset');
  if (value.layoutSpacing !== undefined && !['compact', 'comfortable'].includes(String(value.layoutSpacing))) throw new Error('Invalid editor layout spacing');
  if (value.layoutEngine !== undefined && value.layoutEngine !== 'elk' && value.layoutEngine !== 'dagre') {
    throw new Error('Invalid editor layout engine');
  }
  return { ...value } as EditorDocument['settings'];
}

export function draftFingerprint(document: EditorDocument): string {
  return editorFingerprint(document, { includeReviewHistory: true });
}

export function validateDraft(value: unknown, key: string): DiagramDraft {
  if (!record(value) || value.id !== key || value.schemaVersion !== 1
    || typeof value.revision !== 'number' || !Number.isSafeInteger(value.revision) || value.revision < 1
    || typeof value.updatedAt !== 'number' || !Number.isFinite(value.updatedAt)
    || !record(value.document)) {
    throw new Error('The saved draft has an invalid or unsupported format.');
  }
  const document = value.document;
  if (!Array.isArray(document.nodes) || !Array.isArray(document.edges)
    || document.nodes.length > 2000 || document.edges.length > 5000
    || !Array.isArray(document.workflow) || document.workflow.length > 2000
    || typeof document.architecturePrompt !== 'string' || typeof document.originalPrompt !== 'string'
    || !record(document.titleBlockData) || !record(document.settings)) {
    throw new Error('The saved draft is incomplete.');
  }
  for (const field of ['architectureName', 'author', 'date', 'version']) {
    if (typeof document.titleBlockData[field] !== 'string') throw new Error('Invalid draft title block.');
  }
  const settings = validateEditorSettings(document.settings, { pricingMode: 'payg', stylePreset: 'detailed', edgeStyle: 'orthogonal' });
  const nodes = validateRestoredNodes(document.nodes);
  const edges = validateRestoredEdges(document.edges, new Set(nodes.map(node => node.id)));
  const workflow = validateRestoredWorkflow(document.workflow);
  const reviewHistory = document.reviewHistory === undefined ? undefined : parseValidationReview(document.reviewHistory);
  if (document.lineageId !== undefined && (typeof document.lineageId !== 'string' || document.lineageId.length > 320)) {
    throw new Error('Invalid draft document lineage.');
  }
  if (document.validationScore !== undefined && (typeof document.validationScore !== 'number' || !Number.isFinite(document.validationScore)
    || document.validationScore < 0 || document.validationScore > 100)) throw new Error('Invalid draft validation score.');
  if (document.validationSourceFingerprint !== undefined && document.validationSourceFingerprint !== null
    && typeof document.validationSourceFingerprint !== 'string') throw new Error('Invalid draft validation provenance.');
  if (document.pricingScenarios !== undefined && !Array.isArray(document.pricingScenarios)) throw new Error('Invalid draft pricing scenarios.');
  if (document.iacBaseline !== undefined && document.iacBaseline !== null && !record(document.iacBaseline)) throw new Error('Invalid draft infrastructure baseline.');
  if (document.viewport !== undefined && (!record(document.viewport) || !Number.isFinite(document.viewport.x)
    || !Number.isFinite(document.viewport.y) || typeof document.viewport.zoom !== 'number'
    || !Number.isFinite(document.viewport.zoom) || document.viewport.zoom <= 0)) throw new Error('Invalid draft viewport.');
  return {
    id: key, schemaVersion: 1, revision: value.revision, updatedAt: value.updatedAt,
    document: {
      ...(document as unknown as EditorDocument),
      nodes, edges, workflow,
      titleBlockData: {
        architectureName: String(document.titleBlockData.architectureName),
        author: String(document.titleBlockData.author),
        version: String(document.titleBlockData.version),
        date: String(document.titleBlockData.date),
      },
      architecturePrompt: document.architecturePrompt, originalPrompt: document.originalPrompt,
      settings,
      ...(reviewHistory !== undefined ? { reviewHistory } : {}),
    },
  };
}

function openDatabase(): Promise<IDBDatabase> {
  return new Promise((resolve, reject) => {
    const request = indexedDB.open(DRAFT_DATABASE, 1);
    let failed = false;
    request.onupgradeneeded = () => {
      request.result.createObjectStore(DRAFT_STORE, { keyPath: 'id' });
    };
    request.onblocked = () => {
      failed = true;
      reject(new Error('Draft storage is blocked by another tab. Close that tab and retry.'));
    };
    request.onerror = () => reject(request.error);
    request.onsuccess = () => {
      const database = request.result;
      database.onversionchange = () => database.close();
      if (failed) database.close();
      else resolve(database);
    };
  });
}

export async function readDraft(key: string): Promise<DiagramDraft | null> {
  const database = await openDatabase();
  return new Promise((resolve, reject) => {
    const transaction = database.transaction(DRAFT_STORE, 'readonly');
    const request = transaction.objectStore(DRAFT_STORE).get(key);
    let result: DiagramDraft | null = null;
    let error: unknown;
    request.onsuccess = () => {
      try {
        result = request.result === undefined ? null : validateDraft(request.result, key);
      } catch (cause) {
        error = cause;
        transaction.abort();
      }
    };
    transaction.oncomplete = () => { database.close(); resolve(result); };
    transaction.onabort = () => { database.close(); reject(error ?? transaction.error ?? new Error('Draft read aborted.')); };
  });
}

async function mutateDraft(key: string, expectedRevision: number | null, document: EditorDocument | null): Promise<DiagramDraft | null> {
  const snapshot = document ? cloneEditorDocument(document) : null;
  if (snapshot && JSON.stringify(snapshot).length > 10 * 1024 * 1024) {
    throw new Error('The draft is larger than 10 MB. Download the diagram to keep a copy.');
  }
  const database = await openDatabase();
  return new Promise((resolve, reject) => {
    const transaction = database.transaction(DRAFT_STORE, 'readwrite');
    const store = transaction.objectStore(DRAFT_STORE);
    const request = store.get(key);
    let result: DiagramDraft | null = null;
    let error: unknown;
    request.onsuccess = () => {
      try {
        const current = request.result === undefined ? null : validateDraft(request.result, key);
        if ((current?.revision ?? null) !== expectedRevision) throw new DraftConflictError();
        if (!snapshot) {
          store.delete(key);
        } else {
          result = validateDraft({
            id: key, schemaVersion: 1, revision: (current?.revision ?? 0) + 1,
            updatedAt: Date.now(), document: snapshot,
          }, key);
          store.put(result);
        }
      } catch (cause) {
        error = cause;
        transaction.abort();
      }
    };
    transaction.oncomplete = () => { database.close(); resolve(result); };
    transaction.onabort = () => { database.close(); reject(error ?? transaction.error ?? new Error('Draft save aborted.')); };
  });
}

export async function writeDraft(key: string, document: EditorDocument, expectedRevision: number | null): Promise<DiagramDraft> {
  const result = await mutateDraft(key, expectedRevision, document);
  if (!result) throw new Error('The draft transaction did not produce a saved document.');
  return result;
}

export async function discardDraft(key: string, expectedRevision: number): Promise<void> {
  await mutateDraft(key, expectedRevision, null);
}
