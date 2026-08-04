// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

/**
 * Architecture Importer
 *
 * Reverses the two export formats this server produces back into the canonical
 * { services, connections, groups } shape that every other tool consumes:
 *
 *   - `generate_manifest`  → az prototype interchange manifest (clean, lossless)
 *   - `export_reactflow_scene` → React Flow scene (nodes/edges); service type is
 *     read from data.serviceName/data.azureServiceType, else from the icon path.
 *
 * Tolerant by design: it accepts web-app-native scenes and manifests too, and
 * collects warnings rather than throwing on partially-recognized input.
 */

export interface ImportedService {
  name: string;
  type: string;
  description?: string;
  groupId?: string;
}
export interface ImportedConnection {
  from: string;
  to: string;
  label?: string;
  type?: string;
}
export interface ImportedGroup {
  id: string;
  label: string;
}
export interface ImportResult {
  format: 'manifest' | 'reactflow';
  projectName?: string;
  location?: string;
  services: ImportedService[];
  connections: ImportedConnection[];
  groups: ImportedGroup[];
  warnings: string[];
}

type AnyObj = Record<string, unknown>;
const MAX_IMPORTED_SERVICES = 250;
const MAX_IMPORTED_CONNECTIONS = 1_000;
const MAX_IMPORTED_GROUPS = 100;

function assertImportBounds(
  serviceCount: number,
  connectionCount: number,
  groupCount: number,
): void {
  if (serviceCount > MAX_IMPORTED_SERVICES) {
    throw new RangeError(`Architecture imports support at most ${MAX_IMPORTED_SERVICES} services.`);
  }
  if (connectionCount > MAX_IMPORTED_CONNECTIONS) {
    throw new RangeError(`Architecture imports support at most ${MAX_IMPORTED_CONNECTIONS} connections.`);
  }
  if (groupCount > MAX_IMPORTED_GROUPS) {
    throw new RangeError(`Architecture imports support at most ${MAX_IMPORTED_GROUPS} groups.`);
  }
}

function asObject(input: unknown): AnyObj {
  if (typeof input === 'string') {
    try {
      return JSON.parse(input) as AnyObj;
    } catch (e) {
      throw new Error(`Input is not valid JSON: ${(e as Error).message}`);
    }
  }
  if (input && typeof input === 'object') return input as AnyObj;
  throw new Error('Input must be a JSON string or object.');
}

export function buildIconFileToTypeMap(
  iconMap: Record<string, { iconFile?: string }>,
): Record<string, string> {
  const candidates = new Map<string, string[]>();
  for (const [serviceName, entry] of Object.entries(iconMap)) {
    if (!entry.iconFile) continue;
    const names = candidates.get(entry.iconFile) ?? [];
    names.push(serviceName);
    candidates.set(entry.iconFile, names);
  }

  return Object.fromEntries(
    [...candidates.entries()]
      .filter(([, serviceNames]) => serviceNames.length === 1)
      .map(([iconFile, [serviceName]]) => [iconFile, serviceName]),
  );
}

/** Reverse a React Flow node's Azure service type. */
function typeFromNode(node: AnyObj, iconFileToType?: Record<string, string>): string | null {
  const d = (node.data ?? {}) as AnyObj;
  const explicit =
    (d.serviceName as string) ??
    (d.azureServiceType as string) ??
    (d.serviceType as string) ??
    (d.azureType as string) ??
    (d.type as string) ??
    (node.azureServiceType as string);
  if (typeof explicit === 'string' && explicit.trim()) return explicit.trim();

  const ip = (d.iconPath as string) ?? (d.icon as string);
  if (typeof ip === 'string' && iconFileToType) {
    const m = ip.match(/\/([^/]+)\.svg$/i);
    if (m && iconFileToType[m[1]]) return iconFileToType[m[1]];
  }
  const label = d.label as string;
  if (typeof label === 'string' && label.trim()) return label.trim();
  return null;
}

function isGroupNode(node: AnyObj): boolean {
  const d = (node.data ?? {}) as AnyObj;
  return node.type === 'groupNode' || node.type === 'group' || d.isGroup === true;
}

/**
 * Import an architecture from a manifest or React Flow scene into the canonical
 * shape. Pass `iconFileToType` (reverse of the icon map) to recover service
 * types from icon paths when a scene lacks an explicit type field.
 */
export function importArchitecture(
  input: unknown,
  opts: { iconFileToType?: Record<string, string> } = {},
): ImportResult {
  const obj = asObject(input);
  const warnings: string[] = [];

  // ── Manifest format ───────────────────────────────────────────────────
  if (obj.architecture && typeof obj.architecture === 'object') {
    const arch = obj.architecture as AnyObj;
    const project = (obj.project ?? {}) as AnyObj;
    const rawServices = Array.isArray(arch.services) ? (arch.services as AnyObj[]) : [];
    const rawConns = Array.isArray(arch.connections) ? (arch.connections as AnyObj[]) : [];
    const rawGroups = Array.isArray(arch.groups) ? (arch.groups as AnyObj[]) : [];
    assertImportBounds(rawServices.length, rawConns.length, rawGroups.length);

    const services: ImportedService[] = rawServices.map(s => ({
      name: String(s.name ?? s.id ?? 'Unnamed'),
      type: String(s.type ?? 'Unknown'),
      description: s.description ? String(s.description) : undefined,
      groupId: s.groupId ? String(s.groupId) : undefined,
    }));
    const connections: ImportedConnection[] = rawConns.map(c => ({
      from: String(c.from),
      to: String(c.to),
      label: c.label ? String(c.label) : undefined,
      type: c.type ? String(c.type) : undefined,
    }));
    const groups: ImportedGroup[] = rawGroups.map(g => ({
      id: String(g.id),
      label: String(g.label ?? g.id),
    }));

    if (services.length === 0) warnings.push('Manifest contained no services.');
    return {
      format: 'manifest',
      projectName: project.name ? String(project.name) : undefined,
      location: project.location ? String(project.location) : undefined,
      services,
      connections,
      groups,
      warnings,
    };
  }

  // ── React Flow scene format ───────────────────────────────────────────
  if (Array.isArray(obj.nodes)) {
    const nodes = obj.nodes as AnyObj[];
    const edges = Array.isArray(obj.edges) ? (obj.edges as AnyObj[]) : [];
    const serviceNodeCount = nodes.reduce(
      (count, node) => count + (isGroupNode(node) ? 0 : 1),
      0,
    );
    const groupNodeCount = nodes.length - serviceNodeCount;
    assertImportBounds(serviceNodeCount, edges.length, groupNodeCount);

    const groups: ImportedGroup[] = [];
    const services: ImportedService[] = [];
    const nameByNodeId = new Map<string, string>();
    const groupNodeIds = new Set<string>();

    for (const node of nodes) {
      const id = String(node.id ?? '');
      if (!id) continue;
      if (isGroupNode(node)) {
        groupNodeIds.add(id);
        const d = (node.data ?? {}) as AnyObj;
        groups.push({ id, label: String(d.label ?? id) });
        continue;
      }
      const d = (node.data ?? {}) as AnyObj;
      const name = String(d.label ?? node.id ?? 'Unnamed');
      const type = typeFromNode(node, opts.iconFileToType);
      if (!type) warnings.push(`Node "${name}" had no resolvable service type; using label.`);
      const parent = (node.parentNode ?? node.parentId) as string | undefined;
      services.push({
        name,
        type: type ?? name,
        description: d.description ? String(d.description) : undefined,
        groupId: parent && groupNodeIds.has(String(parent)) ? String(parent) : (parent ? String(parent) : undefined),
      });
      nameByNodeId.set(id, name);
    }

    const connections: ImportedConnection[] = [];
    for (const e of edges) {
      const from = nameByNodeId.get(String(e.source));
      const to = nameByNodeId.get(String(e.target));
      if (!from || !to) continue; // skip edges touching group nodes / unknowns
      const d = (e.data ?? {}) as AnyObj;
      connections.push({
        from,
        to,
        label: e.label ? String(e.label) : undefined,
        type: d.connectionType ? String(d.connectionType) : undefined,
      });
    }

    const meta = (obj.metadata ?? {}) as AnyObj;
    if (services.length === 0) warnings.push('Scene contained no service nodes.');
    return {
      format: 'reactflow',
      projectName: meta.architectureName ? String(meta.architectureName) : undefined,
      services,
      connections,
      groups,
      warnings,
    };
  }

  throw new Error(
    'Unrecognized architecture format. Expected an az prototype manifest (has "architecture") or a React Flow scene (has "nodes").',
  );
}
