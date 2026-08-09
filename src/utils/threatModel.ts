// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

import type { Edge, Node } from 'reactflow';

export type ThreatLevel = 'high' | 'medium' | 'low' | 'control';
export type ThreatKind = 'internet' | 'data' | 'identity' | 'secrets' | 'observability';

export interface ThreatMarker {
  nodeId: string;
  level: ThreatLevel;
  kind: ThreatKind;
  serviceName: string;
}

const includesAny = (value: string, terms: string[]) => terms.some(term => value.includes(term));

export function analyzeThreatModel(nodes: Node[], _edges: Edge[]): ThreatMarker[] {
  const markers: ThreatMarker[] = [];
  for (const node of nodes) {
    if (node.type !== 'azureNode') continue;
    const serviceName = String(node.data?.serviceName || node.data?.label || '');
    const normalized = serviceName.toLowerCase();

    if (includesAny(normalized, ['front door', 'application gateway', 'api management', 'public ip', 'traffic manager'])) {
      markers.push({ nodeId: node.id, level: 'high', kind: 'internet', serviceName });
      continue;
    }
    if (includesAny(normalized, ['sql', 'cosmos', 'storage', 'database', 'data lake', 'synapse'])) {
      markers.push({ nodeId: node.id, level: 'medium', kind: 'data', serviceName });
      continue;
    }
    if (includesAny(normalized, ['key vault', 'managed hsm'])) {
      markers.push({ nodeId: node.id, level: 'control', kind: 'secrets', serviceName });
      continue;
    }
    if (includesAny(normalized, ['entra', 'active directory', 'managed identity', 'b2c'])) {
      markers.push({ nodeId: node.id, level: 'control', kind: 'identity', serviceName });
      continue;
    }
    if (includesAny(normalized, ['monitor', 'application insights', 'log analytics', 'sentinel', 'defender'])) {
      markers.push({ nodeId: node.id, level: 'low', kind: 'observability', serviceName });
    }
  }
  return markers;
}
