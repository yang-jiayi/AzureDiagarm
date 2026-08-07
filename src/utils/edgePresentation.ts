// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

export type DiagramConnectionType =
  | 'sync'
  | 'async'
  | 'optional'
  | 'security'
  | 'telemetry';

export interface ConnectionPresentation {
  type: DiagramConnectionType;
  stroke: string;
  strokeDasharray?: string;
  opacity?: number;
  baseFlowAnimated: boolean;
}

export interface ConnectionSemanticHints {
  type?: unknown;
  style?: unknown;
  label?: unknown;
  fromCategory?: unknown;
  toCategory?: unknown;
  fromName?: unknown;
  toName?: unknown;
}

const PRESENTATIONS: Record<DiagramConnectionType, ConnectionPresentation> = {
  sync: {
    type: 'sync',
    stroke: '#64748b',
    baseFlowAnimated: true,
  },
  async: {
    type: 'async',
    stroke: '#64748b',
    strokeDasharray: '6, 5',
    baseFlowAnimated: true,
  },
  optional: {
    type: 'optional',
    stroke: '#64748b',
    strokeDasharray: '2, 4',
    opacity: 0.68,
    baseFlowAnimated: false,
  },
  security: {
    type: 'security',
    stroke: '#dc2626',
    strokeDasharray: '2, 3',
    baseFlowAnimated: false,
  },
  telemetry: {
    type: 'telemetry',
    stroke: '#7c3aed',
    strokeDasharray: '7, 3, 2, 3',
    baseFlowAnimated: true,
  },
};

export function normalizeConnectionType(value: unknown): DiagramConnectionType {
  if (typeof value !== 'string') return 'sync';
  const normalized = value.trim().toLowerCase();
  if (normalized === 'async' || normalized === 'asynchronous' || normalized === 'event') {
    return 'async';
  }
  if (normalized === 'optional' || normalized === 'fallback') return 'optional';
  if (normalized === 'security' || normalized === 'identity' || normalized === 'trust') {
    return 'security';
  }
  if (
    normalized === 'telemetry'
    || normalized === 'monitoring'
    || normalized === 'observability'
  ) {
    return 'telemetry';
  }
  return 'sync';
}

export function getConnectionPresentation(value: unknown): ConnectionPresentation {
  return PRESENTATIONS[normalizeConnectionType(value)];
}

function normalizedText(values: unknown[]): string {
  return values
    .filter((value): value is string => typeof value === 'string')
    .join(' ')
    .trim()
    .toLowerCase();
}

export function inferConnectionType(hints: ConnectionSemanticHints): DiagramConnectionType {
  if (typeof hints.type === 'string' && hints.type.trim()) {
    return normalizeConnectionType(hints.type);
  }

  const categories = normalizedText([hints.fromCategory, hints.toCategory]);
  const context = normalizedText([
    hints.label,
    hints.fromName,
    hints.toName,
    hints.fromCategory,
    hints.toCategory,
  ]);

  if (
    /\b(identity|security)\b/.test(categories)
    || /\b(auth|authenticate|authentication|authorize|authorization|token|oauth|oidc|identity|trust|secret|certificate|policy)\b/.test(context)
  ) {
    return 'security';
  }
  if (
    /\b(monitor|monitoring|observability)\b/.test(categories)
    || /\b(telemetry|metric|metrics|log|logs|trace|traces|diagnostic|diagnostics|monitor|monitoring|observe|observability)\b/.test(context)
  ) {
    return 'telemetry';
  }

  const style = typeof hints.style === 'string' ? hints.style.trim().toLowerCase() : '';
  if (style === 'dotted') return 'optional';
  if (style === 'dashed') return 'async';
  if (/\b(async|asynchronous|event|events|queue|queued|publish|subscribe|message|messages)\b/.test(context)) {
    return 'async';
  }
  return 'sync';
}
