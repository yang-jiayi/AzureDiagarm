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

/**
 * The per-edge animation intent, independent of the global toggle.
 *
 * `flowAnimated` on a live edge is a *product*: the user's intent for this
 * edge AND the app-wide "animate connections" switch. Only the product is
 * serialised, and the switch itself lives in localStorage rather than the
 * file, so save-time and open-time global state legitimately differ. Reading
 * the stored product back as if it were the intent is what made a diagram
 * saved with animation off come back permanently static -- every edge carried
 * `flowAnimated: false`, which masked its own `baseFlowAnimated: true`.
 *
 * `baseFlowAnimated` is the intent, and it is the only field a restore may
 * trust. The stored product is read only for files written before the field
 * existed, where it is the best evidence available.
 *
 * One consequence is unavoidable, and it only affects files saved by older
 * builds. Those builds recorded a per-edge pause solely in the product, leaving
 * the intent `true`, so `{ baseFlowAnimated: true, flowAnimated: false }` means
 * *either* "the user paused this edge" *or* "this was saved while the switch
 * was off" -- the same bytes, and the switch was never written to the file, so
 * nothing distinguishes them. Honouring the product to rescue the old pause is
 * exactly what caused the bug above. Trusting the intent is the right way round;
 * an individually paused edge in a pre-existing file comes back animated once.
 */
export function edgeAnimationIntent(
  data: { baseFlowAnimated?: unknown; flowAnimated?: unknown } | undefined,
  fallback: boolean,
): boolean {
  if (typeof data?.baseFlowAnimated === 'boolean') return data.baseFlowAnimated;
  if (typeof data?.flowAnimated === 'boolean') return data.flowAnimated;
  return fallback;
}

/** Whether an edge animates right now: its own intent, gated by the global switch. */
export function resolveFlowAnimated(
  data: { baseFlowAnimated?: unknown; flowAnimated?: unknown } | undefined,
  fallback: boolean,
  animationsEnabled: boolean,
): boolean {
  return animationsEnabled && edgeAnimationIntent(data, fallback);
}