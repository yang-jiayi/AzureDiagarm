// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

import type { CloudDiagramPayload } from '../services/cloudDiagramService';

export type SensitiveFindingKind =
  | 'credential'
  | 'connection-string'
  | 'email'
  | 'private-address'
  | 'resource-id'
  | 'internal-host';

export interface SensitiveFinding {
  id: string;
  kind: SensitiveFindingKind;
  severity: 'high' | 'medium' | 'low';
  location: string;
  preview: string;
}

interface SensitiveRule {
  kind: SensitiveFindingKind;
  severity: SensitiveFinding['severity'];
  pattern: RegExp;
  replacement: string | ((match: string, index: number) => string);
}

const RULES: SensitiveRule[] = [
  {
    kind: 'connection-string',
    severity: 'high',
    pattern: /\b(?:DefaultEndpointsProtocol|AccountKey|SharedAccessKey|SharedAccessSignature|Endpoint)\s*=\s*[^;\s]+(?:;[^;\r\n]+)*/gi,
    replacement: '[REDACTED CONNECTION STRING]',
  },
  {
    kind: 'credential',
    severity: 'high',
    pattern: /\b(?:api[_ -]?key|client[_ -]?secret|password|passwd|access[_ -]?token|bearer)\s*[:=]\s*(?:"[^"\r\n]{8,}"|'[^'\r\n]{8,}'|[^\s,;'"{}\[\]]{8,})/gi,
    replacement: '[REDACTED CREDENTIAL]',
  },
  {
    kind: 'credential',
    severity: 'high',
    pattern: /\beyJ[A-Za-z0-9_-]{10,}\.[A-Za-z0-9_-]{10,}(?:\.[A-Za-z0-9_-]{10,})?\b/g,
    replacement: '[REDACTED TOKEN]',
  },
  {
    kind: 'resource-id',
    severity: 'medium',
    pattern: /\/subscriptions\/[0-9a-f-]{36}\/resourceGroups\/[^/\s]+(?:\/providers\/[^\s,;]+)?/gi,
    replacement: '[AZURE RESOURCE ID REMOVED]',
  },
  {
    kind: 'email',
    severity: 'medium',
    pattern: /\b[A-Z0-9._%+-]+@[A-Z0-9.-]+\.[A-Z]{2,}\b/gi,
    replacement: (_match, index) => `[EMAIL REMOVED ${index + 1}]`,
  },
  {
    kind: 'private-address',
    severity: 'low',
    pattern: /\b(?:10\.(?:\d{1,3}\.){2}\d{1,3}|192\.168\.(?:\d{1,3}\.)\d{1,3}|172\.(?:1[6-9]|2\d|3[01])\.(?:\d{1,3}\.)\d{1,3})\b/g,
    replacement: '[PRIVATE ADDRESS REMOVED]',
  },
  {
    kind: 'internal-host',
    severity: 'low',
    pattern: /\b[a-z0-9][a-z0-9.-]{2,}\.(?:internal|corp|local|lan)\b/gi,
    replacement: (_match, index) => `[INTERNAL HOST REMOVED ${index + 1}]`,
  },
];

const SAFE_PREVIEWS: Record<SensitiveFindingKind, string> = {
  credential: '[CREDENTIAL REDACTED]',
  'connection-string': '[CONNECTION STRING REDACTED]',
  email: '[EMAIL REDACTED]',
  'private-address': '[PRIVATE ADDRESS REDACTED]',
  'resource-id': '[AZURE RESOURCE ID REDACTED]',
  'internal-host': '[INTERNAL HOST REDACTED]',
};

function scanValue(
  value: unknown,
  path: string,
  findings: SensitiveFinding[],
  depth = 0,
): void {
  if (depth > 10 || findings.length >= 100) return;
  if (typeof value === 'string') {
    for (const rule of RULES) {
      // RULES patterns are module-scoped and global. matchAll() operates on an
      // internal clone, so the shared regex's lastIndex is never mutated and it
      // is safe to reuse directly instead of compiling a fresh RegExp per call.
      for (const match of value.matchAll(rule.pattern)) {
        findings.push({
          id: `${rule.kind}:${path}:${match.index ?? findings.length}`,
          kind: rule.kind,
          severity: rule.severity,
          location: path,
          preview: SAFE_PREVIEWS[rule.kind],
        });
      }
    }
    return;
  }
  if (Array.isArray(value)) {
    value.forEach((item, index) => scanValue(item, `${path}[${index}]`, findings, depth + 1));
    return;
  }
  if (!value || typeof value !== 'object') return;

  for (const [key, child] of Object.entries(value as Record<string, unknown>)) {
    if (typeof child === 'function' || key === 'iconPath') continue;
    scanValue(child, path ? `${path}.${key}` : key, findings, depth + 1);
  }
}

export function detectSensitiveDataInValue(
  value: unknown,
  root = 'content',
): SensitiveFinding[] {
  const findings: SensitiveFinding[] = [];
  scanValue(value, root, findings);
  // Deduplicate by id in a single pass (Set), keeping the first occurrence and
  // preserving order — equivalent to the previous O(n^2) findIndex filter.
  const seen = new Set<string>();
  const deduped: SensitiveFinding[] = [];
  for (const finding of findings) {
    if (seen.has(finding.id)) continue;
    seen.add(finding.id);
    deduped.push(finding);
  }
  return deduped;
}

export function detectSensitiveData(payload: CloudDiagramPayload): SensitiveFinding[] {
  return detectSensitiveDataInValue(payload, 'diagram');
}

export function anonymizeSensitiveText(value: string): string {
  let next = value;
  for (const rule of RULES) {
    // Reusing the module-scoped global regex is safe: String.replace resets a
    // global pattern's lastIndex to 0 before scanning.
    if (typeof rule.replacement === 'string') {
      next = next.replace(rule.pattern, rule.replacement);
      continue;
    }

    let replacementIndex = 0;
    const replaceMatch = rule.replacement;
    next = next.replace(rule.pattern, match => replaceMatch(match, replacementIndex++));
  }
  return next;
}

function anonymizeValue(
  value: unknown,
  aliases: ReadonlyMap<string, string>,
  depth = 0,
): unknown {
  if (depth > 10) return value;
  if (typeof value === 'string') {
    return aliases.get(value) ?? anonymizeSensitiveText(value);
  }
  if (Array.isArray(value)) {
    return value.map(item => anonymizeValue(item, aliases, depth + 1));
  }
  if (!value || typeof value !== 'object') return value;

  const output: Record<string, unknown> = {};
  for (const [key, child] of Object.entries(value as Record<string, unknown>)) {
    if (typeof child === 'function') continue;
    output[key] = anonymizeValue(child, aliases, depth + 1);
  }
  return output;
}

function createStructuralAliases(payload: CloudDiagramPayload): Map<string, string> {
  const nodeIds = payload.nodes
    .map(node => node?.id)
    .filter((id): id is string => typeof id === 'string');
  const edgeIds = payload.edges
    .map(edge => edge?.id)
    .filter((id): id is string => typeof id === 'string');
  const usedIds = new Set([...nodeIds, ...edgeIds]);
  const aliases = new Map<string, string>();
  let nodeIndex = 0;
  let edgeIndex = 0;

  const addAlias = (id: string, prefix: 'node' | 'edge') => {
    if (aliases.has(id) || anonymizeSensitiveText(id) === id) return;
    let alias: string;
    do {
      const index = prefix === 'node' ? ++nodeIndex : ++edgeIndex;
      alias = `${prefix}-anonymized-${index}`;
    } while (usedIds.has(alias));
    usedIds.add(alias);
    aliases.set(id, alias);
  };

  nodeIds.forEach(id => addAlias(id, 'node'));
  edgeIds.forEach(id => addAlias(id, 'edge'));
  return aliases;
}

export function anonymizeDiagramPayload(payload: CloudDiagramPayload): CloudDiagramPayload {
  return anonymizeValue(payload, createStructuralAliases(payload)) as CloudDiagramPayload;
}
