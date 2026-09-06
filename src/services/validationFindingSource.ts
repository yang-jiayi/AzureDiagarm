import type { ValidationFinding } from './architectureValidator';

export function normalizeValidationFindingSource<T extends Omit<ValidationFinding, 'source'> & { source?: unknown }>(
  finding: T,
): Omit<T, 'source'> & Pick<ValidationFinding, 'source'> {
  if (!finding || typeof finding !== 'object' || Array.isArray(finding)) {
    throw new Error('Invalid validation finding');
  }
  if (finding.source === undefined) return { ...finding, source: undefined };
  if (finding.source === 'ai-analysis' || finding.source === 'ai') return { ...finding, source: 'ai' };
  if (finding.source === 'rule-based') return { ...finding, source: 'rule-based' };
  throw new Error('Invalid validation finding source');
}
