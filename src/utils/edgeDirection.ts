/**
 * Classifies an AI-supplied connection label into an arrow direction.
 *
 * Extracted from App so it can be tested directly: the previous substring
 * matching read "ack" out of "backup", "rollback", "feedback", "track" and
 * "package", and "sync" out of "async", so an edge labelled "Daily backup"
 * rendered its arrowhead at the wrong end. That is not only cosmetic — the
 * IaC exporter derives deployment ordering from the same field, so a
 * misclassified label also produced an inverted dependsOn.
 */
export type EdgeDirection = 'forward' | 'reverse' | 'bidirectional';

const REVERSE_KEYWORDS = ['response', 'callback', 'return', 'returns', 'acknowledge', 'ack', 'reply'];
const BIDIRECTIONAL_KEYWORDS = ['sync', 'bidirectional', 'two-way', 'exchange', 'communicate'];

function matches(label: string, keywords: string[]): boolean {
  const words = new Set(label.split(/[^a-z0-9]+/).filter(Boolean));
  return keywords.some((keyword) => (
    // Multi-word keywords cannot be looked up in the word set.
    /[^a-z0-9]/.test(keyword) ? label.includes(keyword) : words.has(keyword)
  ));
}

export function classifyEdgeDirection(label: string): EdgeDirection {
  const lowerLabel = (label || '').toLowerCase();
  if (matches(lowerLabel, BIDIRECTIONAL_KEYWORDS)) return 'bidirectional';
  if (matches(lowerLabel, REVERSE_KEYWORDS)) return 'reverse';
  return 'forward';
}
