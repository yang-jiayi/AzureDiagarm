/**
 * Classifies an AI-supplied connection label into an arrow direction.
 *
 * Extracted from App so it can be tested directly. Matching is done per word
 * rather than per substring: substring matching found "ack" inside "backup",
 * "rollback", "feedback", "track" and "package", and "sync" inside "async", so
 * an edge labelled "Daily backup" rendered its arrowhead at the wrong end.
 * That is not only cosmetic — the IaC exporter derives deployment ordering
 * from the same field, so a misclassified label also produced an inverted
 * dependsOn.
 *
 * Words are matched against stems, not exact spellings, because the label is
 * free text from a model: "Synchronization", "Responses" and "Communicates
 * with" are as likely as the bare verb, and treating them as unmatched
 * reintroduces the same wrong-way arrow from the other side.
 */
export type EdgeDirection = 'forward' | 'reverse' | 'bidirectional';

/**
 * `repl` is deliberately not a stem: "replicate" and "replication" are
 * extremely common on Azure diagrams and mean the opposite of a reply.
 * "synchronous" is likewise excluded — the app's own connection legend
 * defines it as request/response, not a two-way sync.
 */
const REVERSE_WORDS = [
  /^responses?$/,
  /^callbacks?$/,
  /^return(s|ed|ing)?$/,
  /^acknowledg(e|es|ed|ing|ement|ements|ment|ments)$/,
  /^acks?$/,
  /^repl(y|ies|ied|ying)$/,
];

const BIDIRECTIONAL_WORDS = [
  /^syncs?$/,
  /^synchroni[sz](e|es|ed|ing|ation|ations)$/,
  /^exchang(e|es|ed|ing)$/,
  /^communicat(e|es|ed|ing|ion|ions)$/,
  /^bidirectional(ly)?$/,
];

/** Matched against the whole label because they span a word boundary. */
const BIDIRECTIONAL_PHRASES = ['two-way', 'two way', '2-way', '2 way'];

/**
 * Splits camelCase before lowercasing so "sendResponse" and "onCallback"
 * tokenise, then breaks on every non-alphanumeric run so "request/response"
 * and "sync-config" do too.
 */
function tokenize(label: string): string[] {
  return label
    .replace(/([a-z0-9])([A-Z])/g, '$1 $2')
    .toLowerCase()
    .split(/[^a-z0-9]+/)
    .filter(Boolean);
}

export function classifyEdgeDirection(label: string): EdgeDirection {
  const raw = (label || '').toLowerCase();
  const words = tokenize(label || '');

  if (BIDIRECTIONAL_PHRASES.some((phrase) => raw.includes(phrase))) return 'bidirectional';
  if (words.some((word) => BIDIRECTIONAL_WORDS.some((pattern) => pattern.test(word)))) return 'bidirectional';
  if (words.some((word) => REVERSE_WORDS.some((pattern) => pattern.test(word)))) return 'reverse';
  return 'forward';
}
