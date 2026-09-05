import type { ArchitectureValidation, ValidationFinding } from './architectureValidator';

export type ReviewFinding = ValidationFinding;

export interface ValidationReviewRecord {
  key: string;
  pillar: string;
  finding: ReviewFinding;
  status: 'active' | 'not-detected';
  firstSeenAt: number;
  lastSeenAt: number;
  lastReviewedAt: number;
  notDetectedAt?: number;
}

export const VALIDATION_REVIEW_LIMITS = {
  records: 2000,
  text: 16000,
  identifier: 1000,
  resources: 200,
  resourceName: 512,
  details: 8,
  detailText: 800,
  referenceUrl: 500,
  actionText: 160,
  key: 1048576,
  totalText: 4 * 1048576,
} as const;
const MAX_TIMESTAMP = 8640000000000000;

const canonicalPillars: Record<string, string> = {
  'cost optimization': 'cost', 'コストの最適化': 'cost', 'コスト最適化': 'cost',
  'operational excellence': 'operations', 'オペレーショナル エクセレンス': 'operations',
  '運用上の優秀性': 'operations', '運用の優秀性': 'operations',
  'performance efficiency': 'performance', 'パフォーマンス効率': 'performance',
  reliability: 'reliability', '信頼性': 'reliability',
  security: 'security', 'セキュリティ': 'security',
};

function normalized(value: string): string {
  return value.normalize('NFKC').trim().replace(/\s+/g, ' ').toLowerCase();
}

function identity(finding: ReviewFinding, pillar: string): string {
  const scope = finding.resourceIds?.length
    ? ['ids', [...finding.resourceIds].sort()]
    : ['labels', [...(finding.resources ?? [])].map(normalized).sort()];
  // Never infer identity from similar wording or a repeated resource label.
  // A missing resource scope is distinct from an explicitly targeted finding.
  const stableId = finding.findingId || finding.id || finding.ruleId;
  return JSON.stringify(stableId
    ? ['stable', finding.findingId ? 'findingId' : finding.id ? 'id' : 'ruleId',
      stableId, scope]
    : ['exact', canonicalPillars[normalized(pillar)] ?? normalized(pillar),
      finding.source ?? 'unknown', normalized(finding.category),
      normalized(finding.issue), normalized(finding.recommendation), scope]);
}

export function updateValidationReview(
  previous: ValidationReviewRecord[],
  validation: ArchitectureValidation,
  now: number = Date.parse(validation?.timestamp ?? ''),
): ValidationReviewRecord[] {
  validateValidationReviewHistory(previous);
  const timestamp = Number.isFinite(now) && now >= 0 && now <= MAX_TIMESTAMP ? now : Date.now();
  if (!validation || !Array.isArray(validation.pillars) || !Array.isArray(validation.quickWins) ||
      validation.pillars.length > VALIDATION_REVIEW_LIMITS.records ||
      Array.from(validation.pillars).some(pillar => !isObject(pillar) ||
        !boundedString(pillar.pillar, VALIDATION_REVIEW_LIMITS.identifier) || !Array.isArray(pillar.findings))) {
    throw new Error('Invalid validation review findings or review size limit exceeded');
  }
  const incomingCount = validation.pillars.reduce((count, pillar) => count + pillar.findings.length, 0) +
    validation.quickWins.length;
  if (incomingCount > VALIDATION_REVIEW_LIMITS.records ||
      !validation.pillars.every(pillar => Array.from(pillar.findings).every(isReviewFinding)) ||
      !Array.from(validation.quickWins).every(isReviewFinding)) {
    throw new Error('Invalid validation review findings or review size limit exceeded');
  }
  const records = new Map<string, ValidationReviewRecord>(previous.map(record => [record.key, {
    key: record.key, pillar: record.pillar,
    firstSeenAt: record.firstSeenAt, lastSeenAt: record.lastSeenAt,
    finding: copyFinding(record.finding),
    status: 'not-detected' as ValidationReviewRecord['status'],
    lastReviewedAt: Math.max(record.lastReviewedAt, timestamp),
    notDetectedAt: record.notDetectedAt ?? Math.max(record.lastReviewedAt, timestamp),
  }]));
  const occurrences = new Map<string, number>();
  const findings: { pillar: string; finding: ReviewFinding }[] = validation.pillars.flatMap(p => p.findings.map(finding => ({
    pillar: p.pillar, finding: copyFinding(finding as ReviewFinding),
  })));
  const pillarFindings = [...findings];
  // Quick wins often repeat a pillar finding; do not count that projection twice.
  const pillarIdentities = new Set(findings.map(({ finding, pillar }) => identity(finding, pillar)));
  for (const finding of validation.quickWins) {
    const reviewFinding = copyFinding(finding as ReviewFinding);
    const isProjection = pillarFindings.some(entry =>
      JSON.stringify(entry.finding) === JSON.stringify(reviewFinding));
    if (!isProjection && !pillarIdentities.has(identity(reviewFinding, 'quick-wins'))) {
      findings.push({ pillar: 'quick-wins', finding: reviewFinding });
    }
  }
  for (const { pillar, finding } of findings) {
    const base = identity(finding, pillar);
    const occurrence = (occurrences.get(base) ?? 0) + 1;
    occurrences.set(base, occurrence);
    const key = `${base}#${occurrence}`;
    const prior = records.get(key);
    const observedAt = Math.max(prior?.lastReviewedAt ?? 0, timestamp);
    records.set(key, {
      key, pillar,
      finding: copyFinding(finding),
      status: 'active',
      firstSeenAt: prior?.firstSeenAt ?? observedAt,
      lastSeenAt: observedAt,
      lastReviewedAt: observedAt,
    });
  }
  return validateValidationReviewHistory([...records.values()]);
}

function isObject(value: unknown): value is Record<string, unknown> {
  return typeof value === 'object' && value !== null && !Array.isArray(value);
}

function boundedString(value: unknown, maximum: number): value is string {
  return typeof value === 'string' && value.length <= maximum;
}

function validStrings(value: unknown): value is string[] {
  if (!Array.isArray(value) || value.length > VALIDATION_REVIEW_LIMITS.resources) return false;
  for (const item of value) {
    if (!boundedString(item, VALIDATION_REVIEW_LIMITS.resourceName)) return false;
  }
  return true;
}

function validDetails(value: unknown): value is string[] {
  return Array.isArray(value) && value.length <= VALIDATION_REVIEW_LIMITS.details &&
    Array.from(value).every(item => boundedString(item, VALIDATION_REVIEW_LIMITS.detailText));
}

function validAction(value: unknown): value is NonNullable<ReviewFinding['applyAction']> {
  return isObject(value) && typeof value.type === 'string' &&
    ['add-service', 'regenerate', 'configure'].includes(value.type) &&
    boundedString(value.label, VALIDATION_REVIEW_LIMITS.actionText) &&
    (value.serviceType === undefined || boundedString(value.serviceType, VALIDATION_REVIEW_LIMITS.actionText));
}

function isReviewFinding(finding: unknown): finding is ReviewFinding {
  return isObject(finding) && typeof finding.severity === 'string' &&
      ['critical', 'high', 'medium', 'low'].includes(finding.severity) &&
      boundedString(finding.category, VALIDATION_REVIEW_LIMITS.identifier) &&
      ['issue', 'recommendation'].every(key => boundedString(finding[key], VALIDATION_REVIEW_LIMITS.text)) &&
      ['id', 'findingId', 'ruleId'].every(key => finding[key] === undefined ||
        boundedString(finding[key], VALIDATION_REVIEW_LIMITS.identifier)) &&
      ['resources', 'resourceIds'].every(key => finding[key] === undefined || validStrings(finding[key])) &&
      ['evidence', 'remediation'].every(key => finding[key] === undefined || validDetails(finding[key])) &&
      (finding.referenceUrl === undefined || boundedString(finding.referenceUrl, VALIDATION_REVIEW_LIMITS.referenceUrl)) &&
      (finding.applyAction === undefined || validAction(finding.applyAction)) &&
      (finding.source === undefined || (typeof finding.source === 'string' && ['rule-based', 'ai'].includes(finding.source)));
}

function copyFinding(finding: ReviewFinding): ReviewFinding {
  return {
    severity: finding.severity, category: finding.category,
    issue: finding.issue, recommendation: finding.recommendation,
    ...(finding.id !== undefined ? { id: finding.id } : {}),
    ...(finding.findingId !== undefined ? { findingId: finding.findingId } : {}),
    ...(finding.ruleId !== undefined ? { ruleId: finding.ruleId } : {}),
    ...(finding.source !== undefined ? { source: finding.source } : {}),
    ...(finding.resources ? { resources: [...finding.resources] } : {}),
    ...(finding.resourceIds ? { resourceIds: [...finding.resourceIds] } : {}),
    ...(finding.evidence ? { evidence: [...finding.evidence] } : {}),
    ...(finding.remediation ? { remediation: [...finding.remediation] } : {}),
    ...(finding.referenceUrl !== undefined ? { referenceUrl: finding.referenceUrl } : {}),
    ...(finding.applyAction ? { applyAction: {
      type: finding.applyAction.type, label: finding.applyAction.label,
      ...(finding.applyAction.serviceType !== undefined ? { serviceType: finding.applyAction.serviceType } : {}),
    } } : {}),
  };
}

export function isValidationReviewRecord(value: unknown): value is ValidationReviewRecord {
  if (!isObject(value) || !boundedString(value.key, VALIDATION_REVIEW_LIMITS.key) || !value.key ||
      !boundedString(value.pillar, VALIDATION_REVIEW_LIMITS.identifier) || !isReviewFinding(value.finding)) return false;
  const times = [value.firstSeenAt, value.lastSeenAt, value.lastReviewedAt];
  if (!times.every(time => typeof time === 'number' && Number.isFinite(time) && time >= 0 && time <= MAX_TIMESTAMP)) return false;
  if ((value.firstSeenAt as number) > (value.lastSeenAt as number) ||
      (value.lastSeenAt as number) > (value.lastReviewedAt as number)) return false;
  if (value.status === 'active') return value.notDetectedAt === undefined;
  return value.status === 'not-detected' && typeof value.notDetectedAt === 'number' &&
    Number.isFinite(value.notDetectedAt) && value.notDetectedAt >= (value.lastSeenAt as number) &&
    value.notDetectedAt <= (value.lastReviewedAt as number);
}

export function isValidationReviewHistory(value: unknown): value is ValidationReviewRecord[] {
  if (!Array.isArray(value) || value.length > VALIDATION_REVIEW_LIMITS.records) return false;
  const keys = new Set<string>();
  let textLength = 0;
  for (const record of value) {
    if (!isValidationReviewRecord(record) || keys.has(record.key)) return false;
    keys.add(record.key);
    const finding = record.finding;
    textLength += [record.key, record.pillar, finding.category, finding.issue, finding.recommendation,
      finding.id ?? '', finding.findingId ?? '', finding.ruleId ?? '',
      ...(finding.resources ?? []), ...(finding.resourceIds ?? []),
      ...(finding.evidence ?? []), ...(finding.remediation ?? []),
      finding.referenceUrl ?? '', finding.applyAction?.label ?? '', finding.applyAction?.serviceType ?? '',
    ].reduce((length, text) => length + text.length, 0);
    if (textLength > VALIDATION_REVIEW_LIMITS.totalText) return false;
  }
  return true;
}

export function validateValidationReviewHistory(value: unknown): ValidationReviewRecord[] {
  if (!isValidationReviewHistory(value)) throw new Error('Invalid validation review history');
  return value;
}

/** Bounded restore: missing history is empty; malformed history throws; unknown fields are discarded. */
export function parseValidationReview(value: unknown): ValidationReviewRecord[] {
  if (value === undefined || value === null) return [];
  return validateValidationReviewHistory(value).map(record => ({
    key: record.key, pillar: record.pillar, finding: copyFinding(record.finding),
    status: record.status,
    firstSeenAt: record.firstSeenAt, lastSeenAt: record.lastSeenAt,
    lastReviewedAt: record.lastReviewedAt,
    ...(record.notDetectedAt !== undefined ? { notDetectedAt: record.notDetectedAt } : {}),
  }));
}
