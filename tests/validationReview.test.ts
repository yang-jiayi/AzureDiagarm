import test from 'node:test';
import assert from 'node:assert/strict';
import {
  updateValidationReview, isValidationReviewRecord, isValidationReviewHistory,
  validateValidationReviewHistory, parseValidationReview, VALIDATION_REVIEW_LIMITS, type ReviewFinding,
} from '../src/services/validationReview';
import type { ArchitectureValidation } from '../src/services/architectureValidator';

const finding = (overrides: Partial<ReviewFinding> = {}): ReviewFinding => ({
  severity: 'high', category: 'Availability', issue: 'Single instance',
  recommendation: 'Add redundancy', resources: ['Web'], source: 'ai', ...overrides,
});
const validation = (findings: ReviewFinding[], pillar = 'Reliability'): ArchitectureValidation => ({
  overallScore: 70, summary: 'Review',
  pillars: [{ pillar: pillar as ArchitectureValidation['pillars'][number]['pillar'], score: 70, findings }],
  quickWins: [], timestamp: '2026-09-01T00:00:00Z',
});

test('exact findings retain identity and first-seen timestamps without mutating history', () => {
  const first = updateValidationReview([], validation([finding()]), 100);
  const original = JSON.stringify(first);
  const next = updateValidationReview(first, validation([finding()]), 200);
  assert.equal(next[0].key, first[0].key);
  assert.equal(next[0].firstSeenAt, 100);
  assert.equal(next[0].lastSeenAt, 200);
  assert.equal(next[0].lastReviewedAt, 200);
  assert.equal(JSON.stringify(first), original);
  assert.ok(isValidationReviewHistory(next));
});

test('no-longer-detected is not a remediation claim; repeated absence preserves transition date', () => {
  const first = updateValidationReview([], validation([finding()]), 100);
  const absent = updateValidationReview(first, validation([]), 200);
  const stillAbsent = updateValidationReview(absent, validation([]), 300);
  assert.equal(stillAbsent[0].status, 'not-detected');
  assert.equal(stillAbsent[0].notDetectedAt, 200);
  assert.equal(stillAbsent[0].lastSeenAt, 100);
  assert.equal(stillAbsent[0].lastReviewedAt, 300);
  assert.ok(!('resolvedAt' in stillAbsent[0]));
});

test('reopened finding preserves first-seen timestamp and clears absence marker', () => {
  const first = updateValidationReview([], validation([finding({ ruleId: 'REL-1' })]), 100);
  const absent = updateValidationReview(first, validation([]), 200);
  const reopened = updateValidationReview(absent, validation([finding({
    ruleId: 'REL-1', issue: 'A new paraphrase of the same rule',
  })]), 300);
  assert.equal(reopened.length, 1);
  assert.equal(reopened[0].status, 'active');
  assert.equal(reopened[0].firstSeenAt, 100);
  assert.equal(reopened[0].lastSeenAt, 300);
  assert.equal(reopened[0].notDetectedAt, undefined);
});

test('paraphrases without stable identifiers are conservatively separate', () => {
  const first = updateValidationReview([], validation([finding()]), 100);
  const next = updateValidationReview(first, validation([finding({ issue: 'Redundancy is missing' })]), 200);
  assert.equal(next.length, 2);
  assert.deepEqual(next.map(record => record.status), ['not-detected', 'active']);
});

test('stable identifiers with resource IDs survive translated findings, pillars and labels', () => {
  const first = updateValidationReview([], validation([finding({
    findingId: 'finding-1', resourceIds: ['node-1'],
  })]), 100);
  const next = updateValidationReview(first, validation([finding({
    findingId: 'finding-1', resourceIds: ['node-1'], resources: ['ウェブ'],
    category: '可用性', issue: '単一のインスタンス', recommendation: '冗長性を追加',
  })], '信頼性'), 200);
  assert.equal(next.length, 1);
  assert.equal(next[0].firstSeenAt, 100);
  assert.equal(next[0].finding.issue, '単一のインスタンス');
});

test('translated findings without stable identifiers never imply the original was fixed', () => {
  const first = updateValidationReview([], validation([finding()]), 100);
  const next = updateValidationReview(first, validation([finding({ issue: '単一のインスタンス' })], '信頼性'), 200);
  assert.equal(next.length, 2);
  assert.equal(next[0].status, 'not-detected');
});

test('duplicate labels with distinct resource IDs are distinct rule instances', () => {
  const first = updateValidationReview([], validation([
    finding({ ruleId: 'REL-1', resourceIds: ['node-1'] }),
    finding({ ruleId: 'REL-1', resourceIds: ['node-2'] }),
  ]), 100);
  const next = updateValidationReview(first, validation([
    finding({ ruleId: 'REL-1', resourceIds: ['node-2'] }),
  ]), 200);
  assert.equal(next.length, 2);
  assert.equal(next[0].status, 'not-detected');
  assert.equal(next[1].status, 'active');
});

test('identical duplicate findings do not collapse, and a missing resource does not match', () => {
  const duplicate = updateValidationReview([], validation([finding(), finding()]), 100);
  assert.equal(duplicate.length, 2);
  assert.notEqual(duplicate[0].key, duplicate[1].key);
  assert.ok(isValidationReviewHistory(duplicate));
  const scoped = updateValidationReview([], validation([finding({ ruleId: 'REL-1' })]), 100);
  const unscoped = updateValidationReview(scoped, validation([finding({ ruleId: 'REL-1', resources: undefined })]), 200);
  assert.equal(unscoped.length, 2);
  assert.equal(unscoped[0].status, 'not-detected');
});

test('resource order is irrelevant, quick-win projections do not double-count', () => {
  const first = updateValidationReview([], validation([finding({ resources: ['Web', 'DB'] })]), 100);
  const result = validation([finding({ resources: ['DB', 'Web'] })]);
  result.quickWins = [...result.pillars[0].findings];
  assert.equal(updateValidationReview(first, result, 200).length, 1);
});

test('restore validators reject malformed records, duplicate keys and invalid timestamps', () => {
  const records = updateValidationReview([], validation([finding()]), 100);
  assert.equal(validateValidationReviewHistory(records), records);
  for (const invalid of [
    null, {}, [{ id: 'old-record' }], [...records, ...records],
    [{ ...records[0], firstSeenAt: Infinity }],
    [{ ...records[0], lastSeenAt: -1 }],
    [{ ...records[0], status: 'resolved' }],
    [{ ...records[0], status: 'not-detected' }],
    [{ ...records[0], finding: { ...records[0].finding, resources: [123] } }],
  ]) {
    assert.equal(isValidationReviewHistory(invalid), false);
    assert.throws(() => validateValidationReviewHistory(invalid), /Invalid/);
  }
  assert.equal(isValidationReviewRecord(records[0]), true);
  assert.deepEqual(parseValidationReview(undefined), []);
  assert.deepEqual(parseValidationReview(null), []);
  assert.deepEqual(parseValidationReview(records), records);
  assert.throws(() => parseValidationReview([{ id: 'legacy-unrecognized' }]), /Invalid/);
});

test('restore rejects oversized histories, fields, resource lists, sparse arrays and invalid date ranges', () => {
  const [record] = updateValidationReview([], validation([finding()]), 100);
  for (const invalid of [
    Array(VALIDATION_REVIEW_LIMITS.records + 1).fill(record),
    new Array(1),
    [{ ...record, key: 'x'.repeat(VALIDATION_REVIEW_LIMITS.key + 1) }],
    [{ ...record, finding: { ...record.finding, issue: 'x'.repeat(VALIDATION_REVIEW_LIMITS.text + 1) } }],
    [{ ...record, finding: { ...record.finding, resources: Array(VALIDATION_REVIEW_LIMITS.resources + 1).fill('Web') } }],
    [{ ...record, finding: { ...record.finding, resourceIds: new Array(1) } }],
    [{ ...record, lastReviewedAt: Number.MAX_VALUE }],
  ]) {
    assert.equal(isValidationReviewHistory(invalid), false);
    assert.throws(() => parseValidationReview(invalid), /Invalid/);
  }
});

test('restore enforces a total text budget and strips unknown nested data without mutating input', () => {
  const [record] = updateValidationReview([], validation([finding()]), 100);
  const manyLargeRecords = Array.from({ length: 5 }, (_, index) => ({
    ...record, key: String(index) + 'x'.repeat(VALIDATION_REVIEW_LIMITS.key - 1),
  }));
  assert.throws(() => parseValidationReview(manyLargeRecords), /Invalid/);
  const extra = { ...record, hidden: { arbitrary: 'data' }, finding: { ...record.finding, unexpected: ['metadata'] } };
  const [restored] = parseValidationReview([extra]);
  assert.deepEqual(restored, record);
  assert.notEqual(restored.finding, record.finding);
  assert.notEqual(restored.finding.resources, record.finding.resources);
  assert.ok('hidden' in extra);
  assert.ok(!('hidden' in restored));
});

test('repeated or older timestamps never move observation time backwards', () => {
  const first = updateValidationReview([], validation([finding()]), 200);
  const next = updateValidationReview(first, validation([finding()]), 100);
  assert.equal(next[0].lastSeenAt, 200);
  assert.ok(isValidationReviewHistory(next));
});

test('remote evidence, remediation, links and actions survive bounded cloned history', () => {
  const detailed = finding({
    ruleId: 'REL-1', evidence: ['Only one instance is shown.'],
    remediation: ['Add a second instance.'],
    referenceUrl: 'https://learn.microsoft.com/azure/well-architected/reliability/',
    applyAction: { type: 'add-service', label: 'Add redundancy', serviceType: 'App Service' },
  });
  const records = updateValidationReview([], validation([detailed]), 100);
  const restored = parseValidationReview(records);
  assert.deepEqual(restored[0].finding, detailed);
  assert.notEqual(restored[0].finding.evidence, detailed.evidence);
  assert.notEqual(restored[0].finding.remediation, detailed.remediation);
  assert.notEqual(restored[0].finding.applyAction, detailed.applyAction);
  restored[0].finding.evidence![0] = 'Changed';
  restored[0].finding.applyAction!.label = 'Changed';
  assert.equal(records[0].finding.evidence![0], 'Only one instance is shown.');
  assert.equal(detailed.applyAction!.label, 'Add redundancy');
  const absent = updateValidationReview(records, validation([]), 200);
  assert.equal(absent[0].status, 'not-detected');
  assert.deepEqual(absent[0].finding.remediation, detailed.remediation);
  assert.ok(!('resolvedAt' in absent[0]));
});

test('remote finding details obey field and aggregate budgets rather than being discarded', () => {
  const [record] = updateValidationReview([], validation([finding()]), 100);
  for (const detail of [
    { evidence: ['x'.repeat(VALIDATION_REVIEW_LIMITS.detailText + 1)] },
    { remediation: Array(VALIDATION_REVIEW_LIMITS.details + 1).fill('step') },
    { evidence: new Array(1) },
    { referenceUrl: 'x'.repeat(VALIDATION_REVIEW_LIMITS.referenceUrl + 1) },
    { applyAction: { type: 'unknown', label: 'Bad action' } },
    { applyAction: { type: 'configure', label: 'x'.repeat(VALIDATION_REVIEW_LIMITS.actionText + 1) } },
  ]) {
    assert.throws(() => parseValidationReview([{ ...record, finding: { ...record.finding, ...detail } }]), /Invalid/);
  }
  const crowded = Array.from({ length: 400 }, (_, index) => ({
    ...record, key: String(index),
    finding: {
      ...record.finding, evidence: Array(8).fill('x'.repeat(800)),
      remediation: Array(8).fill('x'.repeat(800)),
    },
  }));
  assert.throws(() => parseValidationReview(crowded), /Invalid/);
});

test('malformed or sparse reports never turn existing findings into absence records', () => {
  const first = updateValidationReview([], validation([finding()]), 100);
  const original = JSON.stringify(first);
  for (const invalid of [
    null, {}, { ...validation([]), pillars: [null] },
    { ...validation([]), pillars: new Array(1) },
    { ...validation([]), quickWins: new Array(1) },
    { ...validation([]), pillars: [{ pillar: 'Security', findings: new Array(1) }] },
  ]) {
    assert.throws(() => updateValidationReview(first, invalid as never, 200), /Invalid/);
    assert.equal(JSON.stringify(first), original);
  }
});
