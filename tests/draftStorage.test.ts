import test from 'node:test';
import assert from 'node:assert/strict';
import { draftFingerprint, validateDraft, validateRestoredNodes, validateRestoredEdges, type DiagramDraft } from '../src/services/draftStorage';
import { editorFingerprint } from '../src/services/editorHistory';
import { parseValidationReview, updateValidationReview } from '../src/services/validationReview';
import type { ArchitectureValidation } from '../src/services/architectureValidator';

function draft(): DiagramDraft {
  return {
    id: 'test', schemaVersion: 1, revision: 1, updatedAt: 1,
    document: {
      nodes: [{ id: 'one', type: 'azureNode', position: { x: 0, y: 0 }, data: { label: 'One' } }],
      edges: [], workflow: [], architecturePrompt: 'original', originalPrompt: 'original',
      titleBlockData: { architectureName: 'App', author: '', date: '', version: '1' },
      settings: { pricingMode: 'payg', stylePreset: 'detailed', edgeStyle: 'orthogonal' },
    },
  };
}

test('drafts retain document settings, history, and original prompts', () => {
  const stored = draft();
  stored.document.settings.pricingRegion = 'japaneast';
  stored.document.reviewHistory = updateValidationReview([], {
    overallScore: 80, summary: 'Review', timestamp: '2026-01-01T00:00:00Z',
    pillars: [{ pillar: 'Security', score: 80, findings: [{
      id: 'security-check', severity: 'high', category: 'Identity', issue: 'Missing authentication', recommendation: 'Enable it',
    }] }],
    quickWins: [],
  });
  assert.deepEqual(validateDraft(stored, 'test'), stored);
  const updated = draft();
  assert.notEqual(draftFingerprint(stored.document), draftFingerprint(updated.document));
});

test('restoring unchanged WAF history never creates a draft edit, but new review observations do', () => {
  const report: ArchitectureValidation = {
    overallScore: 80, summary: 'Review', timestamp: '2026-01-01T00:00:00Z',
    pillars: [{
      pillar: 'Security', score: 80,
      findings: [{ severity: 'high', category: 'Identity', issue: 'Missing authentication', recommendation: 'Enable authentication', resources: ['One'] }],
    }],
    quickWins: [],
  };
  const active = updateValidationReview([], report);
  const notDetected = updateValidationReview(active, { ...report, pillars: [], timestamp: '2026-01-02T00:00:00Z' });
  const stored = { ...draft().document, reviewHistory: notDetected };
  const restored = { ...stored, reviewHistory: parseValidationReview(notDetected) };
  assert.deepEqual(restored.reviewHistory, stored.reviewHistory);
  assert.equal(draftFingerprint(restored), draftFingerprint(stored));
  assert.equal(draftFingerprint({ ...stored, reviewHistory: undefined }), draftFingerprint({ ...stored, reviewHistory: [] }));

  const reobserved = { ...restored, reviewHistory: updateValidationReview(restored.reviewHistory, { ...report, timestamp: '2026-01-03T00:00:00Z' }) };
  assert.notEqual(draftFingerprint(restored), draftFingerprint(reobserved));
  assert.equal(editorFingerprint(restored), editorFingerprint(reobserved), 'Review-only changes remain outside editing undo history.');
});

test('malformed and unsupported drafts fail explicitly', () => {
  assert.throws(() => validateDraft({ ...draft(), schemaVersion: 2 }, 'test'), /unsupported/);
  assert.throws(() => validateDraft(draft(), 'another-user'), /unsupported/);
  assert.throws(() => validateDraft({ ...draft(), revision: -1 }, 'test'), /unsupported/);
  const invalidRegion = draft();
  invalidRegion.document.settings.pricingRegion = '../invalid';
  assert.throws(() => validateDraft(invalidRegion, 'test'), /pricing region/);
  const bad = draft();
  bad.document.nodes[0].position.x = Infinity;
  assert.throws(() => validateDraft(bad, 'test'), /invalid graph/);
});

test('duplicate IDs, dangling edges and cyclic groups are rejected on recovery', () => {
  const stored = draft();
  stored.document.nodes.push({ ...stored.document.nodes[0] });
  assert.throws(() => validateDraft(stored, 'test'), /duplicate/);
  stored.document.nodes.pop();
  stored.document.edges.push({ id: 'bad', source: 'one', target: 'missing' });
  assert.throws(() => validateDraft(stored, 'test'), /dangling/);
  stored.document.edges = [];
  stored.document.nodes[0].parentNode = 'two';
  stored.document.nodes.push({ id: 'two', type: 'groupNode', parentNode: 'one', position: { x: 0, y: 0 }, data: {} });
  assert.throws(() => validateDraft(stored, 'test'), /cyclic/);
});

test('nullable prices preserve production quantities through 100,000 and reject larger values', () => {
  const stored = draft();
  stored.document.nodes[0].data.pricing = { estimatedCost: null, quantity: 100_000, region: 'eastus2' };
  const restored = validateDraft(stored, 'test');
  assert.equal(restored.document.nodes[0].data.pricing.estimatedCost, null, 'Unpriced must never become free.');
  assert.equal(restored.document.nodes[0].data.pricing.quantity, 100_000);
  stored.document.nodes[0].data.pricing.quantity = 100_001;
  assert.throws(() => validateDraft(stored, 'test'), /Quantity/);
  for (const invalid of [-1, 1.5, NaN, Infinity, '2']) {
    stored.document.nodes[0].data.pricing.quantity = invalid;
    assert.throws(() => validateDraft(stored, 'test'), /[Qq]uantity/);
  }
});

test('legacy price normalization is immutable and zero remains a known estimate', () => {
  const node = draft().document.nodes[0];
  node.data.pricing = { estimatedCost: 0 };
  const restored = validateRestoredNodes([node]);
  assert.deepEqual(node.data.pricing, { estimatedCost: 0 });
  assert.deepEqual(restored[0].data.pricing, { estimatedCost: 0, quantity: 1, region: 'Unknown' });
  node.data.pricing = { estimatedCost: undefined };
  assert.throws(() => validateRestoredNodes([node]), /estimatedCost/);
});

test('restores retain bounded sizes, tags, edge selection and exact valid handles', () => {
  const node = draft().document.nodes[0];
  node.style = { width: 320, height: 200 };
  node.data.tags = ['Production', 'Owner: Team A'];
  const nodes = validateRestoredNodes([node]);
  assert.deepEqual(nodes[0].data.tags, node.data.tags);
  for (const invalid of [0, -1, Infinity, 100_001, 'Infinity', '100%']) {
    assert.throws(() => validateRestoredNodes([{ ...node, style: { width: invalid } }]), /width/);
  }
  for (const invalid of [0, -1, Infinity, NaN, 100_001]) {
    assert.throws(() => validateRestoredNodes([{ ...node, height: invalid }]), /height/);
  }
  assert.throws(() => validateRestoredNodes([{ ...node, data: { tags: ['x'.repeat(41)] } }]), /tags/);
  const edge = { id: 'edge', source: 'one', target: 'one', sourceHandle: 'top-source', targetHandle: 'right-target', selected: true };
  assert.deepEqual(validateRestoredEdges([edge], new Set(['one'])), [edge]);
  assert.throws(() => validateRestoredEdges([{ ...edge, sourceHandle: 3 }], new Set(['one'])), /sourceHandle/);
});

test('extended document settings, lineage, IaC and scenarios survive drafts', () => {
  const stored = draft();
  stored.document.lineageId = 'local:stable-document';
  stored.document.settings = {
    ...stored.document.settings, pricingRegion: 'japaneast', animateConnections: false,
    showCostBadges: false, layoutPreset: 'swimlanes', layoutSpacing: 'compact', layoutEngine: 'elk',
  };
  stored.document.iacBaseline = null;
  stored.document.pricingScenarios = [];
  assert.deepEqual(validateDraft(stored, 'test'), stored);
  assert.throws(() => validateDraft({ ...stored, document: { ...stored.document, reviewHistory: [{ id: 'not-a-review' }] } }, 'test'), /review history/);
});
