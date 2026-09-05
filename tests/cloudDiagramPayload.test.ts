import assert from 'node:assert/strict';
import test from 'node:test';
import { toCloudDiagramPayload } from '../src/services/cloudDiagramPayload';
import type { EditorDocument } from '../src/services/editorHistory';
import { canonicalStringify } from '../src/utils/canonicalJson';

function document(): EditorDocument {
  return {
    nodes: [{
      id: 'app',
      type: 'azureNode',
      position: { x: 20, y: 40 },
      data: { label: 'Application' },
    }],
    edges: [],
    titleBlockData: {
      architectureName: 'Source diagram', author: 'Author', version: '1.0', date: '2026-09-05',
    },
    workflow: [],
    architecturePrompt: '',
    originalPrompt: '',
    settings: { pricingMode: 'payg', stylePreset: 'detailed', edgeStyle: 'orthogonal' },
  };
}

test('cloud hydration and live editor snapshots have the same canonical baseline', () => {
  const loaded = document();
  const live: EditorDocument = {
    ...loaded,
    nodes: [{ ...loaded.nodes[0], selected: true, dragging: true, width: 200, height: 100 }],
    settings: { edgeStyle: 'orthogonal', stylePreset: 'detailed', pricingMode: 'payg' },
    reviewHistory: [],
    validationSourceFingerprint: null,
    iacBaseline: null,
    lineageId: 'cloud:source',
    viewport: { x: 5, y: 10, zoom: 1.2 },
  };
  assert.equal(canonicalStringify(toCloudDiagramPayload(live)), canonicalStringify(toCloudDiagramPayload(loaded)));
  assert.equal(live.nodes[0].selected, true);
  assert.equal(live.nodes[0].width, 200);
});

test('cloud baselines still distinguish real edits and array ordering', () => {
  const before = document();
  const after = structuredClone(before);
  after.nodes[0].data.label = 'Edited application';
  assert.notEqual(canonicalStringify(toCloudDiagramPayload(before)), canonicalStringify(toCloudDiagramPayload(after)));
  assert.notEqual(canonicalStringify({ steps: [1, 2] }), canonicalStringify({ steps: [2, 1] }));
});

test('cloud payload retains recorded validation, provenance and original prompt', () => {
  const source = document();
  source.validationScore = 0;
  source.validationSourceFingerprint = 'original-graph';
  source.architecturePrompt = 'Current prompt';
  source.originalPrompt = 'Original prompt';
  source.reviewHistory = [{ summary: 'Recorded review' }];
  source.edges = [{ id: 'edge', source: 'app', target: 'app', selected: true, label: 'Route' }];
  const payload = toCloudDiagramPayload(source);
  assert.equal(payload.validationScore, 0);
  assert.equal(payload.validationSourceFingerprint, 'original-graph');
  assert.equal(payload.originalPrompt, 'Original prompt');
  assert.deepEqual(payload.reviewHistory, source.reviewHistory);
  assert.equal(payload.edges[0].label, 'Route');
  assert.equal('selected' in payload.edges[0], false);
  assert.equal(source.edges[0].selected, true);
});
