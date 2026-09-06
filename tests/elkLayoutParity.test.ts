import assert from 'node:assert/strict';
import test from 'node:test';
import { applyLayoutPreset } from '../src/utils/layoutPresets';
import { applyAutomaticEdgeLabelOffsets } from '../src/utils/edgeLabelLayout';
import { disposeElkLayout } from '../src/utils/elkLayoutRuntime';
import { elkLayoutFixtures } from './fixtures/elkLayoutParity';

for (const fixture of elkLayoutFixtures()) {
  test(`Node ELK retains deterministic, non-destructive output: ${fixture.name}`, async () => {
    const input = structuredClone(fixture);
    const first = await applyLayoutPreset(fixture.nodes, fixture.edges, fixture.options);
    const second = await applyLayoutPreset(fixture.nodes, fixture.edges, fixture.options);
    assert.deepEqual(second, first);
    assert.deepEqual(fixture, input);
    assert.deepEqual(first.nodes.map(node => node.id), fixture.nodes.map(node => node.id));
    assert.deepEqual(first.edges.map(edge => edge.id), fixture.edges.map(edge => edge.id));
    for (const node of first.nodes) {
      const original = input.nodes.find(candidate => candidate.id === node.id);
      assert.ok(original);
      assert.deepEqual(node.data, original.data);
      assert.equal(node.parentNode, original.parentNode);
      assert.ok(Number.isFinite(node.position.x) && Number.isFinite(node.position.y));
    }
    for (const edge of first.edges) {
      const original = input.edges.find(candidate => candidate.id === edge.id);
      assert.ok(original);
      assert.equal(edge.source, original.source);
      assert.equal(edge.target, original.target);
      assert.equal(edge.label, original.label);
      assert.equal(edge.data?.direction, original.data?.direction);
    }
    const firstLabels = applyAutomaticEdgeLabelOffsets(first.nodes, first.edges);
    const secondLabels = applyAutomaticEdgeLabelOffsets(second.nodes, second.edges);
    assert.deepEqual(secondLabels, firstLabels);
  });
}

test('Node ELK can be disposed and initialized again without a native Worker', async () => {
  const fixture = elkLayoutFixtures()[0];
  const before = await applyLayoutPreset(fixture.nodes, fixture.edges, fixture.options);
  disposeElkLayout();
  const after = await applyLayoutPreset(fixture.nodes, fixture.edges, fixture.options);
  assert.deepEqual(after, before);
  disposeElkLayout();
});
