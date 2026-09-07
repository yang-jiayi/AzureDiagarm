import assert from 'node:assert/strict';
import { createHash } from 'node:crypto';
import test from 'node:test';
import { applyLayoutPreset } from '../src/utils/layoutPresets';
import { applyAutomaticEdgeLabelOffsets } from '../src/utils/edgeLabelLayout';
import { disposeElkLayout } from '../src/utils/elkLayoutRuntime';
import { layoutArchitecture } from '../src/utils/elkLayoutEngine';
import { elkLayoutFixtures } from './fixtures/elkLayoutParity';

// Exact serialized outputs captured before replacing the adapter's array/tree scans.
const preIndexDigests: Record<string, string> = {
  '20-ungrouped': '31bf198bff01f133bf71228062349bf28682e14572ba5d0b62b6e2c44ac264e0',
  '80-ungrouped': '01296c0b4da62c22e1ff133ca67f06b1c7bb91831aae84898dea0d369795206c',
  '250-ungrouped': 'd3807b73ff5bea6c570bbe0e34a2fd76f363f5792bfb9b082711ab1ee83ddb4f',
  '20-grouped': '01b3cd343ee8fc09fb9af2a8d9ff5d4262562a742e8e875c0b07aa42e20675ec',
  '80-grouped': 'a6bdfd17695f58334d3275e73dbc8424ed627ad3200d8e5d20ca7e11ebaa237c',
  '250-grouped': '787679c490028a4bed0fd2815c03f7e1072e95151b69e0bf3b13877709d6a3f8',
  'nested-hierarchy': '563a7ffcc5737b9bc563cde7470bf1da4a74f718778930388c0636af49bb2ed5',
  'cycles-parallel-and-duplicate-services': 'c01f18030ce6901576d39455ee57ba87deb57ad8bda73fb37f4b49af45a2e64b',
  'manual-dimensions-handles-and-labels': '35ac70f57c66172ec024ba9fbcefef83825dd27dc134f783e1c578009350de7f',
  'disconnected': '4e0d2e70ba25b49d6ae7f5dcb460305041ab44f5e354622d3918658c6f82ae6c',
  'group-order-empty-and-cross-boundary': '0b09d358cdd970cad6ea3793220af540298cb60a20d89ee14da192de0188cfd1',
};

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
    assert.equal(
      createHash('sha256').update(JSON.stringify({ ...first, edges: firstLabels })).digest('hex'),
      preIndexDigests[fixture.name],
      `Layout coordinates, hierarchy, order, handles and labels changed for ${fixture.name}`,
    );
  });
}

test('ELK indexes group membership without scanning services or filtering all connections per group', async () => {
  const services = [
    { id: 'a', name: 'A', groupId: 'zone-z' },
    { id: 'b', name: 'B', groupId: 'zone-z' },
    { id: 'c', name: 'C', groupId: 'zone-a' },
    { id: 'd', name: 'D', groupId: 'zone-a' },
    { id: 'free', name: 'Free' },
  ];
  const connections = [
    { from: 'b', to: 'c' },
    { from: 'a', to: 'b' },
    { from: 'a', to: 'b' },
    { from: 'c', to: 'd' },
    { from: 'd', to: 'free' },
    { from: 'free', to: 'a' },
  ];
  const groups = [{ id: 'zone-z', label: 'Z' }, { id: 'zone-a', label: 'A' }];
  const expected = await layoutArchitecture(services, connections, groups);
  let connectionFilters = 0;
  const indexedServices = new Proxy(services, {
    get(target, property, receiver) {
      if (property === 'find') throw new Error('Services must not be searched per connection');
      return Reflect.get(target, property, receiver);
    },
  });
  const indexedConnections = new Proxy(connections, {
    get(target, property, receiver) {
      if (property === 'filter') connectionFilters++;
      return Reflect.get(target, property, receiver);
    },
  });
  assert.deepEqual(await layoutArchitecture(indexedServices, indexedConnections, groups), expected);
  assert.equal(connectionFilters, 1, 'Only the top-level edge pass may filter all connections');
});

test('Node ELK can be disposed and initialized again without a native Worker', async () => {
  const fixture = elkLayoutFixtures()[0];
  const before = await applyLayoutPreset(fixture.nodes, fixture.edges, fixture.options);
  disposeElkLayout();
  const after = await applyLayoutPreset(fixture.nodes, fixture.edges, fixture.options);
  assert.deepEqual(after, before);
  disposeElkLayout();
});
