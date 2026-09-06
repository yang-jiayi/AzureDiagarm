import test from 'node:test';
import assert from 'node:assert/strict';
import type { Edge, Node } from 'reactflow';
import { applyDiagramChanges, buildDiagramChanges, DiagramSelectionError, isCompleteDiagramChangeSet, type DiagramGraph } from '../src/services/diagramChanges';
import { MAX_PRICING_AMOUNT, MAX_PRICING_QUANTITY } from '../src/services/pricingConfiguration';

const node = (id: string, label = id, extra: Partial<Node> = {}): Node => ({
  id, type: 'azureNode', position: { x: 0, y: 0 }, data: { label }, ...extra,
});
const edge = (id: string, source: string, target: string, label = ''): Edge => ({ id, source, target, label });
const graph = (nodes: Node[] = [], edges: Edge[] = []): DiagramGraph => ({ nodes, edges });
const all = (set: ReturnType<typeof buildDiagramChanges>) => new Set(set.changes.map(change => change.id));

test('all, none and mixed selections actually produce only selected changes', () => {
  const before = graph([node('a'), node('b'), node('delete')], [edge('e', 'a', 'b', 'old')]);
  const proposed = graph([node('a', 'Renamed'), node('b'), node('new')], [edge('e', 'a', 'b', 'new')]);
  const set = buildDiagramChanges(before, proposed);
  assert.equal(set.changes.length, 4);
  assert.deepEqual(applyDiagramChanges(set, new Set()), before);
  assert.deepEqual(applyDiagramChanges(set, all(set)), proposed);
  const partial = applyDiagramChanges(set, new Set(['node:change:a', 'node:add:new']));
  assert.deepEqual(partial.nodes.map(node => node.id), ['a', 'b', 'delete', 'new']);
  assert.equal(partial.nodes[0].data.label, 'Renamed');
  assert.equal(partial.edges[0].label, 'old');
});

test('every subset of independent changes is faithfully applied', () => {
  const set = buildDiagramChanges(
    graph([node('a'), node('b')]), graph([node('a', 'A2'), node('c')]),
  );
  for (let mask = 0; mask < 1 << set.changes.length; mask++) {
    const selected = new Set(set.changes.filter((_, index) => mask & (1 << index)).map(change => change.id));
    const result = applyDiagramChanges(set, selected);
    assert.equal(result.nodes[0].data.label, selected.has('node:change:a') ? 'A2' : 'a');
    assert.equal(result.nodes.some(node => node.id === 'b'), !selected.has('node:delete:b'));
    assert.equal(result.nodes.some(node => node.id === 'c'), selected.has('node:add:c'));
  }
});

test('new edges require their selected new endpoints', () => {
  const set = buildDiagramChanges(graph([node('a')]), graph([node('a'), node('b')], [edge('e', 'a', 'b')]));
  assert.throws(() => applyDiagramChanges(set, new Set(['edge:add:e'])), /requires node "b"/);
  assert.equal(applyDiagramChanges(set, new Set(['node:add:b'])).edges.length, 0);
  assert.equal(applyDiagramChanges(set, all(set)).edges.length, 1);
});

test('deleting a node requires deleting or reconnecting every retained edge', () => {
  const before = graph([node('a'), node('b'), node('c')], [edge('e', 'a', 'b')]);
  const reconnect = buildDiagramChanges(before, graph([node('a'), node('c')], [edge('e', 'a', 'c')]));
  assert.throws(() => applyDiagramChanges(reconnect, new Set(['node:delete:b'])), DiagramSelectionError);
  assert.deepEqual(applyDiagramChanges(reconnect, all(reconnect)).edges[0].target, 'c');
  const deletion = buildDiagramChanges(before, graph([node('a'), node('c')]));
  assert.throws(() => applyDiagramChanges(deletion, new Set(['node:delete:b'])), DiagramSelectionError);
  assert.equal(applyDiagramChanges(deletion, all(deletion)).edges.length, 0);
});

test('new parents are dependencies and parents are returned before children', () => {
  const proposed = graph([node('child', 'child', { parentNode: 'group', extent: 'parent' }), node('group', 'Group', { type: 'groupNode' })]);
  const set = buildDiagramChanges(graph(), proposed);
  assert.throws(() => applyDiagramChanges(set, new Set(['node:add:child'])), /requires parent/);
  assert.deepEqual(applyDiagramChanges(set, all(set)).nodes.map(node => node.id), ['group', 'child']);
});

test('parent deletion requires child deletion, reparenting or explicit ungrouping', () => {
  const before = graph([node('g', 'Group', { type: 'groupNode' }), node('c', 'Child', { parentNode: 'g', extent: 'parent' })]);
  for (const proposed of [graph(), graph([node('c', 'Child')]), graph([
    node('next', 'Next', { type: 'groupNode' }), node('c', 'Child', { parentNode: 'next', extent: 'parent' }),
  ])]) {
    const set = buildDiagramChanges(before, proposed);
    assert.throws(() => applyDiagramChanges(set, new Set(['node:delete:g'])), /requires parent/);
    const result = applyDiagramChanges(set, all(set));
    assert.equal(result.nodes.some(node => node.id === 'g'), false);
    if (proposed.nodes.length === 1) {
      assert.equal(result.nodes[0].parentNode, undefined);
      assert.equal(result.nodes[0].extent, undefined);
    }
  }
});

test('valid proposals can still yield invalid cyclic subsets, which are rejected', () => {
  const before = graph([node('a'), node('b', 'b', { parentNode: 'a' })]);
  const proposed = graph([node('a', 'a', { parentNode: 'b' }), node('b')]);
  const set = buildDiagramChanges(before, proposed);
  assert.throws(() => applyDiagramChanges(set, new Set(['node:change:a'])), /cycle/);
  assert.deepEqual(applyDiagramChanges(set, all(set)).nodes.map(node => node.id), ['b', 'a']);
});

test('duplicate IDs, missing endpoints/parents and parent cycles reject proposals', () => {
  for (const invalid of [
    graph([node('a'), node('a')]),
    graph([node('a')], [edge('e', 'a', 'a'), edge('e', 'a', 'a')]),
    graph([node('a')], [edge('e', 'a', 'missing')]),
    graph([node('a', 'a', { parentNode: 'missing' })]),
    graph([node('a', 'a', { parentNode: 'b' }), node('b', 'b', { parentNode: 'a' })]),
  ]) assert.throws(() => buildDiagramChanges(graph(), invalid), DiagramSelectionError);
  assert.throws(() => applyDiagramChanges(buildDiagramChanges(graph(), graph()), new Set(['typo'])), /Unknown change/);
});

test('ordinary connection cycles and parallel edges are legal', () => {
  const proposed = graph([node('a'), node('b')], [edge('e1', 'a', 'b'), edge('e2', 'b', 'a'), edge('e3', 'a', 'b')]);
  const set = buildDiagramChanges(graph(), proposed);
  assert.equal(applyDiagramChanges(set, all(set)).edges.length, 3);
});

test('stable IDs win even when labels are swapped', () => {
  const set = buildDiagramChanges(graph([node('a', 'First'), node('b', 'Second')]), graph([node('a', 'Second'), node('b', 'First')]));
  assert.deepEqual(set.changes.map(change => [change.kind, change.entityId]), [['change', 'a'], ['change', 'b']]);
});

test('canonical service aliases retain manual instances, labels, pricing and connections', () => {
  const pricing = { estimatedCost: 120, quantity: 2, region: 'eastus2', isCustom: true };
  const before = graph([node('manual-web', 'Customer Portal', {
    position: { x: 320, y: 240 },
    data: { label: 'Customer Portal', serviceName: 'App Services', pricing },
  })]);
  for (const label of ['App Service', 'Azure App Service']) {
    const set = buildDiagramChanges(before, graph([
      node('generated-web', label, { data: { label, serviceName: 'App Service' } }),
      node('generated-db', 'SQL Database', { data: { label: 'SQL Database', serviceName: 'SQL Database' } }),
    ], [edge('generated-edge', 'generated-web', 'generated-db')]));
    const result = applyDiagramChanges(set, all(set));
    assert.deepEqual(result.nodes.map(item => item.id), ['manual-web', 'generated-db']);
    assert.equal(result.nodes[0].data.label, 'Customer Portal');
    assert.deepEqual(result.nodes[0].position, before.nodes[0].position);
    assert.deepEqual(result.nodes[0].data.pricing, pricing);
    assert.equal(result.edges[0].source, 'manual-web');
    assert.equal(before.nodes[0].data.serviceName, 'App Services');
  }
});

test('canonical service aliases cannot make multiple instances look unique', () => {
  const before = graph([
    node('first', 'First portal', { data: { label: 'First portal', serviceName: 'App Services' } }),
    node('second', 'Second portal', { data: { label: 'Second portal', serviceName: 'App Service' } }),
  ]);
  const set = buildDiagramChanges(before, graph([
    node('generated', 'App Service', { data: { label: 'App Service', serviceName: 'App Service' } }),
  ]));
  assert.equal(set.proposed.nodes[0].id, 'generated');
  assert.deepEqual(set.changes.filter(change => change.kind === 'delete').map(change => change.entityId), ['first', 'second']);
});

test('service aliases do not reinterpret group names', () => {
  const set = buildDiagramChanges(
    graph([node('old-group', 'App Services', { type: 'groupNode' })]),
    graph([node('new-group', 'App Service', { type: 'groupNode' })]),
  );
  assert.equal(set.proposed.nodes[0].id, 'new-group');
});

test('imported proposals preserve their own IDs, connections and metadata', () => {
  const before = graph([
    node('old-web', 'Customer portal', {
      position: { x: 400, y: 300 },
      data: { label: 'Customer portal', serviceName: 'App Service', customNote: 'Previous document' },
    }),
    node('shared-id', 'Old database', { data: { label: 'Old database', customNote: 'Do not inherit' } }),
  ]);
  const imported = graph([
    node('file-web', 'App Service', { data: { label: 'App Service', serviceName: 'App Service' } }),
    node('shared-id', 'Imported database'),
  ], [edge('file-edge', 'file-web', 'shared-id')]);
  const set = buildDiagramChanges(before, imported, { reconcile: false });
  assert.deepEqual(set.proposed, imported);
  const applied = applyDiagramChanges(set, all(set));
  assert.deepEqual(new Map(applied.nodes.map(value => [value.id, value])), new Map(imported.nodes.map(value => [value.id, value])));
  assert.deepEqual(applied.edges, imported.edges);
  assert.equal(isCompleteDiagramChangeSet(set, applied), true);
  assert.equal(set.changes.some(change => change.id === 'node:delete:old-web'), true);
  assert.equal(set.changes.some(change => change.id === 'node:add:file-web'), true);
  set.proposed.nodes[0].data.label = 'Edited copy';
  assert.equal(imported.nodes[0].data.label, 'App Service');
  assert.equal(before.nodes[0].data.customNote, 'Previous document');
});

test('regenerated unique service and group IDs are matched and references remapped', () => {
  const before = graph([
    node('g', 'App', { type: 'groupNode' }),
    node('web', 'Customer API', { data: { label: 'Customer API', serviceName: 'App Service' }, parentNode: 'g' }),
    node('db', 'SQL'),
  ], [edge('route', 'web', 'db', 'Query')]);
  const proposed = graph([
    node('group-new', 'App', { type: 'groupNode' }),
    node('web-new', 'App Service', { data: { label: 'App Service', serviceName: 'App Service' }, parentNode: 'group-new' }),
    node('db-new', 'SQL'),
  ], [edge('edge-0', 'web-new', 'db-new', 'Query')]);
  const set = buildDiagramChanges(before, proposed);
  assert.equal(set.changes.length, 0);
  assert.deepEqual(set.proposed.nodes.map(node => node.id), ['g', 'web', 'db']);
  assert.equal(set.proposed.nodes[1].data.label, 'Customer API');
  assert.equal(set.proposed.nodes[1].parentNode, 'g');
  assert.equal(set.proposed.edges[0].source, 'web');
  assert.equal(set.proposed.edges[0].id, 'route');
});

test('ambiguous service identities are never guessed or collapsed', () => {
  const before = graph([node('a', 'SQL'), node('b', 'SQL')]);
  const proposed = graph([node('c', 'SQL'), node('d', 'SQL')]);
  const set = buildDiagramChanges(before, proposed);
  assert.equal(set.changes.filter(change => change.kind === 'delete').length, 2);
  assert.equal(set.changes.filter(change => change.kind === 'add').length, 2);
  assert.deepEqual(applyDiagramChanges(set, all(set)), proposed);
});

test('same-service instances may be distinguished by unique instance labels', () => {
  const instance = (id: string, label: string) => node(id, label, { data: { label, serviceName: 'SQL' } });
  const set = buildDiagramChanges(graph([instance('a', 'East'), instance('b', 'West')]),
    graph([instance('new-b', 'West'), instance('new-a', 'East')]));
  assert.equal(set.changes.length, 0);
  assert.deepEqual(set.proposed.nodes.map(node => node.id), ['b', 'a']);
});

test('generated index IDs do not turn reordered connections into edits', () => {
  const nodes = [node('a'), node('b'), node('c')];
  const before = graph(nodes, [edge('edge-0', 'a', 'b', 'Read'), edge('edge-1', 'b', 'c', 'Write')]);
  before.edges[0].data = { labelOffsetX: 42, onLabelChange: () => {} };
  before.edges[0].style = { stroke: 'red' };
  const proposed = graph(nodes, [edge('edge-0', 'b', 'c', 'Write'), edge('edge-1', 'a', 'b', 'Read')]);
  const set = buildDiagramChanges(before, proposed);
  assert.equal(set.changes.length, 0);
  assert.equal(set.proposed.edges[1].data.labelOffsetX, 42);
  assert.deepEqual(set.proposed.edges[1].style, { stroke: 'red' });
});

test('parallel regenerated edges remain distinct through reordering and label edits', () => {
  const nodes = [node('a'), node('b')];
  const before = graph(nodes, [edge('edge-0', 'a', 'b', 'Read'), edge('edge-1', 'a', 'b', 'Write')]);
  const reordered = buildDiagramChanges(before, graph(nodes, [
    edge('edge-0', 'a', 'b', 'Write'), edge('edge-1', 'a', 'b', 'Read'),
  ]));
  assert.equal(reordered.changes.length, 0);
  const changed = buildDiagramChanges(before, graph(nodes, [
    edge('edge-0', 'a', 'b', 'New action'), edge('edge-1', 'a', 'b', 'Write'),
  ]));
  const output = applyDiagramChanges(changed, all(changed));
  assert.equal(output.edges.length, 2);
  assert.equal(new Set(output.edges.map(edge => edge.id)).size, 2);
  assert.deepEqual(output.edges.map(edge => edge.label).sort(), ['New action', 'Write']);
});

test('generated routing defaults preserve hand-routed edges and per-node presentation', () => {
  const nodes = [node('a', 'API', { data: { label: 'API', serviceName: 'App Service', stylePreset: 'compact' } }), node('b')];
  const old = { ...edge('edge-0', 'a', 'b', 'Query'), sourceHandle: 'bottom', targetHandle: 'top',
    style: { stroke: 'red' }, labelStyle: { fontSize: 20 }, data: { labelOffsetX: 70, pathStyle: 'bezier' } };
  const proposed = { ...edge('edge-0', 'a', 'b', 'Query'), sourceHandle: 'right', targetHandle: 'left',
    style: {}, labelStyle: { fontSize: 14 }, data: { labelOffsetX: 0, pathStyle: 'orthogonal' } };
  const set = buildDiagramChanges(graph(nodes, [old]), graph([
    { ...nodes[0], data: { ...nodes[0].data, stylePreset: 'detailed' } }, nodes[1],
  ], [proposed]));
  assert.equal(set.changes.length, 0);
  assert.deepEqual(set.proposed.edges[0], old);
});

test('identical parallel edges preserve multiplicity and never reuse one edge twice', () => {
  const nodes = [node('a'), node('b')];
  const set = buildDiagramChanges(graph(nodes, [edge('one', 'a', 'b'), edge('two', 'a', 'b')]),
    graph(nodes, [edge('edge-0', 'a', 'b'), edge('edge-1', 'a', 'b')]));
  assert.equal(set.changes.length, 0);
  assert.equal(new Set(set.proposed.edges.map(edge => edge.id)).size, 2);
});

test('new generated index IDs cannot collide with retained or deleted old edges', () => {
  const nodes = [node('a'), node('b'), node('c')];
  const set = buildDiagramChanges(graph(nodes, [edge('edge-0', 'a', 'b')]),
    graph(nodes, [edge('edge-0', 'a', 'c'), edge('edge-1', 'a', 'b')]));
  const output = applyDiagramChanges(set, all(set));
  assert.equal(output.edges.length, 2);
  assert.equal(new Set(output.edges.map(edge => edge.id)).size, 2);
  assert.equal(output.edges.find(edge => edge.target === 'b')!.id, 'edge-0');
});

test('callbacks, key order and React Flow measurements are not content edits', () => {
  const callback = () => {};
  const before = graph([node('a', 'A', { data: { label: 'A', onEdit: callback }, width: 100 })]);
  const set = buildDiagramChanges(before, graph([node('a', 'A', {
    data: { onEdit: () => {}, label: 'A' }, width: 200, selected: true, dragging: true,
  })]));
  assert.equal(set.changes.length, 0);
  assert.equal(set.proposed.nodes[0].data.onEdit, callback);
  assert.deepEqual(applyDiagramChanges(set, all(set)), before);
});

test('explicit content/style edits retain existing unspecified metadata and callbacks', () => {
  const callback = () => {};
  const before = graph([node('a', 'A', { data: { label: 'A', onEdit: callback, description: 'Keep' }, style: { color: 'red', background: 'white' } })]);
  const set = buildDiagramChanges(before, graph([node('a', 'B', { style: { color: 'blue' } })]));
  const output = applyDiagramChanges(set, all(set));
  assert.equal(output.nodes[0].data.label, 'B');
  assert.equal(output.nodes[0].data.description, 'Keep');
  assert.equal(output.nodes[0].data.onEdit, callback);
  assert.deepEqual(output.nodes[0].style, { color: 'blue', background: 'white' });
});

test('proposals and outputs are isolated snapshots without losing callbacks', () => {
  const before = graph([node('a')]);
  const proposed = graph([node('a', 'New')]);
  const set = buildDiagramChanges(before, proposed);
  before.nodes[0].data.label = 'External mutation';
  proposed.nodes[0].data.label = 'Another mutation';
  const output = applyDiagramChanges(set, all(set));
  output.nodes[0].data.label = 'Output mutation';
  assert.equal(set.before.nodes[0].data.label, 'a');
  assert.equal(set.proposed.nodes[0].data.label, 'New');
  assert.equal(applyDiagramChanges(set, all(set)).nodes[0].data.label, 'New');
});

test('cost deltas include quantity, zero, adds/deletes and unknown estimates', () => {
  const priced = (id: string, estimatedCost: number, quantity = 1) => node(id, id, { data: { label: id, pricing: { estimatedCost, quantity } } });
  const set = buildDiagramChanges(graph([priced('a', 10, 2), priced('b', 0), node('unknown')]),
    graph([priced('a', 20, 3), priced('c', 7), node('unknown', 'Changed')]));
  assert.equal(set.changes.find(change => change.entityId === 'a')?.costDelta, 40);
  assert.equal(set.changes.find(change => change.entityId === 'b')?.costDelta, 0);
  assert.equal(set.changes.find(change => change.entityId === 'c')?.costDelta, 7);
  assert.equal(set.changes.find(change => change.entityId === 'unknown')?.costDelta, undefined);
});

test('AI previews and accepted changes preserve 100,000-unit pricing multipliers', () => {
  const proposed = node('legacy', 'Legacy', {
    data: { label: 'Legacy', pricing: { estimatedCost: 2, quantity: 100_000 } },
  });
  const set = buildDiagramChanges(graph(), graph([proposed]));
  assert.equal(set.changes[0].costDelta, 200_000);
  const applied = applyDiagramChanges(set, all(set));
  assert.equal(applied.nodes.length, 1);
  assert.equal(applied.nodes[0].data.pricing.quantity, 100_000);
});

test('explicitly unpriced services never use stale legacy costs or invalid quantities', () => {
  for (const pricing of [
    { estimatedCost: null, quantity: 1 },
    { estimatedCost: 5, quantity: MAX_PRICING_QUANTITY + 1 },
    { estimatedCost: 5, quantity: 1.5 },
    { estimatedCost: MAX_PRICING_AMOUNT + 1, quantity: 1 },
  ]) {
    const proposed = node('unpriced', 'Unpriced', { data: { label: 'Unpriced', monthlyCost: 100, pricing } });
    const set = buildDiagramChanges(graph(), graph([proposed]));
    assert.equal(set.changes[0].costDelta, undefined);
  }
});

test('reusing an ID for another service does not inherit its old custom pricing', () => {
  const before = node('service', 'API', {
    data: { label: 'API', serviceName: 'App Service', pricing: { estimatedCost: 100, quantity: 2, isCustom: true } },
  });
  const proposed = node('service', 'SQL', { data: { label: 'SQL', serviceName: 'SQL Database' } });
  const set = buildDiagramChanges(graph([before]), graph([proposed]));
  assert.equal(set.proposed.nodes[0].data.pricing, undefined);
  assert.equal(applyDiagramChanges(set, all(set)).nodes[0].data.pricing, undefined);
  assert.equal(before.data.pricing.estimatedCost, 100);
});

test('normalized parents also update compatibility metadata without changing the source proposal', () => {
  const before = graph([node('group', 'Group', { type: 'groupNode' })]);
  const proposed = graph([
    node('generated-group', 'Group', { type: 'groupNode' }),
    node('child', 'Child', { parentNode: 'generated-group', data: { label: 'Child', parentNode: 'generated-group' } }),
  ]);
  const set = buildDiagramChanges(before, proposed);
  assert.equal(set.proposed.nodes[1].parentNode, 'group');
  assert.equal(set.proposed.nodes[1].data.parentNode, 'group');
  assert.equal(proposed.nodes[1].data.parentNode, 'generated-group');
});

test('completeness recognizes full normalized proposals despite ordering and regenerated IDs', () => {
  const before = graph([node('a', 'App'), node('b', 'SQL')], [edge('edge-0', 'a', 'b', 'Read')]);
  const proposed = graph([node('new-b', 'SQL'), node('new-a', 'App'), node('cache', 'Cache')], [
    edge('edge-0', 'new-a', 'cache', 'Cache read'),
    edge('edge-1', 'new-a', 'new-b', 'Read'),
  ]);
  const set = buildDiagramChanges(before, proposed);
  const full = applyDiagramChanges(set, all(set));
  assert.equal(isCompleteDiagramChangeSet(set, full), true);
  assert.equal(isCompleteDiagramChangeSet(set, { nodes: [...full.nodes].reverse(), edges: [...full.edges].reverse() }), true);
  assert.equal(isCompleteDiagramChangeSet(set, set.proposed), true);
  assert.equal(isCompleteDiagramChangeSet(set, proposed), false, 'raw generated IDs are not canonical IDs');
  assert.equal(isCompleteDiagramChangeSet(set, before), false);
});

test('completeness detects rejected node/edge edits even when counts are identical', () => {
  const before = graph([node('a', 'A'), node('b', 'B')], [edge('route', 'a', 'b', 'Read')]);
  const set = buildDiagramChanges(before, graph([node('a', 'Renamed'), node('b', 'B')], [edge('route', 'a', 'b', 'Write')]));
  assert.equal(isCompleteDiagramChangeSet(set, applyDiagramChanges(set, all(set))), true);
  for (const selected of [new Set<string>(), new Set(['node:change:a']), new Set(['edge:change:route'])]) {
    const partial = applyDiagramChanges(set, selected);
    assert.equal(partial.nodes.length, set.proposed.nodes.length);
    assert.equal(partial.edges.length, set.proposed.edges.length);
    assert.equal(isCompleteDiagramChangeSet(set, partial), false);
  }
});

test('completeness ignores callbacks and measurements but detects content, routes and styles', () => {
  const before = graph([node('a'), node('b')], [edge('route', 'a', 'b', 'Read')]);
  const set = buildDiagramChanges(graph(), before);
  const full = applyDiagramChanges(set, all(set));
  full.nodes[0] = { ...full.nodes[0], selected: true, width: 300, height: 200,
    data: { ...full.nodes[0].data, onEdit: () => {} } };
  full.edges[0] = { ...full.edges[0], selected: true, data: { onLabelChange: () => {} } };
  assert.equal(isCompleteDiagramChangeSet(set, full), true);
  assert.equal(buildDiagramChanges(before, full).changes.length, 0);
  const callbacksSet = buildDiagramChanges(graph(), full);
  const withNewCallbacks = applyDiagramChanges(callbacksSet, all(callbacksSet));
  withNewCallbacks.nodes[0].data.onEdit = () => {};
  withNewCallbacks.nodes[0].width = 500;
  withNewCallbacks.edges[0].data.onLabelChange = () => {};
  assert.equal(isCompleteDiagramChangeSet(callbacksSet, withNewCallbacks), true);
  for (const mutate of [
    (value: DiagramGraph) => { value.nodes[0].data.label = 'Not accepted'; },
    (value: DiagramGraph) => { value.nodes[0].style = { color: 'red' }; },
    (value: DiagramGraph) => { value.edges[0].sourceHandle = 'bottom'; },
    (value: DiagramGraph) => { value.edges[0].data.labelOffsetX = 50; },
  ]) {
    const changed = applyDiagramChanges(callbacksSet, all(callbacksSet));
    mutate(changed);
    assert.equal(isCompleteDiagramChangeSet(callbacksSet, changed), false);
  }
});

test('completeness covers empty baselines, no-op proposals and complete deletion', () => {
  const initial = buildDiagramChanges(graph(), graph([node('a')]));
  assert.equal(isCompleteDiagramChangeSet(initial, graph()), false);
  assert.equal(isCompleteDiagramChangeSet(initial, applyDiagramChanges(initial, all(initial))), true);
  const empty = buildDiagramChanges(graph(), graph());
  assert.equal(isCompleteDiagramChangeSet(empty, graph()), true);
  const noOp = buildDiagramChanges(graph([node('a')]), graph([node('a')]));
  assert.equal(isCompleteDiagramChangeSet(noOp, graph([node('a')])), true);
  const deletion = buildDiagramChanges(graph([node('a')]), graph());
  assert.equal(isCompleteDiagramChangeSet(deletion, graph()), true);
  assert.equal(isCompleteDiagramChangeSet(deletion, deletion.before), false);
});

test('completeness rejects malformed graphs and same-sized ID substitution', () => {
  const set = buildDiagramChanges(graph(), graph([node('a'), node('b')], [edge('route', 'a', 'b')]));
  for (const invalid of [
    graph([node('a'), node('a')], set.proposed.edges),
    graph(set.proposed.nodes, [edge('route', 'a', 'missing')]),
    graph([node('a', 'a', { parentNode: 'b' }), node('b', 'b', { parentNode: 'a' })], set.proposed.edges),
    graph([node('a'), node('other', 'b')], [edge('route', 'a', 'other')]),
  ]) assert.equal(isCompleteDiagramChangeSet(set, invalid), false);
});
