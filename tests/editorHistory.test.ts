import test from 'node:test';
import assert from 'node:assert/strict';
import { EditorHistory, editorFingerprint, deleteSelectedElements, duplicateSelectedElements, type EditorDocument } from '../src/services/editorHistory';

function document(label = 'Web'): EditorDocument {
  return {
    nodes: [{ id: 'web', type: 'azureNode', position: { x: 20, y: 40 }, data: { label } }],
    edges: [], titleBlockData: { architectureName: 'App', author: '', date: '', version: '1' },
    workflow: [], architecturePrompt: '', originalPrompt: '',
    settings: { pricingMode: 'payg', stylePreset: 'detailed', edgeStyle: 'orthogonal' },
  };
}

test('manual delete and document edits undo and redo without mutating snapshots', () => {
  const before = document();
  const history = new EditorHistory(before);
  const deleted = { ...before, nodes: [] };
  history.record(deleted);
  before.nodes[0].data.label = 'Not part of the saved state';
  assert.equal(history.undo()!.nodes[0].data.label, 'Web');
  assert.equal(history.redo()!.nodes.length, 0);
  history.record(document('Replacement'));
  assert.equal(history.canRedo, false);
  assert.equal(history.undo()!.nodes.length, 0);
});

test('selection, measurements, callback functions and key order do not create edits', () => {
  const before = document();
  const history = new EditorHistory(before);
  const measured = document();
  measured.nodes[0] = { ...measured.nodes[0], selected: true, width: 220, height: 160 };
  measured.nodes[0].data = { onEdit: () => {}, label: 'Web' };
  assert.equal(editorFingerprint(before), editorFingerprint(measured));
  assert.equal(history.record(measured), false);
  assert.equal(history.canUndo, false);
  measured.nodes[0].style = { width: 300 };
  assert.equal(history.record(measured), true);
});

test('region selection and 100,000-unit prices restore together as one document edit', () => {
  const before = document();
  before.settings.pricingRegion = 'eastus2';
  before.nodes[0].data.pricing = { estimatedCost: 20, quantity: 100_000, region: 'eastus2' };
  const history = new EditorHistory(before);
  const after = document();
  after.settings.pricingRegion = 'japaneast';
  after.nodes[0].data.pricing = { estimatedCost: 25, quantity: 100_000, region: 'japaneast' };
  history.record(after);
  assert.deepEqual(history.undo(), before);
  assert.equal(history.canUndo, false);
  assert.deepEqual(history.redo(), after);
});

test('drag and resize gestures create one reversible operation', () => {
  const before = document();
  const history = new EditorHistory(before);
  for (let x = 40; x <= 200; x += 20) {
    const next = document();
    next.nodes[0].position.x = x;
    next.nodes[0].dragging = true;
    history.record(next, true);
  }
  const final = document();
  final.nodes[0].position.x = 200;
  history.record(final);
  assert.equal(history.undo()!.nodes[0].position.x, 20);
  assert.equal(history.canUndo, false);
  assert.equal(history.redo()!.nodes[0].position.x, 200);
});

test('undo finishes an active gesture and new edits discard the redo branch', () => {
  const history = new EditorHistory(document());
  history.record(document('Typing'), true);
  assert.equal(history.undo()!.nodes[0].data.label, 'Web');
  assert.equal(history.canRedo, true);
  history.record(document('Different'));
  assert.equal(history.canRedo, false);
});

test('history is bounded and resets cleanly after recovery', () => {
  const history = new EditorHistory(document(), 2);
  for (const label of ['one', 'two', 'three']) history.record(document(label));
  assert.equal(history.undo()!.nodes[0].data.label, 'two');
  assert.equal(history.undo()!.nodes[0].data.label, 'one');
  assert.equal(history.undo(), null);
  history.reset(document('Restored'));
  assert.equal(history.canUndo, false);
  assert.equal(history.canRedo, false);
});

test('deleting a nested group preserves child absolute positions and valid edges', () => {
  const nodes = [
    { id: 'outer', type: 'groupNode', position: { x: -200, y: 100 }, data: {} },
    { id: 'inner', type: 'groupNode', parentNode: 'outer', position: { x: 30, y: 50 }, selected: true, data: {} },
    { id: 'web', type: 'azureNode', parentNode: 'inner', position: { x: 20, y: 40 }, data: {} },
    { id: 'db', type: 'azureNode', position: { x: 600, y: 100 }, data: {} },
  ];
  const edges = [{ id: 'query', source: 'web', target: 'db' }, { id: 'removed', source: 'inner', target: 'db' }];
  const result = deleteSelectedElements(nodes, edges);
  assert.deepEqual(result.nodes.find(node => node.id === 'web')!.position, { x: 50, y: 90 });
  assert.equal(result.nodes.find(node => node.id === 'web')!.parentNode, 'outer');
  assert.deepEqual(result.edges.map(edge => edge.id), ['query']);
  assert.equal(nodes[2].parentNode, 'inner');
});

test('duplicating groups remaps descendants and internal connections as one graph change', () => {
  const nodes = [
    { id: 'group', type: 'groupNode', position: { x: 50, y: 20 }, selected: true, data: {} },
    { id: 'web', type: 'azureNode', parentNode: 'group', position: { x: 30, y: 40 }, data: { label: 'Web' } },
    { id: 'db', type: 'azureNode', parentNode: 'group', position: { x: 250, y: 40 }, data: { label: 'DB' } },
  ];
  let sequence = 0;
  const result = duplicateSelectedElements(nodes, [{ id: 'query', source: 'web', target: 'db' }], () => `copy-${sequence++}`);
  assert.equal(result.nodes.length, 6);
  assert.deepEqual(result.nodes[3].position, { x: 100, y: 70 });
  assert.equal(result.nodes[4].parentNode, 'copy-0');
  assert.deepEqual(result.nodes[4].position, { x: 30, y: 40 });
  assert.equal(result.edges[1].source, 'copy-1');
  assert.equal(result.edges[1].target, 'copy-2');
  assert.equal(result.nodes[0].selected, false);
  assert.equal(nodes[0].selected, true);
});

test('automatic price completion enriches the add operation without adding an undo step', () => {
  const empty = { ...document(), nodes: [] };
  const history = new EditorHistory(empty);
  history.record(document());
  const priced = document();
  priced.nodes[0].data.pricing = { estimatedCost: 30, quantity: 1, isCustom: false };
  history.record(priced);
  assert.equal(history.undo()!.nodes.length, 0);
  assert.equal(history.redo()!.nodes[0].data.pricing.estimatedCost, 30);
  const custom = document();
  custom.nodes[0].data.pricing = { estimatedCost: 50, quantity: 1, isCustom: true };
  history.record(custom);
  assert.equal(history.undo()!.nodes[0].data.pricing.estimatedCost, 30);
});

test('a late automatic price enriches redo even when the added node was already undone', () => {
  const before = document();
  before.settings.pricingRegion = 'eastus2';
  const history = new EditorHistory({ ...before, nodes: [] });
  history.record(before);
  history.undo();
  const revision = history.revision;
  history.enrichPricing(before.nodes[0], {
    estimatedCost: 30, quantity: 1, isCustom: false, region: 'eastus2',
    tier: 'Standard', skuName: 'Standard', unit: 'instance', lastUpdated: '2026-09-05',
  });
  assert.equal(history.revision, revision, 'Automatic completion is not a user edit.');
  assert.equal(history.current.nodes.length, 0);
  assert.equal(history.redo()!.nodes[0].data.pricing.estimatedCost, 30);
  assert.equal(history.undo()!.nodes.length, 0);
  assert.equal(history.canUndo, false);
});

test('automatic prices cannot leak across an identity or region change', () => {
  const before = document();
  before.settings.pricingRegion = 'eastus2';
  before.nodes[0].data.serviceName = 'App Service';
  const history = new EditorHistory(before);
  const changed = document();
  changed.settings.pricingRegion = 'japaneast';
  changed.nodes[0].data.serviceName = 'SQL Database';
  history.record(changed);
  history.enrichPricing(changed.nodes[0], {
    estimatedCost: 50, quantity: 1, isCustom: false, region: 'japaneast',
    tier: 'Standard', skuName: 'Standard', unit: 'database', lastUpdated: '2026-09-05',
  });
  assert.equal(history.current.nodes[0].data.pricing.estimatedCost, 50);
  assert.equal(history.undo()!.nodes[0].data.pricing, undefined);
});

test('a simultaneous rename and automatic enrichment leaves one reversible rename', () => {
  const before = document();
  before.nodes[0].data.serviceName = 'App Service';
  const history = new EditorHistory(before);
  const renamed = document('Portal');
  renamed.nodes[0].data.serviceName = 'App Service';
  renamed.nodes[0].data.pricing = { estimatedCost: 12, quantity: 1, isCustom: false };
  history.record(renamed);
  const restored = history.undo()!;
  assert.equal(restored.nodes[0].data.label, 'Web');
  assert.equal(restored.nodes[0].data.pricing.estimatedCost, 12);
  assert.equal(history.canUndo, false);
});

test('review observations do not enter or get lost through the edit history', () => {
  const before = document();
  const history = new EditorHistory(before);
  history.record(document('Portal'));
  const reviewed = { ...document('Portal'), reviewHistory: [{ key: 'observed' }], validationScore: 80, validationSourceFingerprint: 'source' };
  assert.equal(history.record(reviewed), false);
  const restored = history.undo()!;
  assert.equal(restored.nodes[0].data.label, 'Web');
  assert.deepEqual(restored.reviewHistory, reviewed.reviewHistory);
  assert.equal(restored.validationSourceFingerprint, 'source');
  assert.equal(history.canUndo, false);
});
