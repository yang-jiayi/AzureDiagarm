import assert from 'node:assert/strict';
import test from 'node:test';
import type { Edge, Node } from 'reactflow';
import { mergeLayoutEdges, mergeLayoutNodes } from '../src/utils/layoutResultMerge.ts';

function node(id: string, x: number, label = id): Node {
  return {
    id,
    type: 'azureNode',
    position: { x, y: 0 },
    data: { label },
  } as Node;
}

test('layout merge preserves concurrent graph edits without resurrecting deleted nodes', () => {
  const source = [node('a', 0), node('deleted', 10)];
  const laidOut = [node('a', 100), node('deleted', 200)];
  const current = [
    { ...node('a', 0, 'Renamed'), selected: true },
    node('new', 300),
  ];

  const merged = mergeLayoutNodes(current, source, laidOut);
  assert.deepEqual(merged.map(item => item.id), ['a', 'new']);
  assert.deepEqual(merged[0].position, { x: 100, y: 0 });
  assert.equal(merged[0].data.label, 'Renamed');
  assert.equal(merged[0].selected, true);
  assert.deepEqual(merged[1].position, { x: 300, y: 0 });
});

test('layout merge does not overwrite geometry moved during asynchronous layout', () => {
  const source = [node('a', 0)];
  const laidOut = [node('a', 100)];
  const current = [node('a', 25)];

  const merged = mergeLayoutNodes(current, source, laidOut);
  assert.deepEqual(merged[0].position, { x: 25, y: 0 });
});

test('edge layout merge updates layout styling while preserving concurrent labels', () => {
  const source: Edge[] = [{
    id: 'edge',
    source: 'a',
    target: 'b',
    label: 'old',
    data: { pathStyle: 'straight' },
  }];
  const laidOut: Edge[] = [{
    ...source[0],
    data: { pathStyle: 'orthogonal', primaryPath: true },
  }];
  const current: Edge[] = [{
    ...source[0],
    label: 'renamed',
  }, {
    id: 'new-edge',
    source: 'b',
    target: 'c',
  }];

  const merged = mergeLayoutEdges(current, source, laidOut);
  assert.equal(merged[0].label, 'renamed');
  assert.deepEqual(merged[0].data, {
    pathStyle: 'orthogonal',
    primaryPath: true,
  });
  assert.equal(merged[1], current[1]);
});
