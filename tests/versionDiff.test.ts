// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

import assert from 'node:assert/strict';
import test from 'node:test';
import type { Edge, Node } from 'reactflow';
import {
  applySelectedVersionChanges,
  compareDiagramVersions,
} from '../src/utils/versionDiff';

function node(id: string, label: string, x = 0, parentNode?: string): Node {
  return {
    id,
    type: 'azureNode',
    position: { x, y: 0 },
    parentNode,
    extent: parentNode ? 'parent' : undefined,
    data: { label },
  };
}

function edge(id: string, source: string, target: string, label = ''): Edge {
  return { id, source, target, label, type: 'editableEdge' };
}

test('version diff reports added removed and changed elements while ignoring runtime state', () => {
  const currentNodes = [
    { ...node('a', 'App'), selected: true, data: { label: 'App', onClick: () => undefined } },
    node('b', 'Database'),
  ];
  const targetNodes = [
    node('a', 'App'),
    node('c', 'Cache'),
  ];
  const currentEdges = [edge('ab', 'a', 'b')];
  const targetEdges = [edge('ac', 'a', 'c')];

  const diff = compareDiagramVersions(currentNodes, currentEdges, targetNodes, targetEdges);

  assert.deepEqual(diff.counts, { added: 2, removed: 2, changed: 0 });
  assert.deepEqual(
    diff.items.map(item => [item.key, item.status]),
    [
      ['node:b', 'removed'],
      ['node:c', 'added'],
      ['edge:ab', 'removed'],
      ['edge:ac', 'added'],
    ],
  );
});

test('selective restore applies chosen elements and removes dangling edges safely', () => {
  const currentNodes = [node('a', 'App'), node('b', 'Database')];
  const targetNodes = [node('a', 'Updated App', 40), node('c', 'Cache')];
  const currentEdges = [edge('ab', 'a', 'b')];
  const targetEdges = [edge('ac', 'a', 'c')];

  const result = applySelectedVersionChanges(
    currentNodes,
    currentEdges,
    targetNodes,
    targetEdges,
    ['node:a', 'node:b', 'node:c', 'edge:ac'],
  );

  assert.deepEqual(result.nodes.map(item => item.id), ['a', 'c']);
  assert.equal(result.nodes[0].data.label, 'Updated App');
  assert.deepEqual(result.edges.map(item => item.id), ['ac']);
  assert.deepEqual(result.autoRemovedEdgeIds, ['ab']);
  assert.deepEqual(result.skippedKeys, []);
});

test('selective restore automatically includes a required historical parent group', () => {
  const group: Node = {
    id: 'group',
    type: 'groupNode',
    position: { x: 20, y: 20 },
    data: { label: 'Application tier' },
  };
  const child = node('child', 'App', 10, 'group');
  const result = applySelectedVersionChanges([], [], [group, child], [], ['node:child']);

  assert.deepEqual(result.nodes.map(item => item.id), ['group', 'child']);
  assert.deepEqual(result.autoAddedNodeIds, ['group']);
});
