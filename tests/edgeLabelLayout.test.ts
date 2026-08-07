// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

import assert from 'node:assert/strict';
import test from 'node:test';
import type { Edge, Node } from 'reactflow';
import {
  applyAutomaticEdgeLabelOffsets,
  calculateAutomaticEdgeLabelOffsets,
  shouldRecalculateAutomaticEdgeLabels,
} from '../src/utils/edgeLabelLayout';

function service(id: string, x: number, y: number): Node {
  return {
    id,
    type: 'azureNode',
    position: { x, y },
    width: 180,
    height: 110,
    data: { label: id },
  };
}

function connection(id: string, source: string, target: string, label: string): Edge {
  return {
    id,
    source,
    target,
    label,
    data: { labelOffsetAuto: true },
  };
}

test('automatic edge labels move away from a service at the path midpoint', () => {
  const nodes = [
    service('source', 0, 100),
    service('blocker', 300, 100),
    service('target', 600, 100),
  ];
  const offsets = calculateAutomaticEdgeLabelOffsets(
    nodes,
    [connection('edge', 'source', 'target', 'Send request')],
  );

  const offset = offsets.get('edge');
  assert.ok(offset);
  assert.ok(Math.abs(offset.x) > 0 || Math.abs(offset.y) > 0);
});

test('parallel relationship labels receive separate automatic offsets', () => {
  const nodes = [service('source', 0, 0), service('target', 500, 0)];
  const offsets = calculateAutomaticEdgeLabelOffsets(nodes, [
    connection('first', 'source', 'target', 'Read data'),
    connection('second', 'source', 'target', 'Write data'),
  ]);

  assert.notDeepEqual(offsets.get('first'), offsets.get('second'));
});

test('manual label positions are preserved when automatic layout is reapplied', () => {
  const nodes = [service('source', 0, 0), service('target', 500, 0)];
  const edge = connection('edge', 'source', 'target', 'Read data');
  edge.data = {
    labelOffsetX: 90,
    labelOffsetY: 40,
    labelOffsetAuto: false,
  };

  const [result] = applyAutomaticEdgeLabelOffsets(nodes, [edge]);
  assert.equal(result.data?.labelOffsetX, 90);
  assert.equal(result.data?.labelOffsetY, 40);
});

test('automatic edge labels are refreshed only after completed geometry changes', () => {
  assert.equal(shouldRecalculateAutomaticEdgeLabels([
    { id: 'node', type: 'position', position: { x: 20, y: 20 }, dragging: true },
  ]), false);
  assert.equal(shouldRecalculateAutomaticEdgeLabels([
    { id: 'node', type: 'position', position: { x: 20, y: 20 }, dragging: false },
  ]), true);
  assert.equal(shouldRecalculateAutomaticEdgeLabels([
    { id: 'node', type: 'dimensions', dimensions: { width: 200, height: 120 }, resizing: false },
  ]), true);
  assert.equal(shouldRecalculateAutomaticEdgeLabels([
    { id: 'node', type: 'select', selected: true },
  ]), false);
});

test('automatic edge label offsets adapt after connected nodes move', () => {
  const initialNodes = [
    service('source', 0, 100),
    service('blocker', 300, 100),
    service('target', 600, 100),
  ];
  const edge = connection('edge', 'source', 'target', 'Send request');
  const [initial] = applyAutomaticEdgeLabelOffsets(initialNodes, [edge]);
  const movedNodes = initialNodes.map(node => (
    node.id === 'target' ? { ...node, position: { x: 0, y: 600 } } : node
  ));
  const [moved] = applyAutomaticEdgeLabelOffsets(movedNodes, [initial]);

  assert.notDeepEqual(
    [initial.data?.labelOffsetX, initial.data?.labelOffsetY],
    [moved.data?.labelOffsetX, moved.data?.labelOffsetY],
  );
  assert.equal(
    Math.hypot(
      Number(moved.data?.labelOffsetX),
      Number(moved.data?.labelOffsetY),
    ),
    0,
  );
});
