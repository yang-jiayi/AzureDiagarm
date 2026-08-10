// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

import assert from 'node:assert/strict';
import test from 'node:test';
import {
  orderBlueprintColumns,
  type BlueprintArchitecture,
  type BpNode,
} from '../src/services/blueprintArchitectureAI';

const NODE_W = 180;
const NODE_H = 120;

type Point = { x: number; y: number };
const center = (node: BpNode): Point => ({ x: node.x + NODE_W / 2, y: node.y + NODE_H / 2 });

function orientation(a: Point, b: Point, c: Point): number {
  return (b.y - a.y) * (c.x - b.x) - (b.x - a.x) * (c.y - b.y);
}

function segmentsCross(a: Point, b: Point, c: Point, d: Point): boolean {
  const o1 = orientation(a, b, c);
  const o2 = orientation(a, b, d);
  const o3 = orientation(c, d, a);
  const o4 = orientation(c, d, b);
  return (
    ((o1 > 0 && o2 < 0) || (o1 < 0 && o2 > 0))
    && ((o3 > 0 && o4 < 0) || (o3 < 0 && o4 > 0))
  );
}

/** Count edge-pair crossings, ignoring edges that share an endpoint. */
function countCrossings(bp: BlueprintArchitecture): number {
  const byId = new Map(bp.nodes.map(node => [node.id, node]));
  const edges = bp.edges.filter(
    edge => byId.has(edge.from) && byId.has(edge.to) && edge.from !== edge.to,
  );
  let count = 0;
  for (let i = 0; i < edges.length; i += 1) {
    for (let j = i + 1; j < edges.length; j += 1) {
      const left = edges[i];
      const right = edges[j];
      if (
        left.from === right.from
        || left.from === right.to
        || left.to === right.from
        || left.to === right.to
      ) continue;
      if (segmentsCross(
        center(byId.get(left.from)!),
        center(byId.get(left.to)!),
        center(byId.get(right.from)!),
        center(byId.get(right.to)!),
      )) count += 1;
    }
  }
  return count;
}

// A backbone chain n0→…→n5 with two extra forward edges, whose X positions the
// model scrambled so downstream tiles sit left of upstream ones. This exact
// input was verified to produce two segment crossings before ordering.
function scrambledCrossingCase(): BlueprintArchitecture {
  return {
    title: 'Crossing case',
    canvas: { width: 1600, height: 1000 },
    zones: [],
    nodes: [
      { id: 'n0', name: 'n0', category: 'compute', x: 750, y: 250, zone: 'z' },
      { id: 'n1', name: 'n1', category: 'compute', x: 250, y: 0, zone: 'z' },
      { id: 'n2', name: 'n2', category: 'compute', x: 1000, y: 0, zone: 'z' },
      { id: 'n3', name: 'n3', category: 'compute', x: 500, y: 0, zone: 'z' },
      { id: 'n4', name: 'n4', category: 'compute', x: 500, y: 250, zone: 'z' },
      { id: 'n5', name: 'n5', category: 'compute', x: 1250, y: 0, zone: 'z' },
    ],
    edges: [
      { id: 'e0', from: 'n0', to: 'n1' },
      { id: 'e1', from: 'n1', to: 'n2' },
      { id: 'e2', from: 'n2', to: 'n3' },
      { id: 'e3', from: 'n3', to: 'n4' },
      { id: 'e4', from: 'n4', to: 'n5' },
      { id: 'x1', from: 'n0', to: 'n4' },
      { id: 'x2', from: 'n1', to: 'n5' },
    ],
  };
}

function positionsSignature(bp: BlueprintArchitecture): string {
  return bp.nodes.map(node => `${node.id}:${node.x},${node.y}`).join('|');
}

test('orderBlueprintColumns is deterministic — identical input yields identical output', () => {
  const first = scrambledCrossingCase();
  const second = scrambledCrossingCase();
  orderBlueprintColumns(first);
  orderBlueprintColumns(second);
  assert.equal(positionsSignature(first), positionsSignature(second));
});

test('orderBlueprintColumns is idempotent — a second pass changes nothing', () => {
  const bp = scrambledCrossingCase();
  orderBlueprintColumns(bp);
  const once = positionsSignature(bp);
  orderBlueprintColumns(bp);
  assert.equal(positionsSignature(bp), once);
});

test('orderBlueprintColumns reduces crossings on a known crossing case', () => {
  const bp = scrambledCrossingCase();
  const before = countCrossings(bp);
  orderBlueprintColumns(bp);
  const after = countCrossings(bp);
  assert.ok(before >= 2, `expected the scrambled case to start with crossings, got ${before}`);
  assert.ok(after < before, `expected fewer crossings after ordering (${after} < ${before})`);
  assert.equal(after, 0, 'expected the known case to become crossing-free');
});

test('orderBlueprintColumns preserves node set, Y banding, and the zone X-slot set', () => {
  const original = scrambledCrossingCase();
  const bp = scrambledCrossingCase();
  orderBlueprintColumns(bp);

  // Same nodes and ids, in the same array order.
  assert.deepEqual(bp.nodes.map(n => n.id), original.nodes.map(n => n.id));

  // Every node keeps its original Y (hot/cool banding is preserved).
  for (const node of bp.nodes) {
    const before = original.nodes.find(n => n.id === node.id)!;
    assert.equal(node.y, before.y, `Y of ${node.id} must be preserved`);
  }

  // The multiset of X positions within the zone is unchanged, so the zone's
  // bounding box does not grow or shift — only the assignment of tiles to
  // columns changes.
  const sortedX = (nodes: BpNode[]) => nodes.map(n => n.x).sort((a, b) => a - b);
  assert.deepEqual(sortedX(bp.nodes), sortedX(original.nodes));

  // Downstream tiles now sit at or right of their upstream predecessors along
  // the main chain (no backward edges on the backbone).
  const xOf = (id: string) => bp.nodes.find(n => n.id === id)!.x;
  assert.ok(xOf('n0') <= xOf('n1'));
  assert.ok(xOf('n1') <= xOf('n2'));
  assert.ok(xOf('n2') <= xOf('n3'));
});

test('orderBlueprintColumns leaves a trivial (single-node) diagram untouched', () => {
  const bp: BlueprintArchitecture = {
    title: 'trivial',
    canvas: { width: 1600, height: 1000 },
    zones: [],
    nodes: [{ id: 'only', name: 'Solo', category: 'compute', x: 42, y: 99 }],
    edges: [],
  };
  orderBlueprintColumns(bp);
  assert.equal(bp.nodes[0].x, 42);
  assert.equal(bp.nodes[0].y, 99);
});
