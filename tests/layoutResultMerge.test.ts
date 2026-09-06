import assert from 'node:assert/strict';
import test from 'node:test';
import type { Edge, Node } from 'reactflow';
import { mergeLayoutEdges, mergeLayoutNodes } from '../src/utils/layoutResultMerge.ts';
import { OperationGeneration } from '../src/utils/operationGeneration.ts';

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

test('edge layout merge applies automatic label offsets without replacing manual edits', () => {
  const source: Edge[] = [{
    id: 'edge',
    source: 'a',
    target: 'b',
    data: {
      labelOffsetX: 0,
      labelOffsetY: 0,
      labelOffsetAuto: true,
    },
  }];
  const laidOut: Edge[] = [{
    ...source[0],
    data: {
      labelOffsetX: 24,
      labelOffsetY: -18,
      labelOffsetAuto: true,
    },
  }];

  const automatic = mergeLayoutEdges([{ ...source[0] }], source, laidOut);
  assert.deepEqual(automatic[0].data, laidOut[0].data);

  const manual = mergeLayoutEdges([{
    ...source[0],
    data: {
      labelOffsetX: 40,
      labelOffsetY: 12,
      labelOffsetAuto: false,
    },
  }], source, laidOut);
  assert.deepEqual(manual[0].data, {
    labelOffsetX: 40,
    labelOffsetY: 12,
    labelOffsetAuto: false,
  });
});

test('a superseded asynchronous layout cannot replace a newer layout result', async () => {
  const generation = new OperationGeneration();
  let currentNodes = [node('a', 0)];
  let finishFirst!: (nodes: Node[]) => void;
  let finishSecond!: (nodes: Node[]) => void;
  const first = new Promise<Node[]>(resolve => { finishFirst = resolve; });
  const second = new Promise<Node[]>(resolve => { finishSecond = resolve; });
  const apply = async (pending: Promise<Node[]>) => {
    const sourceNodes = currentNodes;
    const token = generation.advance();
    const result = await pending;
    if (!generation.isCurrent(token)) return;
    currentNodes = mergeLayoutNodes(currentNodes, sourceNodes, result);
  };

  const firstApply = apply(first);
  const secondApply = apply(second);
  finishSecond([node('a', 200)]);
  await secondApply;
  finishFirst([node('a', 100)]);
  await firstApply;
  assert.deepEqual(currentNodes[0].position, { x: 200, y: 0 });
});

test('pending layout preserves drag, rename, addition, deletion and manual edge edits together', async () => {
  const sourceNodes = [node('a', 0), node('b', 20), node('removed', 40)];
  const sourceEdges: Edge[] = [{
    id: 'edge', source: 'a', target: 'b',
    sourceHandle: 'right', targetHandle: 'left', label: 'before',
    data: { pathStyle: 'straight', labelOffsetX: 0, labelOffsetY: 0, labelOffsetAuto: true },
  }];
  let currentNodes = sourceNodes;
  let currentEdges = sourceEdges;
  let finish!: (result: { nodes: Node[]; edges: Edge[] }) => void;
  const pending = new Promise<{ nodes: Node[]; edges: Edge[] }>(resolve => { finish = resolve; });
  const apply = pending.then(result => {
    currentNodes = mergeLayoutNodes(currentNodes, sourceNodes, result.nodes);
    currentEdges = mergeLayoutEdges(currentEdges, sourceEdges, result.edges);
  });

  currentNodes = [
    { ...node('a', 35, 'Renamed while arranging'), selected: true },
    node('b', 20),
    node('added', 500),
  ];
  currentEdges = [{
    ...sourceEdges[0],
    sourceHandle: 'bottom',
    label: 'Edited while arranging',
    data: { pathStyle: 'smooth', labelOffsetX: 40, labelOffsetY: 30, labelOffsetAuto: false },
  }];
  finish({
    nodes: [node('a', 100), node('b', 200), node('removed', 300)],
    edges: [{
      ...sourceEdges[0],
      sourceHandle: 'left-source',
      targetHandle: 'right-target',
      data: { pathStyle: 'orthogonal', labelOffsetX: -10, labelOffsetY: -20, labelOffsetAuto: true },
    }],
  });
  await apply;

  assert.deepEqual(currentNodes.map(item => item.id), ['a', 'b', 'added']);
  assert.equal(currentNodes[0].data.label, 'Renamed while arranging');
  assert.equal(currentNodes[0].selected, true);
  assert.deepEqual(currentNodes.map(item => item.position.x), [35, 200, 500]);
  assert.equal(currentEdges[0].label, 'Edited while arranging');
  assert.equal(currentEdges[0].sourceHandle, 'bottom');
  assert.equal(currentEdges[0].targetHandle, 'left');
  assert.deepEqual(currentEdges[0].data, {
    pathStyle: 'smooth', labelOffsetX: 40, labelOffsetY: 30, labelOffsetAuto: false,
  });
});

test('an in-flight layout is ignored after replacing the document even when IDs and positions match', async () => {
  let lineage = 'original';
  let currentNodes = [node('a', 0)];
  const sourceNodes = currentNodes;
  const sourceLineage = lineage;
  let finish!: (nodes: Node[]) => void;
  const pending = new Promise<Node[]>(resolve => { finish = resolve; });
  const apply = pending.then(result => {
    if (lineage !== sourceLineage) return;
    currentNodes = mergeLayoutNodes(currentNodes, sourceNodes, result);
  });
  lineage = 'loaded-document';
  currentNodes = [node('a', 0, 'New document')];
  finish([node('a', 100)]);
  await apply;
  assert.equal(currentNodes[0].data.label, 'New document');
  assert.deepEqual(currentNodes[0].position, { x: 0, y: 0 });
});
