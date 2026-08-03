import assert from 'node:assert/strict';
import test from 'node:test';
import type { Edge, Node } from 'reactflow';
import { duplicateSelectedSubgraph } from '../src/utils/groupUtils.ts';

function node(
  id: string,
  type: string,
  x: number,
  y: number,
  parentNode?: string,
): Node {
  return {
    id,
    type,
    position: { x, y },
    data: { label: id },
    parentNode,
    extent: parentNode ? 'parent' : undefined,
  } as Node;
}

test('duplicating a group copies nested descendants and internal edges once', () => {
  const nodes = [
    { ...node('outer', 'groupNode', 100, 100), selected: true },
    node('inner', 'groupNode', 20, 30, 'outer'),
    node('service-a', 'azureNode', 10, 15, 'inner'),
    node('service-b', 'azureNode', 200, 80, 'outer'),
    node('outside', 'azureNode', 700, 100),
  ];
  const edges: Edge[] = [
    { id: 'internal', source: 'service-a', target: 'service-b' },
    { id: 'external', source: 'service-b', target: 'outside' },
  ];
  let sequence = 0;

  const result = duplicateSelectedSubgraph(
    nodes,
    edges,
    ['outer'],
    (kind, sourceId) => `${kind}-${sourceId}-${++sequence}`,
  );

  assert.equal(result.nodes.length, 9);
  assert.equal(result.edges.length, 3);
  const outerCopy = result.nodes.find(item => item.id.startsWith('node-outer-'));
  const innerCopy = result.nodes.find(item => item.id.startsWith('node-inner-'));
  const serviceACopy = result.nodes.find(item => item.id.startsWith('node-service-a-'));
  const serviceBCopy = result.nodes.find(item => item.id.startsWith('node-service-b-'));
  assert.deepEqual(outerCopy?.position, { x: 150, y: 150 });
  assert.deepEqual(innerCopy?.position, { x: 20, y: 30 });
  assert.equal(innerCopy?.parentNode, outerCopy?.id);
  assert.equal(serviceACopy?.parentNode, innerCopy?.id);
  assert.equal(serviceBCopy?.parentNode, outerCopy?.id);
  assert.equal(outerCopy?.selected, true);
  assert.equal(innerCopy?.selected, false);

  const internalCopy = result.edges.find(item => item.id.startsWith('edge-internal-'));
  assert.equal(internalCopy?.source, serviceACopy?.id);
  assert.equal(internalCopy?.target, serviceBCopy?.id);
  assert.equal(result.edges.some(item => item.id.startsWith('edge-external-')), false);
});

test('duplicating one child keeps it in the original group with a relative offset', () => {
  const result = duplicateSelectedSubgraph(
    [
      node('group', 'groupNode', 100, 100),
      { ...node('service', 'azureNode', 20, 30, 'group'), selected: true },
    ],
    [],
    ['service'],
    (_kind, sourceId) => `${sourceId}-copy`,
  );

  const copy = result.nodes.find(item => item.id === 'service-copy');
  assert.equal(copy?.parentNode, 'group');
  assert.deepEqual(copy?.position, { x: 70, y: 80 });
});
