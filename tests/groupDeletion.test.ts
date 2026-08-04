import test from 'node:test';
import assert from 'node:assert/strict';
import type { Node } from 'reactflow';
import {
  collectNodeAndDescendantIds,
  deleteNodesPreservingGroupChildren,
  detachNodeFromGroup,
} from '../src/utils/groupUtils.ts';

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

test('deleting a group preserves its child nodes at absolute positions', () => {
  const result = deleteNodesPreservingGroupChildren([
    node('group', 'groupNode', 250, 150),
    node('service', 'azureNode', 40, 60, 'group'),
    node('outside', 'azureNode', 700, 100),
  ], ['group']);

  assert.deepEqual(result.map(item => item.id), ['service', 'outside']);
  const service = result.find(item => item.id === 'service');
  assert.deepEqual(service?.position, { x: 290, y: 210 });
  assert.equal(service?.parentNode, undefined);
  assert.equal(service?.extent, undefined);
});

test('explicitly selected children are still deleted with their group', () => {
  const result = deleteNodesPreservingGroupChildren([
    node('group', 'groupNode', 250, 150),
    node('service', 'azureNode', 40, 60, 'group'),
    node('outside', 'azureNode', 700, 100),
  ], ['group', 'service']);

  assert.deepEqual(result.map(item => item.id), ['outside']);
});

test('nested groups retain their descendants when the outer group is deleted', () => {
  const result = deleteNodesPreservingGroupChildren([
    node('outer', 'groupNode', 100, 100),
    node('inner', 'groupNode', 30, 40, 'outer'),
    node('service', 'azureNode', 5, 6, 'inner'),
  ], ['outer']);

  const inner = result.find(item => item.id === 'inner');
  const service = result.find(item => item.id === 'service');
  assert.deepEqual(inner?.position, { x: 130, y: 140 });
  assert.equal(inner?.parentNode, undefined);
  assert.equal(service?.parentNode, 'inner');
  assert.deepEqual(service?.position, { x: 5, y: 6 });
});

test('collectNodeAndDescendantIds includes nested group contents only', () => {
  const ids = collectNodeAndDescendantIds([
    node('outer', 'groupNode', 100, 100),
    node('inner', 'groupNode', 30, 40, 'outer'),
    node('service', 'azureNode', 5, 6, 'inner'),
    node('outside', 'azureNode', 700, 100),
  ], ['outer']);

  assert.deepEqual([...ids].sort(), ['inner', 'outer', 'service']);
});

test('ungrouping a service preserves its absolute position through nested groups', () => {
  const result = detachNodeFromGroup([
    node('outer', 'groupNode', 100, 100),
    node('inner', 'groupNode', 30, 40, 'outer'),
    node('service', 'azureNode', 5, 6, 'inner'),
  ], 'service');

  const service = result.find(item => item.id === 'service');
  assert.deepEqual(service?.position, { x: 135, y: 146 });
  assert.equal(service?.parentNode, undefined);
  assert.equal(service?.extent, undefined);
});
