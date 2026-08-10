import assert from 'node:assert/strict';
import test from 'node:test';
import type { Node } from 'reactflow';
import {
  alignSelectedNodes,
  applyBulkNodeEdits,
  BULK_GROUP_COLORS,
  normalizeBulkTags,
} from '../src/utils/bulkNodeEditing';

test('alignment uses absolute positions across different parent groups', () => {
  const nodes: Node[] = [
    { id: 'g1', type: 'groupNode', position: { x: 100, y: 100 }, data: {} },
    { id: 'g2', type: 'groupNode', position: { x: 500, y: 200 }, data: {} },
    {
      id: 'a',
      type: 'azureNode',
      parentNode: 'g1',
      position: { x: 30, y: 60 },
      selected: true,
      data: {},
    },
    {
      id: 'b',
      type: 'azureNode',
      parentNode: 'g2',
      position: { x: 20, y: 40 },
      selected: true,
      data: {},
    },
  ];
  const aligned = alignSelectedNodes(nodes, 'left');
  const a = aligned.find(node => node.id === 'a')!;
  const b = aligned.find(node => node.id === 'b')!;
  assert.equal(100 + a.position.x, 500 + b.position.x);
});

test('bulk group assignment preserves absolute position and fits the group', () => {
  const nodes: Node[] = [
    {
      id: 'group',
      type: 'groupNode',
      position: { x: 300, y: 100 },
      style: { width: 220, height: 180 },
      data: { label: 'Target' },
    },
    {
      id: 'service',
      type: 'azureNode',
      position: { x: 80, y: 420 },
      data: { label: 'App' },
    },
  ];
  const updated = applyBulkNodeEdits(nodes, new Set(['service']), {
    targetGroupId: 'group',
  });
  const group = updated.find(node => node.id === 'group')!;
  const service = updated.find(node => node.id === 'service')!;
  assert.equal(service.parentNode, 'group');
  assert.equal(service.data.groupId, 'group');
  assert.equal(group.position.x + service.position.x, 80);
  assert.equal(group.position.y + service.position.y, 420);
});

test('bulk edits normalize tags and apply service and group styling', () => {
  const nodes: Node[] = [
    {
      id: 'group',
      type: 'groupNode',
      position: { x: 0, y: 0 },
      data: { tags: ['Legacy'] },
    },
    {
      id: 'service',
      type: 'azureNode',
      position: { x: 0, y: 0 },
      data: { tags: ['Production'] },
    },
  ];
  const updated = applyBulkNodeEdits(nodes, new Set(['group', 'service']), {
    stylePreset: 'presentation',
    groupColor: BULK_GROUP_COLORS[1],
    tags: {
      mode: 'add',
      values: ['Critical', ' production ', 'Critical'],
    },
  });
  const group = updated.find(node => node.id === 'group')!;
  const service = updated.find(node => node.id === 'service')!;
  assert.deepEqual(group.data.customColor, BULK_GROUP_COLORS[1]);
  assert.deepEqual(group.data.tags, ['Legacy', 'Critical', 'production']);
  assert.equal(service.data.stylePreset, 'presentation');
  assert.deepEqual(service.data.tags, ['Production', 'Critical']);
  assert.deepEqual(normalizeBulkTags([' A ', 'a', '', 'B']), ['A', 'B']);
});
