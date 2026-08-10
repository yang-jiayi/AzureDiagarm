// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

import assert from 'node:assert/strict';
import test from 'node:test';
import type { Edge, Node } from 'reactflow';
import {
  analyzeDiagramQuality,
  applyDiagramQualityFixes,
} from '../src/utils/diagramQuality';

function node(
  id: string,
  x: number,
  y: number,
  label = id,
  extra: Partial<Node> = {},
): Node {
  return {
    id,
    type: 'azureNode',
    position: { x, y },
    data: { label },
    ...extra,
  };
}

function edge(id: string, source: string, target: string): Edge {
  return { id, source, target, type: 'editableEdge' };
}

test('quality analysis detects overlaps crossings orphans crowded labels and contrast', () => {
  const nodes = [
    node('a', 0, 0, 'A', { style: { color: '#777777', background: '#888888' } }),
    node('b', 20, 20, 'B'),
    node('c', 0, 240, 'A very long service label without enough room to remain readable'),
    node('d', 240, 0, 'D'),
    node('e', 240, 240, 'E'),
    node('orphan', 520, 0, 'Orphan'),
  ];
  const edges = [
    edge('ae', 'a', 'e'),
    edge('cd', 'c', 'd'),
  ];
  const report = analyzeDiagramQuality(nodes, edges);
  const categories = new Set(report.findings.map(finding => finding.category));

  assert.equal(categories.has('overlap'), true);
  assert.equal(categories.has('crossing'), true);
  assert.equal(categories.has('orphan'), true);
  assert.equal(categories.has('label'), true);
  assert.equal(categories.has('contrast'), true);
});

test('quality analysis detects density and insufficient group padding', () => {
  const denseNodes = Array.from({ length: 12 }, (_, index) => (
    node(`n${index}`, (index % 4) * 90, Math.floor(index / 4) * 80)
  ));
  const group: Node = {
    id: 'group',
    type: 'groupNode',
    position: { x: 600, y: 0 },
    style: { width: 260, height: 220 },
    data: { label: 'Group' },
  };
  const child = node('child', 4, 30, 'Child', { parentNode: 'group', extent: 'parent' });
  const report = analyzeDiagramQuality([...denseNodes, group, child], []);
  const categories = new Set(report.findings.map(finding => finding.category));

  assert.equal(categories.has('density'), true);
  assert.equal(categories.has('group-padding'), true);
});

test('quality fixes expand labels fit groups and reset custom contrast', () => {
  const group: Node = {
    id: 'group',
    type: 'groupNode',
    position: { x: 0, y: 0 },
    style: { width: 240, height: 180 },
    data: { label: 'Group' },
  };
  const child = node(
    'child',
    2,
    20,
    'A very long service label without enough room to remain readable',
    {
      parentNode: 'group',
      extent: 'parent',
      style: { color: '#777777', background: '#888888' },
    },
  );
  const nodes = [group, child];
  const report = analyzeDiagramQuality(nodes, []);
  const selected = report.findings
    .filter(finding => finding.fixKind && finding.fixKind !== 'layout')
    .map(finding => finding.id);
  const fixed = applyDiagramQualityFixes(nodes, [], report.findings, selected);
  const fixedChild = fixed.nodes.find(item => item.id === 'child');
  const fixedGroup = fixed.nodes.find(item => item.id === 'group');

  assert.ok(Number(fixedChild?.data.labelMaxWidth) >= 160);
  assert.equal((fixedChild?.style as Record<string, unknown>)?.color, undefined);
  assert.ok(Number((fixedGroup?.style as Record<string, unknown>)?.width) > 240);
});
