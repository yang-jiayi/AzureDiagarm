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

  assert.ok(fixedChild, 'the child node was dropped by the fix pass');
  assert.ok(fixedGroup, 'the group node was dropped by the fix pass');
  assert.ok(Number(fixedChild.data.labelMaxWidth) >= 160);
  // A missing node satisfies `?.color === undefined` as readily as a cleared
  // one; the missing case is the regression.
  assert.equal((fixedChild.style as Record<string, unknown>)?.color, undefined);
  assert.ok(Number((fixedGroup.style as Record<string, unknown>)?.width) > 240);
});

test('quality analysis flags unrecognized services by name', () => {
  const known = node('known', 0, 0, 'API Management', {
    data: { label: 'API Management', serviceName: 'API Management', iconPath: '/x.svg' },
  });
  const unknown = node('unknown', 400, 0, 'Frobnicator 9000', {
    data: { label: 'Frobnicator 9000', serviceName: 'Frobnicator 9000', iconPath: '' },
  });
  const report = analyzeDiagramQuality([known, unknown], [edge('e', 'known', 'unknown')]);
  const finding = report.findings.find(item => item.category === 'unrecognized-service');
  assert.ok(finding, 'expected an unrecognized-service finding');
  assert.deepEqual(finding?.nodeIds, ['unknown']);
  assert.equal(finding?.detail.includes('Frobnicator 9000'), true);
  // A resolvable node with an icon path must NOT be flagged.
  assert.equal(
    report.findings.some(item => item.category === 'unrecognized-service' && item.nodeIds.includes('known')),
    false,
  );
});

test('quality analysis flags generic and empty edge labels', () => {
  const a = node('a', 0, 0, 'Azure Functions', { data: { label: 'Azure Functions', iconPath: '/a.svg' } });
  const b = node('b', 400, 0, 'Service Bus', { data: { label: 'Service Bus', iconPath: '/b.svg' } });
  const c = node('c', 800, 0, 'Azure Cosmos DB', { data: { label: 'Azure Cosmos DB', iconPath: '/c.svg' } });
  const generic: Edge = { id: 'g', source: 'a', target: 'b', type: 'editableEdge', label: 'connects to' };
  const empty: Edge = { id: 'em', source: 'b', target: 'c', type: 'editableEdge' };
  const specific: Edge = { id: 'sp', source: 'a', target: 'c', type: 'editableEdge', label: 'Persist result (Delta)' };
  const report = analyzeDiagramQuality([a, b, c], [generic, empty, specific]);
  const flagged = report.findings
    .filter(item => item.category === 'generic-edge')
    .flatMap(item => item.edgeIds)
    .sort();
  assert.deepEqual(flagged, ['em', 'g']);
});

test('quality analysis flags a service overlapping an unrelated group', () => {
  const groupA: Node = {
    id: 'groupA',
    type: 'groupNode',
    position: { x: 0, y: 0 },
    style: { width: 300, height: 300 },
    data: { label: 'Zone A' },
  };
  const groupB: Node = {
    id: 'groupB',
    type: 'groupNode',
    position: { x: 500, y: 0 },
    style: { width: 300, height: 300 },
    data: { label: 'Zone B' },
  };
  // Child of groupA but positioned (relative to groupA) to overlap groupB.
  const stray = node('stray', 560, 100, 'Stray', {
    parentNode: 'groupA',
    extent: 'parent',
    data: { label: 'Stray', iconPath: '/s.svg' },
    style: { width: 120, height: 80 },
  });
  const report = analyzeDiagramQuality([groupA, groupB, stray], []);
  const overlaps = report.findings.filter(item => item.category === 'overlap');
  assert.equal(
    overlaps.some(item => item.id.startsWith('overlap:node-group:')),
    true,
    'expected a service-vs-unrelated-group overlap finding',
  );
});

test('quality analysis flags overlapping non-sibling groups', () => {
  // Two groups nested under different (missing) parents so the base same-parent
  // loop skips them; the extended detector must still catch the overlap.
  const groupA: Node = {
    id: 'groupA',
    type: 'groupNode',
    position: { x: 0, y: 0 },
    parentNode: 'pA',
    style: { width: 300, height: 300 },
    data: { label: 'Group A' },
  };
  const groupB: Node = {
    id: 'groupB',
    type: 'groupNode',
    position: { x: 100, y: 100 },
    parentNode: 'pB',
    style: { width: 300, height: 300 },
    data: { label: 'Group B' },
  };
  const report = analyzeDiagramQuality([groupA, groupB], []);
  assert.equal(
    report.findings.some(item => item.category === 'overlap' && item.id.startsWith('overlap:groups:')),
    true,
  );
});

test('quality analysis flags dangling, self-loop, and duplicate edges', () => {
  const a = node('a', 0, 0, 'Azure Functions', { data: { label: 'Azure Functions', iconPath: '/a.svg' } });
  const b = node('b', 400, 0, 'Service Bus', { data: { label: 'Service Bus', iconPath: '/b.svg' } });
  const edges: Edge[] = [
    { id: 'dangling', source: 'a', target: 'ghost', type: 'editableEdge', label: 'to nowhere' },
    { id: 'loop', source: 'a', target: 'a', type: 'editableEdge', label: 'self' },
    { id: 'dup1', source: 'a', target: 'b', type: 'editableEdge', label: 'first' },
    { id: 'dup2', source: 'a', target: 'b', type: 'editableEdge', label: 'second' },
  ];
  const report = analyzeDiagramQuality([a, b], edges);
  const categories = new Set(report.findings.map(finding => finding.category));
  assert.equal(categories.has('dangling-edge'), true);
  assert.equal(categories.has('self-loop'), true);
  assert.equal(categories.has('duplicate-edge'), true);
  const dup = report.findings.find(finding => finding.category === 'duplicate-edge');
  assert.deepEqual(dup?.edgeIds, ['dup2']);
});

test('quality analysis flags under-specified and over-crowded diagrams', () => {
  const sparse = [
    node('s1', 0, 0, 'Azure Functions', { data: { label: 'Azure Functions', iconPath: '/a.svg' } }),
    node('s2', 400, 0, 'Service Bus', { data: { label: 'Service Bus', iconPath: '/b.svg' } }),
  ];
  const sparseReport = analyzeDiagramQuality(sparse, [edge('e', 's1', 's2')]);
  assert.equal(
    sparseReport.findings.some(finding => finding.category === 'under-specified'),
    true,
  );

  const crowded = Array.from({ length: 26 }, (_, index) => (
    node(`c${index}`, (index % 6) * 300, Math.floor(index / 6) * 240, `Svc ${index}`, {
      data: { label: `Svc ${index}`, iconPath: '/x.svg' },
    })
  ));
  const crowdedReport = analyzeDiagramQuality(crowded, []);
  assert.equal(
    crowdedReport.findings.some(finding => finding.category === 'over-crowded'),
    true,
  );
});

test('quality analysis flags empty groups', () => {
  const emptyGroup: Node = {
    id: 'emptyGroup',
    type: 'groupNode',
    position: { x: 0, y: 0 },
    style: { width: 200, height: 160 },
    data: { label: 'Empty Zone' },
  };
  const report = analyzeDiagramQuality([emptyGroup], []);
  const finding = report.findings.find(item => item.category === 'empty-group');
  assert.ok(finding, 'expected an empty-group finding');
  assert.deepEqual(finding?.nodeIds, ['emptyGroup']);
});
