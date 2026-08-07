import assert from 'node:assert/strict';
import test from 'node:test';
import type { Edge, Node } from 'reactflow';
import { relayoutDiagram as relayoutWithDagre } from '../src/utils/layoutEngine';
import {
  ELK_QUALITY_LAYOUT_OPTIONS,
  layoutArchitecture as layoutArchitectureWithElk,
  relayoutDiagram as relayoutWithElk,
} from '../src/utils/elkLayoutEngine';
import {
  captureGroupLayout,
  fitAllGroupsToContent,
  fitGroupToContent,
  restoreGroupLayout,
} from '../src/utils/groupUtils';
import {
  buildNestedHierarchyLayout,
  collectNestedHierarchyNodeIds,
  layoutNodeDimensions,
} from '../src/utils/layoutHierarchy';
import {
  applyLayoutPreset,
  straightenPrimaryPath,
} from '../src/utils/layoutPresets';
import { buildAbsolutePositionMap } from '../src/utils/preserveManualLayout';

function node(
  id: string,
  type: 'azureNode' | 'groupNode',
  x: number,
  y: number,
  options: Partial<Node> = {},
): Node {
  return {
    id,
    type,
    position: { x, y },
    data: { label: id },
    ...options,
  };
}

function edge(id: string, source: string, target: string): Edge {
  return { id, source, target, data: { direction: 'forward' } };
}

function overlaps(
  left: { x: number; y: number; width: number; height: number },
  right: { x: number; y: number; width: number; height: number },
): boolean {
  return left.x < right.x + right.width
    && left.x + left.width > right.x
    && left.y < right.y + right.height
    && left.y + left.height > right.y;
}

function boundsFor(
  nodes: Node[],
  nodeIds: string[],
): { x: number; y: number; width: number; height: number } {
  const absolute = buildAbsolutePositionMap(nodes);
  const bounds = nodeIds.map(id => {
    const candidate = nodes.find(node => node.id === id);
    const position = absolute.get(id);
    assert.ok(candidate && position);
    const dimensions = layoutNodeDimensions(candidate);
    return {
      x: position.x,
      y: position.y,
      width: dimensions.width,
      height: dimensions.height,
    };
  });
  const x = Math.min(...bounds.map(bound => bound.x));
  const y = Math.min(...bounds.map(bound => bound.y));
  const maxX = Math.max(...bounds.map(bound => bound.x + bound.width));
  const maxY = Math.max(...bounds.map(bound => bound.y + bound.height));
  return { x, y, width: maxX - x, height: maxY - y };
}

test('ELK quality options prioritize orthogonal routing and balanced placement', () => {
  assert.equal(ELK_QUALITY_LAYOUT_OPTIONS['elk.edgeRouting'], 'ORTHOGONAL');
  assert.equal(
    ELK_QUALITY_LAYOUT_OPTIONS['elk.layered.nodePlacement.strategy'],
    'NETWORK_SIMPLEX',
  );
  assert.equal(ELK_QUALITY_LAYOUT_OPTIONS['elk.spacing.edgeLabelNode'], '20');
});

test('fitGroupToContent preserves child absolute positions', () => {
  const nodes = [
    node('group', 'groupNode', 100, 100, { style: { width: 400, height: 300 } }),
    node('child', 'azureNode', 100, 100, {
      parentNode: 'group',
      width: 160,
      height: 100,
    }),
  ];
  const before = buildAbsolutePositionMap(nodes);
  const fitted = fitGroupToContent(nodes, 'group');

  assert.ok(fitted);
  const after = buildAbsolutePositionMap(fitted);
  assert.deepEqual(after.get('child'), before.get('child'));
  assert.deepEqual(fitted.find(candidate => candidate.id === 'group')?.position, {
    x: 160,
    y: 110,
  });
});

test('fitAllGroupsToContent fits nested groups from the inside out', () => {
  const nodes = [
    node('outer', 'groupNode', 100, 100, { style: { width: 600, height: 500 } }),
    node('inner', 'groupNode', 120, 130, {
      parentNode: 'outer',
      style: { width: 300, height: 200 },
    }),
    node('service', 'azureNode', 80, 90, {
      parentNode: 'inner',
      width: 160,
      height: 100,
    }),
  ];
  const before = buildAbsolutePositionMap(nodes);
  const fitted = fitAllGroupsToContent(nodes);
  const after = buildAbsolutePositionMap(fitted);

  assert.deepEqual(after.get('service'), before.get('service'));
  assert.equal(
    (fitted.find(candidate => candidate.id === 'inner')?.style as { width?: number }).width,
    240,
  );
});

test('group collapse snapshot restores nested positions and dimensions', () => {
  const nodes = [
    node('outer', 'groupNode', 100, 100, { style: { width: 600, height: 500 } }),
    node('inner', 'groupNode', 120, 130, {
      parentNode: 'outer',
      style: { width: 300, height: 200 },
    }),
    node('service', 'azureNode', 80, 90, {
      parentNode: 'inner',
      width: 160,
      height: 100,
    }),
  ];
  const snapshot = captureGroupLayout(nodes);
  const restored = restoreGroupLayout(fitAllGroupsToContent(nodes), snapshot);

  for (const original of nodes) {
    const result = restored.find(candidate => candidate.id === original.id);
    assert.deepEqual(result?.position, original.position);
  }
  for (const id of ['outer', 'inner']) {
    const original = nodes.find(candidate => candidate.id === id);
    const result = restored.find(candidate => candidate.id === id);
    assert.equal(
      (result?.style as { width?: number }).width,
      (original?.style as { width?: number }).width,
    );
    assert.equal(
      (result?.style as { height?: number }).height,
      (original?.style as { height?: number }).height,
    );
  }
});

test('nested hierarchy protection includes ancestors and all descendants', () => {
  const nodes = [
    node('outer', 'groupNode', 0, 0),
    node('inner', 'groupNode', 20, 20, { parentNode: 'outer' }),
    node('direct', 'azureNode', 40, 40, { parentNode: 'outer' }),
    node('nested', 'azureNode', 30, 30, { parentNode: 'inner' }),
    node('free', 'azureNode', 500, 100),
  ];

  assert.deepEqual(
    [...collectNestedHierarchyNodeIds(nodes)].sort(),
    ['direct', 'inner', 'nested', 'outer'],
  );
});

test('nested layout unit covers descendants outside the root group bounds', () => {
  const nodes = [
    node('outer', 'groupNode', 100, 100, { style: { width: 200, height: 200 } }),
    node('inner', 'groupNode', -80, -60, {
      parentNode: 'outer',
      style: { width: 500, height: 400 },
    }),
    node('nested', 'azureNode', 480, 350, {
      parentNode: 'inner',
      width: 180,
      height: 100,
    }),
  ];
  const layout = buildNestedHierarchyLayout(nodes);

  assert.equal(layout.units.length, 1);
  assert.deepEqual(layout.units[0], {
    rootId: 'outer',
    surrogateId: '__nested_hierarchy__outer',
    width: 660,
    height: 450,
    offsetX: -80,
    offsetY: -60,
  });
});

test('primary path alignment uses absolute coordinates for grouped nodes', () => {
  const nodes = [
    node('group', 'groupNode', 100, 200, { style: { width: 500, height: 400 } }),
    node('a', 'azureNode', 0, 300),
    node('b', 'azureNode', 50, 100, { parentNode: 'group' }),
    node('c', 'azureNode', 400, 500),
  ];
  const result = straightenPrimaryPath(
    nodes,
    [edge('ab', 'a', 'b'), edge('bc', 'b', 'c')],
    'LR',
  );
  const absolute = buildAbsolutePositionMap(result.nodes);

  assert.equal(absolute.get('a')?.y, 300);
  assert.equal(absolute.get('b')?.y, 300);
  assert.equal(absolute.get('c')?.y, 300);
});

test('Dagre relayout preserves nested group hierarchies', () => {
  const nodes = [
    node('outer', 'groupNode', 100, 100, { style: { width: 200, height: 150 } }),
    node('inner', 'groupNode', 50, 60, {
      parentNode: 'outer',
      style: { width: 400, height: 300 },
    }),
    node('nested', 'azureNode', 350, 260, {
      parentNode: 'inner',
      width: 180,
      height: 100,
    }),
    node('direct', 'azureNode', -100, -80, {
      parentNode: 'outer',
      width: 180,
      height: 100,
    }),
    node('free', 'azureNode', 900, 100, { width: 180, height: 100 }),
  ];
  const result = relayoutWithDagre(nodes, [edge('nested-free', 'nested', 'free')], {
    direction: 'LR',
  });

  for (const id of ['inner', 'nested', 'direct']) {
    const before = nodes.find(candidate => candidate.id === id);
    const after = result.find(candidate => candidate.id === id);
    assert.deepEqual(after?.position, before?.position);
    assert.equal(after?.parentNode, before?.parentNode);
  }
  const outer = result.find(candidate => candidate.id === 'outer');
  const free = result.find(candidate => candidate.id === 'free');
  assert.ok(outer && free);
  assert.equal(overlaps(
    boundsFor(result, ['outer', 'inner', 'nested', 'direct']),
    { ...free.position, width: 180, height: 100 },
  ), false);
});

test('ELK relayout preserves nested group hierarchies', async () => {
  const nodes = [
    node('outer', 'groupNode', 100, 100, { style: { width: 200, height: 150 } }),
    node('inner', 'groupNode', 50, 60, {
      parentNode: 'outer',
      style: { width: 400, height: 300 },
    }),
    node('nested', 'azureNode', 350, 260, {
      parentNode: 'inner',
      width: 180,
      height: 100,
    }),
    node('direct', 'azureNode', -100, -80, {
      parentNode: 'outer',
      width: 180,
      height: 100,
    }),
    node('free', 'azureNode', 900, 100, { width: 180, height: 100 }),
  ];
  const result = await relayoutWithElk(nodes, [edge('nested-free', 'nested', 'free')], {
    direction: 'LR',
  });

  for (const id of ['inner', 'nested', 'direct']) {
    const before = nodes.find(candidate => candidate.id === id);
    const after = result.find(candidate => candidate.id === id);
    assert.deepEqual(after?.position, before?.position);
    assert.equal(after?.parentNode, before?.parentNode);
  }
  const outer = result.find(candidate => candidate.id === 'outer');
  const free = result.find(candidate => candidate.id === 'free');
  assert.ok(outer && free);
  assert.equal(overlaps(
    boundsFor(result, ['outer', 'inner', 'nested', 'direct']),
    { ...free.position, width: 180, height: 100 },
  ), false);
});

test('ELK keeps colliding service and group IDs distinct', async () => {
  const result = await layoutArchitectureWithElk(
    [{ id: 'dup', name: 'Service' }],
    [],
    [{ id: 'dup', label: 'Group' }],
    { direction: 'LR' },
  );

  assert.equal(result.services.length, 1);
  assert.equal(result.groups.length, 1);
  assert.notDeepEqual(result.services[0].position, result.groups[0].position);
});

test('swimlanes move a nested hierarchy as one intact lane', async () => {
  const nodes = [
    node('outer', 'groupNode', 300, 300, { style: { width: 500, height: 400 } }),
    node('inner', 'groupNode', 50, 60, {
      parentNode: 'outer',
      style: { width: 250, height: 200 },
    }),
    node('nested', 'azureNode', 30, 40, { parentNode: 'inner' }),
  ];
  const result = await applyLayoutPreset(nodes, [], {
    preset: 'swimlanes',
    spacing: 'comfortable',
    edgeStyle: 'smooth',
    emphasizePrimaryPath: false,
  });

  assert.deepEqual(result.nodes.find(candidate => candidate.id === 'outer')?.position, {
    x: 80,
    y: 80,
  });
  assert.deepEqual(
    result.nodes.find(candidate => candidate.id === 'inner')?.position,
    { x: 50, y: 60 },
  );
  assert.deepEqual(
    result.nodes.find(candidate => candidate.id === 'nested')?.position,
    { x: 30, y: 40 },
  );
});

test('radial layout moves top-level groups as units and centers node bounds', async () => {
  const nodes = [
    node('group', 'groupNode', 100, 100, { style: { width: 400, height: 300 } }),
    node('grouped', 'azureNode', 50, 80, { parentNode: 'group' }),
    node('free', 'azureNode', 800, 100, { width: 180, height: 100 }),
  ];
  const result = await applyLayoutPreset(
    nodes,
    [edge('connected', 'grouped', 'free')],
    {
      preset: 'radial',
      spacing: 'comfortable',
      edgeStyle: 'smooth',
      emphasizePrimaryPath: false,
      selectedNodeId: 'grouped',
    },
  );

  assert.deepEqual(result.nodes.find(candidate => candidate.id === 'group')?.position, {
    x: 320,
    y: 210,
  });
  assert.deepEqual(
    result.nodes.find(candidate => candidate.id === 'grouped')?.position,
    { x: 50, y: 80 },
  );
  assert.equal(result.nodes.find(candidate => candidate.id === 'grouped')?.parentNode, 'group');
  assert.equal(result.nodes.find(candidate => candidate.id === 'free')?.parentNode, undefined);
});

test('radial layout sizes rings to avoid overlap between large groups', async () => {
  const groups = Array.from({ length: 4 }, (_, index) => (
    node(`group-${index}`, 'groupNode', index * 500, 300, {
      style: { width: 1000, height: 1000 },
    })
  ));
  const groupedServices = groups.map((group, index) => (
    node(`grouped-${index}`, 'azureNode', 50, 80, { parentNode: group.id })
  ));
  const center = node('center', 'azureNode', 0, 0);
  const result = await applyLayoutPreset(
    [center, ...groups, ...groupedServices],
    groupedServices.map((service, index) => edge(`edge-${index}`, 'center', service.id)),
    {
      preset: 'radial',
      spacing: 'comfortable',
      edgeStyle: 'smooth',
      emphasizePrimaryPath: false,
      selectedNodeId: 'center',
    },
  );
  const positionedGroups = groups.map(group => {
    const positioned = result.nodes.find(candidate => candidate.id === group.id);
    assert.ok(positioned);
    return { ...positioned.position, width: 1000, height: 1000 };
  });

  for (let left = 0; left < positionedGroups.length; left += 1) {
    for (let right = left + 1; right < positionedGroups.length; right += 1) {
      assert.equal(overlaps(positionedGroups[left], positionedGroups[right]), false);
    }
  }
});
