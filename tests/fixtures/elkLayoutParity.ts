import type { Edge, Node } from 'reactflow';
import type { ApplyLayoutOptions } from '../../src/utils/layoutPresets';

export interface ElkLayoutFixture {
  name: string;
  nodes: Node[];
  edges: Edge[];
  options: ApplyLayoutOptions;
}

const options: ApplyLayoutOptions = {
  preset: 'flow-lr',
  spacing: 'comfortable',
  edgeStyle: 'orthogonal',
  emphasizePrimaryPath: false,
  layoutEngine: 'elk',
};

function service(id: string, index: number, parentNode?: string): Node {
  return {
    id, type: 'azureNode',
    position: { x: (index % 5) * 300, y: Math.floor(index / 5) * 180 },
    width: 180, height: 100,
    ...(parentNode ? { parentNode } : {}),
    data: { label: `Service ${index}`, serviceName: 'App Services', category: 'Compute' },
  };
}

function connection(id: string, source: string, target: string): Edge {
  return {
    id, source, target, sourceHandle: 'right', targetHandle: 'left',
    type: 'editableEdge', label: `Request ${id}`,
    data: { direction: 'forward', pathStyle: 'orthogonal' },
  };
}

function graph(count: number, grouped: boolean): ElkLayoutFixture {
  const groups: Node[] = grouped ? Array.from({ length: Math.ceil(count / 10) }, (_, index) => ({
    id: `group-${index}`, type: 'groupNode', position: { x: index * 700, y: 0 },
    style: { width: 650, height: 800 }, data: { label: `Zone ${index}` },
  })) : [];
  const services = Array.from({ length: count }, (_, index) => service(
    `service-${index}`,
    grouped ? index % 10 : index,
    grouped ? `group-${Math.floor(index / 10)}` : undefined,
  ));
  const edges: Edge[] = [];
  for (let index = 1; index < count; index += 1) {
    edges.push(connection(`edge-${index}`, `service-${index - 1}`, `service-${index}`));
    if (index % 5 === 0) {
      edges.push(connection(`cross-${index}`, `service-${index - 4}`, `service-${index}`));
    }
  }
  return {
    name: `${count}-${grouped ? 'grouped' : 'ungrouped'}`,
    nodes: [...groups, ...services], edges, options: { ...options },
  };
}

export function elkLayoutFixtures(): ElkLayoutFixture[] {
  const nested: ElkLayoutFixture = {
    name: 'nested-hierarchy',
    nodes: [
      { id: 'outer', type: 'groupNode', position: { x: 100, y: 100 }, style: { width: 650, height: 450 }, data: { label: 'Outer' } },
      { id: 'inner', type: 'groupNode', parentNode: 'outer', position: { x: -40, y: 80 }, style: { width: 400, height: 250 }, data: { label: 'Inner' } },
      service('nested-a', 0, 'inner'),
      service('nested-b', 1, 'inner'),
      service('direct-child', 4, 'outer'),
      service('outside', 6),
    ],
    edges: [
      connection('nested-link', 'nested-a', 'nested-b'),
      connection('cross-boundary', 'nested-b', 'outside'),
      connection('direct-link', 'outside', 'direct-child'),
    ],
    options: { ...options },
  };
  const cyclic: ElkLayoutFixture = {
    name: 'cycles-parallel-and-duplicate-services',
    nodes: Array.from({ length: 6 }, (_, index) => ({
      ...service(`cycle-${index}`, index),
      data: { label: 'Duplicate service', serviceName: 'App Services' },
    })),
    edges: [
      connection('cycle-a', 'cycle-0', 'cycle-1'),
      connection('cycle-b', 'cycle-1', 'cycle-2'),
      connection('cycle-c', 'cycle-2', 'cycle-0'),
      { ...connection('parallel', 'cycle-0', 'cycle-1'), data: { direction: 'reverse' } },
      { ...connection('both', 'cycle-2', 'cycle-3'), data: { direction: 'bidirectional' } },
      connection('self', 'cycle-4', 'cycle-4'),
    ],
    options: { ...options, preset: 'flow-tb', emphasizePrimaryPath: true },
  };
  const manual: ElkLayoutFixture = {
    name: 'manual-dimensions-handles-and-labels',
    nodes: [
      { ...service('manual-a', 0), selected: true, width: 260, height: 140, data: { label: 'Wide service', serviceName: 'App Services', pricing: { customPrice: 42 } } },
      { ...service('manual-b', 1), width: 220, height: 160 },
      { ...service('manual-c', 2), width: 190, height: 120 },
    ],
    edges: [
      { ...connection('manual-edge', 'manual-a', 'manual-b'), sourceHandle: 'bottom', targetHandle: 'top', data: { direction: 'forward', labelOffsetAuto: false, labelOffsetX: 45, labelOffsetY: -25 } },
      connection('horizontal', 'manual-b', 'manual-c'),
    ],
    options: { ...options, spacing: 'compact' },
  };
  return [
    graph(20, false), graph(80, false), graph(250, false),
    graph(20, true), graph(80, true), graph(250, true),
    nested, cyclic, manual,
    { ...graph(8, false), name: 'disconnected', edges: [] },
  ];
}
