import assert from 'node:assert/strict';
import test from 'node:test';
import { preserveConnectionRoutes } from '../dist/connectionRouting.js';

import { computeLayout, reflowLayoutForPresentation } from '../dist/layoutEngine.js';

const groups = [
  { id: 'ingress', label: 'Ingress' },
  { id: 'workers', label: 'Workers' },
  { id: 'data', label: 'Data' },
];

const services = [
  { name: 'Gateway', type: 'API Management', groupId: 'ingress' },
  { name: 'Worker A', type: 'Function App', groupId: 'workers' },
  { name: 'Worker B', type: 'Function App', groupId: 'workers' },
  { name: 'Worker C', type: 'Function App', groupId: 'workers' },
  { name: 'Sink', type: 'SQL Database', groupId: 'data' },
];

const connections = [
  { from: 'Gateway', to: 'Worker A' },
  { from: 'Gateway', to: 'Worker B' },
  { from: 'Gateway', to: 'Worker C' },
  { from: 'Worker A', to: 'Sink' },
  { from: 'Worker B', to: 'Sink' },
  { from: 'Worker C', to: 'Sink' },
];

function approximatelyEqual(actual, expected) {
  assert.ok(Math.abs(actual - expected) < 0.001, `${actual} should equal ${expected}`);
}

function verifyDistributedPorts(direction) {
  const layout = computeLayout(services, connections, groups, direction);
  const nodes = new Map(layout.nodes.map((node) => [node.name, node]));
  const gateway = nodes.get('Gateway');
  const sink = nodes.get('Sink');
  const gatewayEdges = layout.edges.filter((edge) => edge.from === 'Gateway');
  const sinkEdges = layout.edges.filter((edge) => edge.to === 'Sink');
  const perpendicular = direction === 'TB' ? 'x' : 'y';

  assert.equal(gatewayEdges.length, 3);
  assert.equal(sinkEdges.length, 3);
  assert.equal(new Set(gatewayEdges.map((edge) => edge.points[0][perpendicular])).size, 3);
  assert.equal(new Set(sinkEdges.map((edge) => edge.points.at(-1)[perpendicular])).size, 3);

  const orderedGatewayEdges = [...gatewayEdges].sort((left, right) => {
    const leftTarget = nodes.get(left.to);
    const rightTarget = nodes.get(right.to);
    return direction === 'TB'
      ? leftTarget.x - rightTarget.x
      : leftTarget.y - rightTarget.y;
  });
  const orderedSinkEdges = [...sinkEdges].sort((left, right) => {
    const leftSource = nodes.get(left.from);
    const rightSource = nodes.get(right.from);
    return direction === 'TB'
      ? leftSource.x - rightSource.x
      : leftSource.y - rightSource.y;
  });

  assert.deepEqual(
    orderedGatewayEdges.map((edge) => edge.points[0][perpendicular]),
    [...orderedGatewayEdges.map((edge) => edge.points[0][perpendicular])].sort((a, b) => a - b),
  );
  assert.deepEqual(
    orderedSinkEdges.map((edge) => edge.points.at(-1)[perpendicular]),
    [...orderedSinkEdges.map((edge) => edge.points.at(-1)[perpendicular])].sort((a, b) => a - b),
  );

  for (const edge of gatewayEdges) {
    const target = nodes.get(edge.to);
    const targetPoint = edge.points.at(-1);
    if (direction === 'TB') {
      approximatelyEqual(edge.points[0].y, gateway.y + gateway.height);
      approximatelyEqual(targetPoint.y, target.y);
      approximatelyEqual(targetPoint.x, target.x + target.width / 2);
    } else {
      approximatelyEqual(edge.points[0].x, gateway.x + gateway.width);
      approximatelyEqual(targetPoint.x, target.x);
      approximatelyEqual(targetPoint.y, target.y + target.height / 2);
    }
  }

  for (const edge of sinkEdges) {
    const source = nodes.get(edge.from);
    const sourcePoint = edge.points[0];
    const targetPoint = edge.points.at(-1);
    if (direction === 'TB') {
      approximatelyEqual(sourcePoint.x, source.x + source.width / 2);
      approximatelyEqual(sourcePoint.y, source.y + source.height);
      approximatelyEqual(targetPoint.y, sink.y);
    } else {
      approximatelyEqual(sourcePoint.x, source.x + source.width);
      approximatelyEqual(sourcePoint.y, source.y + source.height / 2);
      approximatelyEqual(targetPoint.x, sink.x);
    }
  }
}

test('grouped TB layout distributes shared edge ports in endpoint order', () => {
  verifyDistributedPorts('TB');
});

test('grouped LR layout distributes shared edge ports in endpoint order', () => {
  verifyDistributedPorts('LR');
});

test('layout rejects architectures above the defensive service limit', () => {
  const oversized = Array.from({ length: 251 }, (_, index) => ({
    name: `Service ${index}`,
    type: 'App Service',
  }));
  assert.throws(
    () => computeLayout(oversized, [], [], 'TB'),
    /at most 250 services/,
  );
});

test('multi-region presentation keeps overflow data nodes clear of fixed slots', () => {
  const regionalGroups = [
    { id: 'global', label: 'Global Edge' },
    { id: 'primary', label: 'Primary Region' },
    { id: 'secondary', label: 'Secondary Region' },
  ];
  const regionalServices = [
    { name: 'Global Front Door', type: 'Azure Front Door', groupId: 'global' },
    { name: 'Global WAF', type: 'Web Application Firewall', groupId: 'global' },
    { name: 'Global API Management', type: 'API Management', groupId: 'global' },
    ...['primary', 'secondary'].flatMap((groupId) => [
      { name: `${groupId} SQL`, type: 'SQL Database', groupId },
      { name: `${groupId} Redis`, type: 'Redis Cache', groupId },
      { name: `${groupId} Storage`, type: 'Storage Account', groupId },
      { name: `${groupId} Key Vault`, type: 'Key Vault', groupId },
    ]),
  ];
  const layout = reflowLayoutForPresentation(
    computeLayout(regionalServices, [], regionalGroups, 'LR'),
  );

  for (let left = 0; left < layout.nodes.length; left += 1) {
    for (let right = left + 1; right < layout.nodes.length; right += 1) {
      const a = layout.nodes[left];
      const b = layout.nodes[right];
      assert.equal(
        !(
          a.x + a.width <= b.x
          || b.x + b.width <= a.x
          || a.y + a.height <= b.y
          || b.y + b.height <= a.y
        ),
        false,
        `${a.name} overlaps ${b.name}`,
      );
    }
  }
});

function assertConnectionGeometry(layout, expected) {
  assert.deepEqual(layout.edges.map(({ from, to, label, type }) => ({ from, to, label, type })), expected);
  assert.equal(new Set(layout.edges.map(edge => edge.key)).size, expected.length);
  for (const edge of layout.edges) {
    assert.ok(edge.points.length >= 2);
    for (const point of edge.points) {
      assert.ok(Number.isFinite(point.x) && Number.isFinite(point.y));
      assert.ok(point.x >= 0 && point.x <= layout.width && point.y >= 0 && point.y <= layout.height,
        `${edge.label} is inside the computed canvas`);
    }
    if (edge.from === edge.to) {
      const node = layout.nodes.find(value => value.name === edge.from);
      assert.ok(edge.points.length >= 4);
      assert.ok(edge.points.some(point => point.x < node.x || point.x > node.x + node.width
        || point.y < node.y || point.y > node.y + node.height), 'the loop leaves its own tile');
    }
  }
  const routes = layout.edges.map(edge => JSON.stringify(edge.points));
  assert.equal(new Set(routes).size, routes.length, 'parallel and self-loop occurrences retain independent routes');
}

test('existing self-loop polylines survive until their nodes are reanchored', () => {
  const node = { name: 'Worker', type: 'Function App', category: 'compute', x: 40, y: 40, width: 70, height: 70 };
  const edge = {
    key: 'connection-0', from: 'Worker', to: 'Worker', label: 'Retry', type: 'async',
    points: [{ x: 110, y: 60 }, { x: 180, y: 20 }, { x: 190, y: 120 }, { x: 110, y: 90 }],
  };
  assert.deepEqual(preserveConnectionRoutes([node], [edge])[0].points, edge.points);
  const moved = preserveConnectionRoutes([{ ...node, x: 400 }], [edge], true)[0];
  assert.ok(moved.points.every(point => point.x >= 470));
});

test('connection multiset and excursions survive flat/grouped TB/LR layouts', () => {
  const connections = [
    { from: 'Worker', to: 'Queue', label: 'Publish', type: 'async' },
    { from: 'Worker', to: 'Queue', label: 'Health check', type: 'sync' },
    { from: 'Queue', to: 'Worker', label: 'Deliver', type: 'async' },
    { from: 'Worker', to: 'Worker', label: 'Retry', type: 'async' },
    { from: 'Worker', to: 'Worker', label: 'Recover', type: 'optional' },
  ];
  for (const direction of ['TB', 'LR']) {
    for (const grouped of [false, true]) {
      const services = [
        { name: 'Worker', type: 'Function App', ...(grouped ? { groupId: 'app' } : {}) },
        { name: 'Queue', type: 'Service Bus', ...(grouped ? { groupId: 'data' } : {}) },
      ];
      const groups = grouped ? [{ id: 'app', label: 'Application' }, { id: 'data', label: 'Data' }] : [];
      const snapshot = structuredClone({ services, connections, groups });
      const layout = computeLayout(services, connections, groups, direction);
      assertConnectionGeometry(layout, connections);
      assert.deepEqual({ services, connections, groups }, snapshot);
      if (grouped) {
        const group = layout.groups.find(value => value.id === 'app');
        for (const edge of layout.edges.filter(value => value.from === value.to)) {
          assert.ok(edge.points.every(point => point.x <= group.x + group.width), 'group layout reserves loop space');
        }
      }
    }
  }
});

test('presentation reanchoring keeps parallel connections and loops inside its recomputed bounds', () => {
  for (const groups of [
    ['Edge', 'Application', 'Data', 'Monitoring'].map((label, index) => ({ id: `group-${index}`, label })),
    [{ id: 'global', label: 'Global Edge' }, { id: 'primary', label: 'Primary Region' }, { id: 'secondary', label: 'Secondary Region' }],
  ]) {
    const services = groups.map((group, index) => ({ name: `Service ${index}`, type: 'Function App', groupId: group.id }));
    const connections = [
      { from: 'Service 0', to: 'Service 1', label: 'Invoke', type: 'sync' },
      { from: 'Service 0', to: 'Service 1', label: 'Notify', type: 'async' },
      { from: 'Service 1', to: 'Service 1', label: 'Retry', type: 'async' },
      { from: 'Service 1', to: 'Service 1', label: 'Recover', type: 'optional' },
      ...services.slice(2).map((service, index) => ({
        from: `Service ${index + 1}`, to: service.name, label: `Next ${index}`, type: 'sync',
      })),
    ];
    const layout = computeLayout(services, connections, groups, 'LR');
    const snapshot = structuredClone(layout);
    const presented = reflowLayoutForPresentation(layout);
    assert.notDeepEqual(presented.nodes.map(node => [node.x, node.y]), layout.nodes.map(node => [node.x, node.y]));
    assertConnectionGeometry(presented, connections);
    assert.deepEqual(layout, snapshot);
  }
});
