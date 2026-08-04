import assert from 'node:assert/strict';
import test from 'node:test';

import { computeLayout } from '../dist/layoutEngine.js';

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
