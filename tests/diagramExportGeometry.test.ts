import test from 'node:test';
import assert from 'node:assert/strict';
import type { Edge, Node } from 'reactflow';
import {
  buildExportRoutes,
  collectExportBoxes,
  computeBounds,
  computeFitTransform,
  partitionBoxes,
  routeOrthogonal,
  type ExportBox,
} from '../src/services/diagramExportGeometry.ts';

function box(id: string, x: number, y: number, w = 150, h = 75): ExportBox {
  return { id, kind: 'service', label: id, category: 'other', x, y, w, h };
}

function node(id: string, data: Record<string, unknown>): Node {
  return { id, type: 'azureNode', position: { x: 0, y: 0 }, data } as unknown as Node;
}

test('child nodes are flattened into absolute page coordinates', () => {
  const nodes = [
    { id: 'zone', type: 'groupNode', position: { x: 100, y: 50 }, style: { width: 500, height: 400 }, data: { label: 'Zone' } },
    { id: 'child', type: 'azureNode', position: { x: 20, y: 30 }, parentNode: 'zone', width: 150, height: 75, data: { label: 'Child' } },
    { id: 'loose', type: 'azureNode', position: { x: 700, y: 10 }, data: {} },
  ] as unknown as Node[];

  const boxes = collectExportBoxes(nodes);
  assert.deepEqual(
    { x: boxes.get('child')!.x, y: boxes.get('child')!.y },
    { x: 120, y: 80 },
    'the parent offset must be added',
  );
  assert.equal(boxes.get('zone')!.w, 500, 'group size comes from style');
  assert.equal(boxes.get('loose')!.w, 150, 'services fall back to the default tile size');
  assert.equal(boxes.get('loose')!.label, 'Service', 'missing labels get a readable default');

  const { groups, services } = partitionBoxes(boxes);
  assert.deepEqual(groups.map((entry) => entry.id), ['zone']);
  assert.deepEqual(services.map((entry) => entry.id), ['child', 'loose']);
});

test('nested groups are flattened through the whole ancestor chain', () => {
  const nodes = [
    { id: 'outer', type: 'groupNode', position: { x: 1000, y: 500 }, style: { width: 800, height: 600 }, data: { label: 'Subscription' } },
    { id: 'inner', type: 'groupNode', position: { x: 40, y: 60 }, parentNode: 'outer', style: { width: 400, height: 300 }, data: { label: 'VNet' } },
    { id: 'svc', type: 'azureNode', position: { x: 20, y: 30 }, parentNode: 'inner', width: 150, height: 75, data: { label: 'VM' } },
  ] as unknown as Node[];

  const boxes = collectExportBoxes(nodes);
  assert.deepEqual({ x: boxes.get('outer')!.x, y: boxes.get('outer')!.y }, { x: 1000, y: 500 });
  assert.deepEqual({ x: boxes.get('inner')!.x, y: boxes.get('inner')!.y }, { x: 1040, y: 560 });
  assert.deepEqual({ x: boxes.get('svc')!.x, y: boxes.get('svc')!.y }, { x: 1060, y: 590 });

  // Nothing may escape the outer zone, otherwise the fit transform shrinks the
  // whole drawing to cover phantom space.
  const bounds = computeBounds(boxes.values());
  assert.deepEqual(bounds, { minX: 1000, minY: 500, maxX: 1800, maxY: 1100 });
});

test('a cyclic parent chain degrades gracefully instead of hanging', () => {
  const nodes = [
    { id: 'a', type: 'groupNode', position: { x: 10, y: 20 }, parentNode: 'b', style: { width: 100, height: 100 }, data: {} },
    { id: 'b', type: 'groupNode', position: { x: 30, y: 40 }, parentNode: 'a', style: { width: 100, height: 100 }, data: {} },
  ] as unknown as Node[];

  const boxes = collectExportBoxes(nodes);
  assert.equal(boxes.size, 2);
  for (const box of boxes.values()) {
    assert.equal(Number.isFinite(box.x) && Number.isFinite(box.y), true);
  }
});

test('bounds cover every box and stay valid when empty', () => {
  const bounds = computeBounds([box('a', 0, 0), box('b', 400, 200)]);
  assert.deepEqual(bounds, { minX: 0, minY: 0, maxX: 550, maxY: 275 });
  assert.deepEqual(computeBounds([]), { minX: 0, minY: 0, maxX: 1, maxY: 1 });
});

test('the fit transform centres the drawing and never magnifies it', () => {
  const frame = { x: 0.5, y: 1, w: 12, h: 5 };

  const wide = computeFitTransform({ minX: 0, minY: 0, maxX: 2400, maxY: 1000 }, frame);
  assert.ok(wide.scale <= 1 / 96 + 1e-9, 'never larger than 96 px per inch');
  assert.ok(2400 * wide.scale <= frame.w + 1e-9, 'content fits horizontally');
  assert.ok(1000 * wide.scale <= frame.h + 1e-9, 'content fits vertically');
  assert.ok(Math.abs((frame.x + frame.w / 2) - (wide.offsetX + 1200 * wide.scale)) < 1e-6, 'centred');

  // A small diagram is capped at 1:1 rather than being blown up.
  const small = computeFitTransform({ minX: 0, minY: 0, maxX: 300, maxY: 150 }, frame);
  assert.ok(Math.abs(small.scale - 1 / 96) < 1e-9);

  // Negative canvas coordinates are shifted into the frame.
  const shifted = computeFitTransform({ minX: -500, minY: -250, maxX: -100, maxY: -50 }, frame);
  assert.ok(shifted.offsetX + -500 * shifted.scale >= frame.x - 1e-9);
});

test('orthogonal routes leave from facing edges and bend on the mid line', () => {
  const horizontal = routeOrthogonal(box('a', 0, 0), box('b', 600, 0));
  assert.equal(horizontal.points.length, 2, 'aligned tiles need no elbow');
  assert.deepEqual(horizontal.points[0], { x: 150, y: 37.5 }, 'exits the right edge');
  assert.deepEqual(horizontal.points[1], { x: 600, y: 37.5 }, 'enters the left edge');

  const elbow = routeOrthogonal(box('a', 0, 0), box('b', 600, 300));
  assert.equal(elbow.points.length, 4);
  assert.equal(elbow.points[0].x, 150);
  assert.equal(elbow.points[1].x, elbow.points[2].x, 'the elbow is a vertical segment');
  assert.equal(elbow.labelAnchor.x, elbow.points[1].x);

  const backwards = routeOrthogonal(box('a', 600, 0), box('b', 0, 0));
  assert.equal(backwards.points[0].x, 600, 'exits the left edge when the target is behind');

  const vertical = routeOrthogonal(box('a', 0, 0), box('b', 0, 400));
  assert.equal(vertical.points.length, 2);
  assert.equal(vertical.points[0].y, 75, 'exits the bottom edge');
});

test('routes carry labels and dash state, and skip dangling edges', () => {
  const boxes = new Map([
    ['a', box('a', 0, 0)],
    ['b', box('b', 600, 0)],
  ]);
  const edges = [
    { id: 'labelled', source: 'a', target: 'b', data: { label: 'HTTPS 443' } },
    { id: 'fallback', source: 'a', target: 'b', label: 'legacy label' },
    { id: 'dashed', source: 'a', target: 'b', animated: true },
    { id: 'styled', source: 'a', target: 'b', style: { strokeDasharray: '4 4' } },
    { id: 'dangling', source: 'a', target: 'missing' },
    { id: 'self', source: 'a', target: 'a' },
  ] as unknown as Edge[];

  const routes = buildExportRoutes(edges, boxes);
  assert.deepEqual(routes.map((route) => route.id), ['labelled', 'fallback', 'dashed', 'styled']);
  assert.equal(routes[0].label, 'HTTPS 443', 'data.label wins over the React Flow label');
  assert.equal(routes[1].label, 'legacy label');
  assert.equal(routes[0].dashed, false);
  assert.equal(routes[2].dashed, true, 'animated edges export dashed');
  assert.equal(routes[3].dashed, true, 'strokeDasharray exports dashed');
});

test('the node category drives export colours, with the icon folder as fallback', () => {
  const boxes = collectExportBoxes([
    // AI-generated services keep `category` but can have an empty iconPath when
    // no icon file matched — they must not fall back to the unstyled "other".
    node('ai', { category: 'Databases', iconPath: '' }),
    node('icon-only', { iconPath: '/Azure_Public_Service_Icons/Icons/networking/00001-icon-service-Virtual-Networks.svg' }),
    node('explicit-wins', {
      category: 'compute',
      iconPath: '/Azure_Public_Service_Icons/Icons/storage/00086-icon-service-Storage-Accounts.svg',
    }),
    node('bare', {}),
  ]);

  assert.equal(boxes.get('ai')?.category, 'databases', 'category is normalised to the palette key');
  assert.equal(boxes.get('ai')?.iconPath, undefined, 'an empty icon path is not treated as a path');
  assert.equal(boxes.get('icon-only')?.category, 'networking');
  assert.equal(boxes.get('explicit-wins')?.category, 'compute');
  assert.equal(boxes.get('bare')?.category, 'other');
});

test('the canonical service name is carried into exports separately from the label', () => {
  const boxes = collectExportBoxes([
    node('renamed', { label: 'Orders DB', serviceName: 'SQL Database' }),
    node('unnamed', { label: 'Orders DB' }),
    node('blank', { label: 'Orders DB', serviceName: '   ' }),
  ]);

  assert.equal(boxes.get('renamed')?.serviceName, 'SQL Database');
  assert.equal(boxes.get('renamed')?.label, 'Orders DB', 'the user label is never overwritten');
  assert.equal(boxes.get('unnamed')?.serviceName, undefined);
  assert.equal(boxes.get('blank')?.serviceName, undefined, 'whitespace is not a service name');
});
