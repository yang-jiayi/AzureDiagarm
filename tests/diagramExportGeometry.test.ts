import test from 'node:test';
import assert from 'node:assert/strict';
import type { Edge, Node } from 'reactflow';
import {
  buildExportRoutes,
  collectExportBoxes,
  computeBounds,
  compactEmptyGutters,
  computeContentBounds,
  computeFitTransform,
  partitionBoxes,
  routeOrthogonal,
  narrateEdgeCallouts,
  workflowListFromEdges,
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
  // Dangling edges are dropped; self-loops are now kept as a visible stub.
  assert.deepEqual(routes.map((route) => route.id), ['labelled', 'fallback', 'dashed', 'styled', 'self']);
  assert.equal(routes[0].label, 'HTTPS 443', 'data.label wins over the React Flow label');
  assert.equal(routes[1].label, 'legacy label');
  assert.equal(routes[0].dashed, false);
  assert.equal(routes[2].dashed, true, 'animated edges export dashed');
  assert.equal(routes[3].dashed, true, 'strokeDasharray exports dashed');

  const self = routes.find((route) => route.id === 'self')!;
  assert.equal(self.isSelfLoop, true, 'a self-referencing edge is a self-loop');
  assert.ok(self.points.length >= 3, 'the self-loop is a visible stub, not a zero-length line');
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

/**
 * Every exporter draws a single arrowhead at the route's target end, so the
 * route — not the stored tuple — decides which way the arrow points. An edge
 * with direction 'reverse' keeps its tuple and moves only the arrowhead, so
 * following the tuple pointed PPTX, VSDX, Draw.io and HTML the opposite way
 * to the canvas.
 */
test('export routes follow the drawn arrow, not the stored endpoint tuple', () => {
  const boxes = new Map([
    ['a', box('a', 0, 0)],
    ['b', box('b', 400, 0)],
  ]);
  const directed = (direction?: string): Edge => ({
    id: 'e1',
    source: 'a',
    target: 'b',
    ...(direction ? { data: { direction } } : {}),
  } as Edge);

  const forward = buildExportRoutes([directed()], boxes)[0];
  assert.equal(forward.sourceId, 'a');
  assert.equal(forward.targetId, 'b');
  assert.equal(forward.bidirectional, false);

  const reverse = buildExportRoutes([directed('reverse')], boxes)[0];
  assert.equal(reverse.sourceId, 'b', 'a reverse edge points from b');
  assert.equal(reverse.targetId, 'a', 'a reverse edge points to a');
  assert.equal(reverse.bidirectional, false);

  // The polyline has to run from the arrow tail too: exporters place the head
  // at the last point and Visio derives its whole local axis from begin->end.
  assert.ok(reverse.points[0].x > reverse.points[reverse.points.length - 1].x,
    'the reverse polyline must start at the right-hand box');
  assert.ok(forward.points[0].x < forward.points[forward.points.length - 1].x);

  const both = buildExportRoutes([directed('bidirectional')], boxes)[0];
  assert.equal(both.bidirectional, true, 'a two-way edge asks for a second arrowhead');
  assert.equal(both.sourceId, 'a');
  assert.equal(both.targetId, 'b');

  // An unknown or missing direction must never be treated as reverse.
  const unknown = buildExportRoutes([directed('sideways')], boxes)[0];
  assert.equal(unknown.sourceId, 'a');
  assert.equal(unknown.bidirectional, false);
});

test('a reverse edge still resolves its endpoints and self-loops are unaffected', () => {
  const boxes = new Map([['a', box('a', 0, 0)], ['b', box('b', 400, 0)]]);
  // Dangling endpoints must still be skipped after orientation.
  assert.equal(buildExportRoutes([{ id: 'e', source: 'a', target: 'gone', data: { direction: 'reverse' } } as Edge], boxes).length, 0);
  const loop = buildExportRoutes([{ id: 'e', source: 'a', target: 'a', data: { direction: 'reverse' } } as Edge], boxes)[0];
  assert.equal(loop.isSelfLoop, true);
  assert.equal(loop.sourceId, 'a');
});

/**
 * Workflow numbering has to survive the trip into the shared geometry, because
 * every exporter reads its badges from ExportRoute and nowhere else.
 */
test('export routes carry only well-formed workflow step numbers', () => {
  const boxes = new Map([['a', box('a', 0, 0)], ['b', box('b', 400, 0)]]);
  const routeFor = (data: Record<string, unknown>) =>
    buildExportRoutes([{ id: 'e1', source: 'a', target: 'b', data } as Edge], boxes)[0];

  assert.equal(routeFor({ stepNumber: 3 }).stepNumber, 3);
  // Persisted diagrams round-trip through JSON, so a numeric string counts.
  assert.equal(routeFor({ stepNumber: '2' }).stepNumber, 2);
  assert.equal(routeFor({}).stepNumber, undefined, 'an unnumbered edge gets no badge');
  for (const bad of [0, -1, 1.5, 'two', null, NaN, Infinity]) {
    assert.equal(routeFor({ stepNumber: bad }).stepNumber, undefined,
      `${String(bad)} is not a step number`);
  }
});

test('two arrows never wear the same number, and no sentence is lost to a duplicate', () => {
  // A re-prompted model happily numbers every hop of a flow "3", each with its
  // own sentence. The workflow list is keyed by number, so all but the first
  // sentence vanished while five badges on the drawing all read 3 - the reader
  // has five identical digits and one row to look them up in.
  const edges = ['a', 'b', 'c', 'd', 'e'].map((k, i) => ({
    id: `dup${i}`, source: `n${i}`, target: `n${i + 1}`, label: `hop ${k}`,
    data: { stepNumber: 3, stepDescription: `sentence ${k}` },
  })) as unknown as Edge[];

  const repaired = narrateEdgeCallouts(edges);
  const numbers = repaired.map((e) => (e.data as { stepNumber: number }).stepNumber);
  assert.equal(new Set(numbers).size, numbers.length, `duplicate callouts remain: ${numbers.join(', ')}`);
  // The author's own numbering is respected: the first keeps 3.
  assert.equal(numbers[0], 3);

  const rows = workflowListFromEdges(repaired);
  assert.deepEqual(
    rows.map((r) => r.description).sort(),
    ['sentence a', 'sentence b', 'sentence c', 'sentence d', 'sentence e'],
    'every authored sentence must reach the workflow list',
  );
});

test('a diagram whose numbering is already distinct is left exactly as it was', () => {
  const edges = [0, 1, 2].map((i) => ({
    id: `e${i}`, source: `n${i}`, target: `n${i + 1}`, label: `hop ${i}`,
    data: { stepNumber: i + 1, stepDescription: `sentence ${i}` },
  })) as unknown as Edge[];
  assert.equal(narrateEdgeCallouts(edges), edges, 'an untouched diagram must return the same array');
});

test('a legitimate second row survives the outlier trim that removes real strays', () => {
  // Ten services on one row, one on a second row, and two nodes flung far to
  // either side. The vertical quartile range of a single-row drawing is zero,
  // so a per-axis fence has zero width vertically and throws the second-row
  // service out with the strays — it is then parked into the margin strip,
  // away from the neighbours it is wired to. One fence width shared by both
  // axes, floored at the median tile extent, keeps it.
  const boxes = [
    ...Array.from({ length: 10 }, (_, i) => box(`row${i}`, i * 200, 0)),
    box('secondRow', 400, 260),
    box('farLeft', -6000, 0),
    box('farRight', 6000, 0),
  ];
  const trimmed = computeContentBounds(boxes);
  const inside = (b: ExportBox): boolean => b.x >= trimmed.minX && b.y >= trimmed.minY
    && b.x + b.w <= trimmed.maxX && b.y + b.h <= trimmed.maxY;

  assert.equal(inside(boxes.find((b) => b.id === 'secondRow')!), true, 'the second row is part of the drawing');
  assert.equal(inside(boxes.find((b) => b.id === 'row0')!), true, 'the first row is part of the drawing');
  assert.equal(inside(boxes.find((b) => b.id === 'farLeft')!), false, 'a node 6000px away is a stray');
  assert.equal(inside(boxes.find((b) => b.id === 'farRight')!), false, 'a node 6000px away is a stray');
});

test('outliers off several corners are peeled instead of setting their own fence', () => {
  // Four strays in a twelve-box drawing drag the upper quartile out far enough
  // that three of the four land inside a single-pass fence, and the "dense
  // cluster" comes back nearly as large as never trimming at all. Recomputing
  // the quartiles from the survivors each pass tightens the fence onto the
  // pack.
  const cluster = Array.from({ length: 8 }, (_, i) => box(`c${i}`, (i % 4) * 200, Math.floor(i / 4) * 180));
  const strays = [box('nw', -4000, -3000), box('ne', 4000, -3000), box('sw', -4000, 3000), box('se', 4000, 3000)];
  const trimmed = computeContentBounds([...cluster, ...strays]);

  assert.equal(trimmed.minX, 0, 'the trimmed bounds start at the cluster');
  assert.equal(trimmed.maxX, 750, 'the trimmed bounds end at the cluster');
  assert.ok(trimmed.maxY - trimmed.minY < 400, 'the trimmed drawing is the two-row cluster, not the corners');
});

test('an empty band wider than a screen is closed, and ordinary spacing is not', () => {
  // A DR region drawn 6000px east of the primary is a second region, not an
  // outlier to trim and not a stray to park: nothing in the drawing says to
  // move it, only the void between them is worth removing. Closing it turned a
  // 72in Visio sheet into 21in without changing what the drawing says.
  const map = new Map<string, ExportBox>();
  for (const b of [
    ...Array.from({ length: 12 }, (_, i) => box(`p${i}`, (i % 4) * 220, Math.floor(i / 4) * 180)),
    box('dr-a', 6020, 20),
    box('dr-b', 6720, 20),
  ]) map.set(b.id, b);

  const compact = compactEmptyGutters(map);
  const before = computeBounds(map.values());
  const after = computeBounds(compact.values());

  assert.ok(after.maxX - after.minX < 2200, 'the void between the two regions is closed');
  assert.ok(before.maxX - before.minX > 6800, 'the untouched drawing really was that wide');
  assert.equal(after.maxY - after.minY, before.maxY - before.minY, 'the axis with no void is untouched');
  assert.ok(
    compact.get('dr-a')!.x > compact.get('p3')!.x,
    'the DR region is still east of the primary — closing a void must not reorder the drawing',
  );
  assert.equal(
    compact.get('dr-b')!.x - compact.get('dr-a')!.x,
    700,
    'spacing inside a region is the author is, and is preserved exactly',
  );
});

test('a hub-and-spoke on the Architecture Center radius is left alone', () => {
  // The widest gap an author is own layout produces is the 1250px between a
  // spoke and the hub. Closing that would pull one arm in and leave the other
  // out, so the threshold sits deliberately above it.
  const R = 1400;
  const map = new Map<string, ExportBox>();
  for (const b of [
    box('hub', 0, 0), box('n', 0, -R), box('s', 0, R), box('e', R, 0), box('w', -R, 0),
  ]) map.set(b.id, b);

  const compact = compactEmptyGutters(map);

  assert.equal(compact, map, 'a radial layout is returned untouched, not rebuilt');
});
