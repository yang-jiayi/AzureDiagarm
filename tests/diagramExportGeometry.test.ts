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
  fitBoxesWithin,
  boxScaleWithin,
  scaleBoxesWithin,
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
  // `?.iconPath === undefined` passes just as well when the BOX is missing, so
  // pin the box first: a dropped tile is exactly the regression being guarded.
  const ai = boxes.get('ai');
  assert.ok(ai, 'the tile was dropped from the export');
  assert.equal(ai.iconPath, undefined, 'an empty icon path is not treated as a path');
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
  // Both of these claim something about a tile that is present. Written as
  // `?.serviceName === undefined` they pass identically when the tile is gone,
  // which is the failure mode, not the success one.
  const unnamed = boxes.get('unnamed');
  const blank = boxes.get('blank');
  assert.ok(unnamed, 'the unnamed tile was dropped from the export');
  assert.ok(blank, 'the blank-serviceName tile was dropped from the export');
  assert.equal(unnamed.serviceName, undefined);
  assert.equal(blank.serviceName, undefined, 'whitespace is not a service name');
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

test('a symmetric hub-and-spoke stays symmetric when its voids are closed', () => {
  // A service tile is 150 wide and 75 tall, so at any given radius the gap
  // between two tiles is 75px larger horizontally than vertically. Judging a
  // void by that gap therefore answered the same question differently on the
  // two axes: at radius 1700 the north and south arms were closed and the east
  // and west arms were not, and a drawing the author made perfectly symmetric
  // came out with 4.1in arms one way and 17.7in the other.
  //
  // Measuring centre to centre removes the tile's own size from the question,
  // so both axes see the same 1700 and make the same decision.
  const R = 1700;
  const map = new Map<string, ExportBox>();
  for (const b of [
    box('hub', 0, 0), box('n', 0, -R), box('s', 0, R), box('e', R, 0), box('w', -R, 0),
  ]) map.set(b.id, b);

  const compact = compactEmptyGutters(map);
  const hub = compact.get('hub')!;
  // Edge to edge, which is what a reader sees as the length of an arm.
  const gap = {
    n: hub.y - (compact.get('n')!.y + 75),
    s: compact.get('s')!.y - (hub.y + 75),
    e: compact.get('e')!.x - (hub.x + 150),
    w: hub.x - (compact.get('w')!.x + 150),
  };

  assert.equal(gap.n, gap.s, 'the two vertical arms stay equal');
  assert.equal(gap.e, gap.w, 'the two horizontal arms stay equal');
  assert.equal(
    gap.n,
    gap.e,
    `a symmetric drawing stays symmetric: vertical arm ${gap.n}px, horizontal arm ${gap.e}px`,
  );
  assert.ok(gap.e < R - 150, 'and the voids really were closed, so this is not symmetric by doing nothing');
});

test('a box whose edge lands exactly on a void is not collapsed to a hairline', () => {
  // The lip of a void is by definition where the last box before it ends, so
  // there is always at least one box sitting on it. Shifting coordinates in a
  // step rather than a ramp moved that box's right edge without moving its
  // left, collapsing it to 1px: in an ordinary two-region drawing two of the
  // twelve services came out 0.008in wide and their labels with them.
  const map = new Map<string, ExportBox>();
  for (const b of [
    box('a', 0, 0), box('b', 200, 0), box('c', 400, 0),
    box('far-a', 6000, 0), box('far-b', 6200, 0),
  ]) map.set(b.id, b);

  const compact = compactEmptyGutters(map);

  for (const id of ['a', 'b', 'c', 'far-a', 'far-b']) {
    assert.equal(compact.get(id)!.w, 150, `${id} keeps its full width across the void`);
    assert.equal(compact.get(id)!.h, 75, `${id} keeps its full height across the void`);
  }
  assert.ok(
    compact.get('far-a')!.x - compact.get('c')!.x < 1000,
    'and the void really was closed',
  );
});

test('a zone drawn around a void shrinks to its contents instead of dragging the void along', () => {
  // A subscription frame, a tenant boundary, an "Azure" box — the commonest
  // annotation there is, and it spans every void in the drawing. Counting it
  // as content found no void at all and exported a sheet nine tenths blank;
  // moving only its origin would have left it 6000px wider than everything
  // inside it.
  const map = new Map<string, ExportBox>();
  map.set('azure', { id: 'azure', kind: 'group', category: '', label: 'Azure', x: -80, y: -80, w: 7060, h: 400 });
  for (const b of [
    box('e0', 0, 0), box('e1', 200, 0), box('e2', 400, 0),
    box('w0', 6000, 0), box('w1', 6200, 0), box('w2', 6400, 0),
  ]) map.set(b.id, b);

  const compact = compactEmptyGutters(map);
  const frame = compact.get('azure')!;
  const west = compact.get('w2')!;

  assert.ok(frame.w < 2200, `the frame closes with its contents, not around the void (${frame.w}px)`);
  assert.ok(
    frame.x <= compact.get('e0')!.x && frame.x + frame.w >= west.x + west.w,
    'and it still contains every service it contained before',
  );
});

test('closing voids answers the same way however many regions the drawing has', () => {
  // Every fixture above pins one shape of drawing: two regions, one hub, one
  // frame. A rule read on a single k cannot see a k-dependent answer, and the
  // void scan carries running state across the sweep - `reach` and `reachSize`
  // are updated inside the loop - so the second void is the first one that can
  // be computed from stale state, and no test had a second void.
  //
  // The gutter is measured from the two-region case rather than written down,
  // so this asserts what it means to say the answer is the same, without
  // copying a constant the module does not export.
  const PITCH = 220;
  const TILE = 150;
  const REGION_W = PITCH + TILE;
  const SEPARATION = 6000;

  const regions = (k: number): Map<string, ExportBox> => {
    const map = new Map<string, ExportBox>();
    for (let r = 0; r < k; r += 1) {
      for (let i = 0; i < 4; i += 1) {
        const b = box(`r${r}s${i}`, r * SEPARATION + (i % 2) * PITCH, Math.floor(i / 2) * 180);
        map.set(b.id, b);
      }
    }
    return map;
  };
  const leftOf = (m: Map<string, ExportBox>, r: number): number => Math.min(
    ...[...m.values()].filter((b) => b.id.startsWith(`r${r}s`)).map((b) => b.x),
  );
  const rightOf = (m: Map<string, ExportBox>, r: number): number => Math.max(
    ...[...m.values()].filter((b) => b.id.startsWith(`r${r}s`)).map((b) => b.x + b.w),
  );

  const two = compactEmptyGutters(regions(2));
  const gutter = leftOf(two, 1) - rightOf(two, 0);
  assert.ok(gutter > 0 && gutter < SEPARATION, `the two-region void really closed (${gutter}px)`);

  for (let k = 2; k <= 7; k += 1) {
    const before = regions(k);
    const compact = compactEmptyGutters(before);
    const tall = computeBounds(before.values());
    const flat = computeBounds(compact.values());

    for (let r = 1; r < k; r += 1) {
      assert.equal(
        leftOf(compact, r) - rightOf(compact, r - 1),
        gutter,
        `k=${k}: every void closes to the same gutter, including void ${r}`,
      );
    }
    for (let r = 0; r < k; r += 1) {
      assert.equal(
        compact.get(`r${r}s1`)!.x - compact.get(`r${r}s0`)!.x,
        PITCH,
        `k=${k}: spacing the author chose inside region ${r} is untouched`,
      );
      assert.equal(compact.get(`r${r}s0`)!.w, TILE, `k=${k}: no tile in region ${r} is compacted away`);
    }
    assert.equal(
      flat.maxX - flat.minX,
      k * REGION_W + (k - 1) * gutter,
      `k=${k}: the drawing is its regions plus one gutter between each pair`,
    );
    assert.equal(
      flat.maxY - flat.minY,
      tall.maxY - tall.minY,
      `k=${k}: the axis with no void is left alone`,
    );
  }
});

test('a void is judged against the box that defines the reach, not the first box on the axis', () => {
  // The sweep carries TWO pieces of running state, and the fixtures above only
  // ever exercised one of them. `reach` is where the drawing has got to; the
  // threshold is `gap + (reachSize + ownSize) / 2`, so `reachSize` has to be
  // the size of whichever box currently defines that reach. Leaving it at the
  // first box's size is invisible on every drawing whose boxes are all the same
  // size - which is every other fixture in this file - and invisible again on
  // separations far above the threshold, where the term is swamped by the gap.
  //
  // It is visible only where the two coincide: a wide box defining the reach, a
  // narrow one opening the drawing, and a gap placed so that the correct size
  // clears the bar and the stale one does not. 600 + (2000 + 100) / 2 = 1650 is
  // over; 600 + (100 + 100) / 2 = 700 is not. Between those two readings sits
  // the difference between closing this void and leaving it open.
  const WIDE = 2000;
  const NARROW = 100;
  const GAP = 600;

  // Measured, not written down: the module does not export its gutter.
  const control = compactEmptyGutters(new Map([
    ['p', box('p', 0, 0)],
    ['q', box('q', 6000, 0)],
  ].map(([, b]) => [(b as ExportBox).id, b as ExportBox])));
  const gutter = control.get('q')!.x - (control.get('p')!.x + control.get('p')!.w);
  assert.ok(gutter > 0 && gutter < 6000, `the control void really closed (${gutter}px)`);

  const first = box('first', 0, 0, NARROW);
  const wide = box('wide', 200, 0, WIDE);
  const far = box('far', wide.x + WIDE + GAP, 0, NARROW);
  const compact = compactEmptyGutters(new Map([first, wide, far].map((b) => [b.id, b])));

  // The gap behind `wide` is 100px and must stay open: nothing about it is a
  // void, and closing it would mean the sweep had stopped distinguishing.
  assert.equal(
    compact.get('wide')!.x - (compact.get('first')!.x + NARROW),
    200 - NARROW,
    'the ordinary gap between the first two boxes is left exactly as authored',
  );
  assert.equal(
    compact.get('far')!.x - (compact.get('wide')!.x + WIDE),
    gutter,
    'the void past the wide box is judged against the wide box and closed',
  );
  assert.equal(compact.get('wide')!.w, WIDE, 'the wide box keeps its width');
  assert.equal(compact.get('far')!.w, NARROW, 'the far box keeps its width');
});

test('an arrow re-anchored onto another side never doubles back through its own tile', () => {
  // A service on the seam of a grid can only be reached by re-anchoring: every
  // lane the router can name finishes inside the neighbour flush against the
  // side the dominant axis chose. The escape is a different connection site,
  // and the obvious way to join two of them — meet on the mid line between the
  // stubs — runs straight back down through the tile it just left whenever the
  // stubs point away from each other. Nothing caught it, because the boxes a
  // route connects are excluded from its own obstacle list.
  const boxes = new Map<string, ExportBox>();
  for (let i = 0; i < 60; i += 1) boxes.set(`g-${i}`, box(`g-${i}`, (i % 10) * 250, Math.floor(i / 10) * 200));
  boxes.set('h-0', box('h-0', 2400, 400));
  const routes = buildExportRoutes(
    [{ id: 'seam', source: 'g-2', target: 'h-0', label: 'Queries' } as Edge],
    boxes,
  );
  assert.equal(routes.length, 1);
  const { points } = routes[0];
  for (const id of ['g-2', 'h-0']) {
    const tile = boxes.get(id)!;
    for (let i = 1; i < points.length; i += 1) {
      const a = points[i - 1];
      const b = points[i];
      const enters = Math.max(a.x, b.x) > tile.x + 1 && Math.min(a.x, b.x) < tile.x + tile.w - 1
        && Math.max(a.y, b.y) > tile.y + 1 && Math.min(a.y, b.y) < tile.y + tile.h - 1;
      assert.ok(!enters, `the hop runs back through ${id}: (${a.x},${a.y})->(${b.x},${b.y})`);
    }
  }
  // And it still reaches both tiles, which is the point of re-anchoring.
  const touches = (p: { x: number; y: number }, t: ExportBox): boolean => p.x >= t.x - 1 && p.x <= t.x + t.w + 1
    && p.y >= t.y - 1 && p.y <= t.y + t.h + 1;
  assert.ok(touches(points[0], boxes.get('g-2')!), 'the hop starts on its source');
  assert.ok(touches(points[points.length - 1], boxes.get('h-0')!), 'the hop ends on its target');
});

test('a fan keeps one lane per member even where a lone hop would be re-anchored', () => {
  // Re-anchoring is gated on the hop standing alone, and "alone" cannot be read
  // off the lane offset: the middle member of an odd fan is handed offset zero
  // and looks exactly like a solitary hop. Moving it puts two arrows, and their
  // two step numbers, on the same line.
  const boxes = new Map<string, ExportBox>();
  for (let i = 0; i < 60; i += 1) boxes.set(`g-${i}`, box(`g-${i}`, (i % 10) * 250, Math.floor(i / 10) * 200));
  boxes.set('h-0', box('h-0', 2400, 400));
  const routes = buildExportRoutes(
    [0, 1, 2].map((i) => ({ id: `seam-${i}`, source: 'g-2', target: 'h-0', label: `Queries ${i}` } as Edge)),
    boxes,
  );
  assert.equal(routes.length, 3);
  // The anchors themselves must stay on the connection site, so separation is
  // in the lane: the paths, and the points the labels hang off, must differ.
  const lanes = routes.map((r) => r.points.map((p) => `${p.x.toFixed(1)},${p.y.toFixed(1)}`).join(' '));
  assert.equal(new Set(lanes).size, 3, `the bundle collapsed onto one lane:\n${lanes.join('\n')}`);
  const anchors = routes.map((r) => `${r.labelAnchor.x.toFixed(1)},${r.labelAnchor.y.toFixed(1)}`);
  assert.equal(new Set(anchors).size, 3, `two labels share a spot: ${anchors.join(' ')}`);
});

test('a drawing too wide for the format gives up its gaps, never its shapes', () => {
  // Visio refuses a page over 200in, and a long cascade at the author's own
  // spacing is well past it. Scaling shrinks the type below the legibility
  // floor and cropping loses services, so the whitespace is what pays.
  const map = new Map<string, ExportBox>();
  for (let i = 0; i < 27; i += 1) map.set(`s${i}`, box(`s${i}`, i * 900, 0));
  const limit = 6000;

  const fitted = fitBoxesWithin(map, limit, limit);
  const after = computeBounds(fitted.values());

  assert.ok(after.maxX - after.minX <= limit + 1, `still ${after.maxX - after.minX}px wide`);
  assert.equal(fitted.size, map.size, 'no service was dropped');
  for (const [id, b] of fitted) {
    assert.equal(b.w, map.get(id)!.w, `${id} kept its width`);
    assert.equal(b.h, map.get(id)!.h, `${id} kept its height`);
  }
  // Reading order and relative position survive: the squeeze is monotonic.
  const xs = [...Array(27).keys()].map((i) => fitted.get(`s${i}`)!.x);
  for (let i = 1; i < xs.length; i += 1) {
    assert.ok(xs[i] > xs[i - 1], `s${i} overtook s${i - 1}`);
    assert.ok(xs[i] - xs[i - 1] >= 150, `s${i} overlaps its neighbour`);
  }
});

test('a drawing that already fits is returned untouched, and a zone keeps its members inside it', () => {
  const map = new Map<string, ExportBox>();
  map.set('zone', { id: 'zone', kind: 'group', category: '', label: 'Zone', x: 0, y: 0, w: 400, h: 300 });
  map.set('a', box('a', 20, 40));
  map.set('far', box('far', 3000, 40));

  assert.equal(fitBoxesWithin(map, 10000, 10000), map, 'a drawing inside the limit is not rebuilt at all');

  const fitted = fitBoxesWithin(map, 800, 10000);
  const zone = fitted.get('zone')!;
  const a = fitted.get('a')!;
  const far = fitted.get('far')!;
  // The zone gives up the empty part of itself, exactly like any other gap. It
  // used to keep all 400px because the union of "shapes that cannot be given
  // up" was taken over every box including this one — which made a rectangle
  // drawn around a whole architecture read as solid, left no whitespace to
  // spend anywhere, and returned the identity map. What a zone must never lose
  // is its contents.
  assert.ok(zone.w < 400, 'the empty part of a zone is whitespace like any other');
  assert.ok(zone.w >= a.w, `the zone was squeezed below the service inside it (${zone.w}px vs ${a.w}px)`);
  assert.ok(zone.x <= a.x && zone.x + zone.w >= a.x + a.w, 'and it still contains the service drawn inside it');
  assert.ok(far.x + far.w <= 801, `the far service came back onto the paper (${far.x + far.w}px)`);
  assert.ok(far.x >= zone.x + zone.w, 'without being pulled inside the zone it was never in');
});

test('shapes alone over the limit are squeezed towards each other but never welded', () => {
  // There is no whitespace left to spend, so the squeeze takes all of it that
  // it is allowed to. What it is not allowed to take is the last sliver: two
  // tiles squeezed flush share an edge, and the hop between them is then a
  // connector from that edge to itself — zero length, no direction, drawn as
  // nothing at all. Every arrow in the drawing disappears to buy tile width
  // the caller's uniform scale would have given back anyway.
  const map = new Map<string, ExportBox>();
  for (let i = 0; i < 10; i += 1) map.set(`s${i}`, box(`s${i}`, i * 400, 0));

  const fitted = fitBoxesWithin(map, 400, 400);
  const xs = [...Array(10).keys()].map((i) => fitted.get(`s${i}`)!);
  for (let i = 1; i < xs.length; i += 1) {
    const gap = xs[i].x - (xs[i - 1].x + xs[i - 1].w);
    assert.ok(gap > 0, `s${i} was welded onto s${i - 1} (gap ${gap}px)`);
    // A gap proportional to the tile stays proportional under the uniform
    // scale the caller applies next, so it is still there on the page.
    assert.ok(gap >= 150 * 0.12 - 1e-9, `s${i} kept only ${gap}px of clearance`);
    assert.ok(gap < 250, `s${i} kept ${gap}px, so the squeeze did no work`);
  }
});

test('a squeeze never widens a gap that was already tighter than the floor', () => {
  // The floor is a limit on how hard the squeeze may push, not a spacing to be
  // imposed: a drawing that was drawn tight stays as tight as it was drawn.
  const map = new Map<string, ExportBox>();
  for (let i = 0; i < 10; i += 1) map.set(`s${i}`, box(`s${i}`, i * 154, 0));

  const fitted = fitBoxesWithin(map, 400, 400);
  for (let i = 1; i < 10; i += 1) {
    const prev = fitted.get(`s${i - 1}`)!;
    const gap = fitted.get(`s${i}`)!.x - (prev.x + prev.w);
    assert.ok(gap > 0 && gap <= 4 + 1e-9, `s${i} came back ${gap}px from its neighbour, not the 4px it was drawn at`);
  }
});

test('an empty band drawn over the drawing does not switch gutter compaction off', () => {
  // A childless zone is treated as the content of the band it names, so that a
  // labelled corridor between two regions is not compacted out of existence.
  // Stretched *over* the drawing instead of standing between things, the same
  // rule made it bridge every void it spanned: two clusters 5,450px apart
  // stayed 5,450px apart under a caption that covered neither of them, and the
  // sheet went out at 74.74in for a 16in drawing.
  const map = new Map<string, ExportBox>();
  map.set('scope', { id: 'scope', kind: 'group', category: '', label: 'Sovereign boundary', x: -80, y: -600, w: 7060, h: 400 });
  for (let i = 0; i < 6; i += 1) map.set(`east${i}`, box(`east${i}`, (i % 3) * 200, Math.floor(i / 3) * 180));
  for (let i = 0; i < 6; i += 1) map.set(`west${i}`, box(`west${i}`, 6000 + (i % 3) * 200, Math.floor(i / 3) * 180));

  const out = compactEmptyGutters(map);
  const span = Math.max(...[...out.values()].map((b) => b.x + b.w)) - Math.min(...[...out.values()].map((b) => b.x));
  assert.ok(span < 3000, `the void under the band was never closed — still ${span.toFixed(0)}px wide`);
  // The band still reaches both clusters: it lost the emptiness underneath it,
  // not its meaning.
  const band = out.get('scope')!;
  assert.ok(band.x <= out.get('east0')!.x, 'the band no longer starts before the drawing it names');
  assert.ok(band.x + band.w >= out.get('west2')!.x + out.get('west2')!.w, 'the band no longer reaches the far cluster');
});

test('a corridor standing between two regions keeps the band it names', () => {
  // The other half of the same rule, and the reason it exists. Nothing is
  // inside this zone and no service overlaps its span on x, so the space it
  // names is the whole point of it and must not be closed.
  const map = new Map<string, ExportBox>();
  map.set('corridor', { id: 'corridor', kind: 'group', category: '', label: 'ExpressRoute circuit', x: 1000, y: -40, w: 3000, h: 400 });
  for (let i = 0; i < 4; i += 1) map.set(`a${i}`, box(`a${i}`, (i % 2) * 200, Math.floor(i / 2) * 180));
  for (let i = 0; i < 4; i += 1) map.set(`b${i}`, box(`b${i}`, 4100 + (i % 2) * 200, Math.floor(i / 2) * 180));

  const out = compactEmptyGutters(map);
  const corridor = out.get('corridor')!;
  assert.ok(corridor.w > 2000, `the labelled corridor was compacted to ${corridor.w.toFixed(0)}px of its 3000`);
});

test('a frame drawn around the drawing does not make the fit a no-op', () => {
  // The union of "what cannot be given up" is taken over the services. Counting
  // an outer rectangle as solid made it span the axis, left no whitespace to
  // spend and returned the identity map, so an ordinary subscription frame cost
  // a 40-service cascade 47% of its tile width to the uniform scaler behind it.
  const build = (framed: boolean): Map<string, ExportBox> => {
    const map = new Map<string, ExportBox>();
    if (framed) map.set('azure', { id: 'azure', kind: 'group', category: '', label: 'Azure', x: -80, y: -80, w: 40 * 900 + 160, h: 400 });
    for (let i = 0; i < 40; i += 1) map.set(`s${i}`, box(`s${i}`, i * 900, 0));
    return map;
  };
  const width = (m: Map<string, ExportBox>): number =>
    Math.max(...[...m.values()].map((b) => b.x + b.w)) - Math.min(...[...m.values()].map((b) => b.x));

  const plain = fitBoxesWithin(build(false), 9000, 100000);
  const framed = fitBoxesWithin(build(true), 9000, 100000);
  assert.ok(width(plain) <= 9001, `the unframed drawing did not fit (${width(plain).toFixed(0)}px)`);
  assert.ok(width(framed) <= 9001, `the framed drawing did not fit (${width(framed).toFixed(0)}px)`);
  // And the frame costs the services only its own padding, not their space. It
  // was costing 22% of the tile width at 27 services and 77% at 90.
  const spread = (m: Map<string, ExportBox>): number => m.get('s39')!.x - m.get('s0')!.x;
  const lost = (spread(plain) - spread(framed)) / spread(plain);
  assert.ok(lost < 0.02, `the frame cost the services ${(lost * 100).toFixed(0)}% of the room they had`);
});

test('a grid packed tighter than the clearance margin still routes around its tiles', () => {
  // 150x75 tiles on a 160x85 pitch leaves 10px gutters — narrower than the 6px
  // clearance the router keeps on each side of a line. clearLanes merged every
  // column into one span and offered no lane at all, so the router took a route
  // that ran the full height of three tiles it does not connect. The clearance
  // is a preference, not a requirement: a line 5px from a tile edge is a far
  // better drawing than a line through it, and one exists here — the 310..320
  // gutter between columns 1 and 2.
  const boxes = new Map<string, ExportBox>();
  for (let i = 0; i < 20; i += 1) boxes.set(`s${i}`, box(`s${i}`, (i % 5) * 160, Math.floor(i / 5) * 85));
  const edges = [
    { id: 'x1', source: 's0', target: 's12' },
    { id: 'x2', source: 's4', target: 's15' },
  ] as Edge[];

  const routes = buildExportRoutes(edges, boxes);
  assert.equal(routes.length, 2, 'both hops must be drawn');
  for (const route of routes) {
    const crossed = [...boxes.values()].filter((tile) => {
      if (tile.id === route.sourceId || tile.id === route.targetId) return false;
      for (let i = 1; i < route.points.length; i += 1) {
        const a = route.points[i - 1];
        const b = route.points[i];
        if (Math.max(a.x, b.x) > tile.x + 1 && Math.min(a.x, b.x) < tile.x + tile.w - 1
          && Math.max(a.y, b.y) > tile.y + 1 && Math.min(a.y, b.y) < tile.y + tile.h - 1) return true;
      }
      return false;
    });
    assert.deepEqual(
      crossed.map((t) => t.id), [],
      `${route.id} is drawn through ${crossed.map((t) => t.id).join(', ')}, which it does not connect`,
    );
  }
});

test('a budget already spent scales down, never inside out', () => {
  // A caller that has spent the page on something else — a Visio sheet whose
  // numbered workflow band is taller than the page Visio will open — hands in a
  // negative budget. A signed ratio turns that into a negative scale, which
  // mirrors every shape about the drawing's own origin and floors every tile at
  // 1px, so the guard against an oversized page produced a larger page than no
  // guard at all.
  const boxes = new Map<string, ExportBox>();
  for (let i = 0; i < 6; i += 1) boxes.set(`s${i}`, box(`s${i}`, i * 300, 0));
  const before = computeBounds(boxes.values());

  for (const budget of [-5000, -1, 0]) {
    const scale = boxScaleWithin(boxes, 4000, budget);
    assert.ok(scale > 0, `budget ${budget} produced a scale of ${scale}`);

    const out = scaleBoxesWithin(boxes, 4000, budget);
    const after = computeBounds(out.values());
    assert.ok(after.maxX >= after.minX && after.maxY >= after.minY,
      `budget ${budget} turned the drawing inside out: ${JSON.stringify(after)}`);
    // Order is what mirroring destroys: s0 must still be left of s5.
    assert.ok(out.get('s0')!.x <= out.get('s5')!.x,
      `budget ${budget} reversed the drawing — s0 landed right of s5`);
    for (const b of out.values()) {
      assert.ok(b.w > 0 && b.h > 0, `budget ${budget} produced a ${b.w}x${b.h} shape`);
    }
  }
  assert.ok(before.maxX > before.minX, 'the fixture must have a positive span to begin with');
});

test('a route to the tile directly below must not leave by the top', () => {
  // routesBetweenSides crosses on a different axis at each of its three shapes:
  // a same-axis pair is joined by a lane perpendicular to its own stubs, while
  // the jog that replaces that lane when it is infeasible runs parallel to
  // them. bestDetour derived one axis flag and handed the same array to all
  // three, so two of them were fed coordinates from the wrong axis: the
  // same-axis lane could never be used at all, and the two-bend lane was wrong
  // whenever the source left by its left or right side. The router then fell
  // through to a detour that left the estate to get somewhere inside it —
  // out of the TOP of a tile, along above the drawing, down its whole height
  // and back up into a target that was directly below where it started.
  const nodes: Node[] = [];
  for (let r = 0; r < 4; r += 1) {
    for (let c = 0; c < 4; c += 1) {
      nodes.push({
        id: `s${r * 4 + c}`, type: 'azureService', width: 150, height: 75,
        position: { x: c * 158, y: r * 83 }, data: { label: 'x', serviceType: 'vm' },
      } as unknown as Node);
    }
  }
  const edges: Edge[] = [];
  for (let i = 0; i < nodes.length; i += 1) {
    for (let j = i + 1; j < nodes.length; j += 1) {
      if ((i * 31 + j * 17) % 5 === 0) edges.push({ id: `e${i}-${j}`, source: `s${i}`, target: `s${j}` } as Edge);
    }
  }
  const boxes = collectExportBoxes(nodes);
  const routes = buildExportRoutes(edges, boxes, { obstacles: [...boxes.values()] });

  let stacked = 0;
  for (const route of routes) {
    const from = boxes.get(route.sourceId)!;
    const to = boxes.get(route.targetId)!;
    if (Math.abs(from.x - to.x) > 1 || to.y <= from.y) continue;
    stacked += 1;
    const above = route.points.filter((p) => p.y < from.y - 1);
    assert.equal(above.length, 0,
      `${route.id} goes up and over to reach a tile below it: ${
        route.points.map((p) => `${Math.round(p.x)},${Math.round(p.y)}`).join(' ')}`);
  }
  assert.ok(stacked >= 4, `the fixture must contain stacked pairs, found ${stacked}`);

  // and the shorter routes must not have been bought by crossing a tile.
  for (const route of routes) {
    const ends = new Set([route.sourceId, route.targetId]);
    for (const b of boxes.values()) {
      if (ends.has(b.id)) continue;
      const inside = route.points.some((p) => p.x > b.x + 1 && p.x < b.x + b.w - 1 && p.y > b.y + 1 && p.y < b.y + b.h - 1);
      assert.ok(!inside, `${route.id} is drawn through ${b.id}`);
    }
  }
});

test('a zone drawn around another is painted behind it, not over it', () => {
  // Authoring order is the trap: `addGroupBoxAtPosition` appends, so a trust
  // boundary drawn around an existing tier is the LAST node. On the canvas the
  // 8-10% zone fill lets the inner tier show through regardless; every export
  // resolves that fill to one opaque colour, so whatever is painted last wins
  // and the inner zone -- with its label -- would simply disappear.
  const nodes = [
    { id: 'inner', type: 'groupNode', position: { x: 120, y: 120 }, style: { width: 160, height: 120 }, data: { label: 'Inner tier' } },
    { id: 'outer', type: 'groupNode', position: { x: 40, y: 40 }, style: { width: 420, height: 340 }, data: { label: 'Outer boundary' } },
  ] as unknown as Node[];

  const { groups } = partitionBoxes(collectExportBoxes(nodes));
  assert.deepEqual(
    groups.map((entry) => entry.id),
    ['outer', 'inner'],
    'the containing zone must be painted first so the zone inside it stays visible',
  );

  const contains = (o: (typeof groups)[number], i: (typeof groups)[number]) =>
    o.x <= i.x && o.y <= i.y && o.x + o.w >= i.x + i.w && o.y + o.h >= i.y + i.h;
  for (let i = 0; i < groups.length; i++) {
    for (let j = i + 1; j < groups.length; j++) {
      assert.ok(
        !contains(groups[j], groups[i]),
        `"${groups[j].label}" is painted after "${groups[i].label}" and encloses it`,
      );
    }
  }
});

test('equal-area zones keep their canvas order', () => {
  // The sort exists only to stop one zone burying another. Zones that cannot
  // enclose each other must stay in authoring order, or exports stop being
  // byte-stable across runs for no benefit.
  const nodes = [
    { id: 'left', type: 'groupNode', position: { x: 0, y: 0 }, style: { width: 200, height: 100 }, data: { label: 'Left' } },
    { id: 'right', type: 'groupNode', position: { x: 400, y: 0 }, style: { width: 200, height: 100 }, data: { label: 'Right' } },
  ] as unknown as Node[];

  const { groups } = partitionBoxes(collectExportBoxes(nodes));
  assert.deepEqual(groups.map((entry) => entry.id), ['left', 'right']);
});