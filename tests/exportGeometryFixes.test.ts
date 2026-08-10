import test from 'node:test';
import assert from 'node:assert/strict';
import type { Edge, Node } from 'reactflow';
import {
  buildExportRoutes,
  categoryStyle,
  collectExportBoxes,
  computeBounds,
  computeContentBounds,
  computeFitTransform,
  connectionStyleFor,
  metaSubline,
  routeOrthogonal,
  truncateLabel,
  textWidth,
  usedConnectionLegend,
  wrapLabel,
  zoneStyleFor,
  type ExportBox,
} from '../src/services/diagramExportGeometry.ts';

function box(id: string, x: number, y: number, w = 150, h = 75): ExportBox {
  return { id, kind: 'service', label: id, category: 'other', x, y, w, h };
}

const FRAME = { x: 0.5, y: 1, w: 12, h: 5 };

// ─── Fix 3: one outlier must not collapse the whole diagram ───────────────────

test('fix 3: an extreme outlier is clamped out of the fit bounds', () => {
  // A dense 4×4 grid of tiles near the origin, plus a single stray node far away.
  const boxes: ExportBox[] = [];
  for (let i = 0; i < 4; i += 1) {
    for (let j = 0; j < 4; j += 1) {
      boxes.push(box(`c${i}-${j}`, i * 180, j * 110));
    }
  }
  boxes.push(box('outlier', 8000, 6000));

  const full = computeBounds(boxes);
  const content = computeContentBounds(boxes);

  // The outlier blows the raw bounds out past 8000px…
  assert.ok(full.maxX >= 8000, 'raw bounds include the outlier');
  // …but the clamped content bounds stay with the dense cluster.
  assert.ok(content.maxX < 2000, 'content bounds drop the far outlier');
  assert.ok(content.maxY < 2000, 'content bounds drop the far outlier vertically');

  const clampedScale = computeFitTransform(content, FRAME).scale;
  const rawScale = computeFitTransform(full, FRAME).scale;
  // The normal tiles are now rendered far larger than they would be otherwise.
  assert.ok(clampedScale > rawScale * 3, 'clamping restores a readable scale');
});

test('fix 3: a normal diagram is left untouched by the clamp', () => {
  const boxes = [box('a', 0, 0), box('b', 300, 0), box('c', 0, 200), box('d', 300, 200), box('e', 150, 100)];
  assert.deepEqual(computeContentBounds(boxes), computeBounds(boxes));
});

test('fix 3: three or fewer boxes are never trimmed', () => {
  const boxes = [box('a', 0, 0), box('b', 9000, 9000)];
  assert.deepEqual(computeContentBounds(boxes), computeBounds(boxes));
});

// ─── Fix 5: non-finite coordinates must never corrupt the file ────────────────

test('fix 5: NaN / Infinity coordinates are coerced to safe numbers', () => {
  const nodes = [
    { id: 'nan', type: 'azureNode', position: { x: NaN, y: 10 }, width: Infinity, height: 75, data: { label: 'Broken' } },
    { id: 'ok', type: 'azureNode', position: { x: 100, y: 100 }, width: 150, height: 75, data: { label: 'Fine' } },
  ] as unknown as Node[];

  const boxes = collectExportBoxes(nodes);
  const broken = boxes.get('nan')!;
  assert.ok(Number.isFinite(broken.x) && Number.isFinite(broken.y), 'position is finite');
  assert.ok(Number.isFinite(broken.w) && Number.isFinite(broken.h) && broken.w > 0 && broken.h > 0, 'size is finite and positive');
});

test('fix 5: the fit transform stays finite for a broken bounds box', () => {
  const t = computeFitTransform({ minX: NaN, minY: 0, maxX: Infinity, maxY: 1 }, FRAME);
  assert.ok(Number.isFinite(t.scale) && t.scale > 0, 'scale is finite and positive');
  assert.ok(Number.isFinite(t.offsetX) && Number.isFinite(t.offsetY), 'offsets are finite');
});

// ─── Fix 4: connection semantics are carried into every export ────────────────

test('fix 4: each connection type maps to its canonical colour + dash', () => {
  assert.deepEqual(
    { color: connectionStyleFor('sync').color, dashed: connectionStyleFor('sync').dashed },
    { color: '#64748b', dashed: false },
  );
  assert.equal(connectionStyleFor('async').dashed, true);
  assert.equal(connectionStyleFor('security').color, '#dc2626');
  assert.equal(connectionStyleFor('telemetry').color, '#7c3aed');
  assert.ok(connectionStyleFor('optional').opacity < 1, 'optional links are faded');
});

test('fix 4: routes carry the connection colour and only the used legend is emitted', () => {
  const boxes = new Map([['a', box('a', 0, 0)], ['b', box('b', 600, 0)]]);
  const edges = [
    { id: 'sec', source: 'a', target: 'b', data: { connectionType: 'security' } },
    { id: 'tel', source: 'a', target: 'b', data: { connectionType: 'telemetry' } },
  ] as unknown as Edge[];

  const routes = buildExportRoutes(edges, boxes);
  assert.equal(routes.find((r) => r.id === 'sec')!.color, '#dc2626');
  assert.equal(routes.find((r) => r.id === 'tel')!.color, '#7c3aed');

  const legend = usedConnectionLegend(edges);
  assert.deepEqual(legend.map((entry) => entry.type).sort(), ['security', 'telemetry']);
  assert.equal(usedConnectionLegend([]).length, 0, 'no edges → no legend');
});

// ─── Fix 6: a zone's custom colour must be honoured everywhere ────────────────

test('fix 6: customColor is plumbed into the box and drives the zone style', () => {
  const nodes = [
    { id: 'zone', type: 'groupNode', position: { x: 0, y: 0 }, style: { width: 400, height: 300 }, data: { label: 'Prod', customColor: { border: '#dc2626' } } },
  ] as unknown as Node[];

  const box0 = collectExportBoxes(nodes).get('zone')!;
  assert.equal(box0.customColor?.border, '#dc2626');
  const style = zoneStyleFor(box0, 0);
  assert.equal(style.border, '#dc2626');
  assert.equal(style.text, '#dc2626');
});

test('fix 6: zones without a custom colour share one deterministic palette', () => {
  const plain = box('z', 0, 0);
  plain.kind = 'group';
  // Same index → same colour in every exporter.
  assert.deepEqual(zoneStyleFor(plain, 2), zoneStyleFor(plain, 2));
  assert.notDeepEqual(zoneStyleFor(plain, 0), zoneStyleFor(plain, 1));
});

// ─── Fix 7: self-loops and parallel edges must stay visible & distinct ────────

test('fix 7: a self-loop becomes a visible multi-point stub', () => {
  const boxes = new Map([['a', box('a', 0, 0)]]);
  const routes = buildExportRoutes([{ id: 'loop', source: 'a', target: 'a' }] as unknown as Edge[], boxes);
  assert.equal(routes.length, 1);
  assert.equal(routes[0].isSelfLoop, true);
  assert.ok(routes[0].points.length >= 4, 'the loop has a real footprint');
  // The stub extends beyond the tile's right edge so it is not hidden behind it.
  assert.ok(Math.max(...routes[0].points.map((p) => p.x)) > 150);
});

test('fix 7: parallel edges are fanned so they do not stack byte-for-byte', () => {
  const boxes = new Map([['a', box('a', 0, 0)], ['b', box('b', 600, 300)]]);
  const edges = [
    { id: 'e1', source: 'a', target: 'b' },
    { id: 'e2', source: 'a', target: 'b' },
    { id: 'e3', source: 'a', target: 'b' },
  ] as unknown as Edge[];

  const routes = buildExportRoutes(edges, boxes);
  const signatures = routes.map((r) => JSON.stringify(r.points));
  assert.equal(new Set(signatures).size, 3, 'every parallel edge has a distinct route');
  assert.deepEqual(routes.map((r) => r.ordinal), [0, 1, 2]);
});

// ─── Fix 9: connectors should detour around intervening tiles ─────────────────

test('fix 9: an obstacle on the mid-line pushes the route around it', () => {
  const source = box('a', 0, 0);
  const target = box('b', 800, 0);
  // A blocker straddling the straight-line path between the two tiles.
  const blocker = box('mid', 380, -20, 120, 120);

  const plain = routeOrthogonal(source, target);
  const routed = routeOrthogonal(source, target, { obstacles: [blocker] });

  // The naive route is a straight horizontal line (2 points); the obstacle-aware
  // route must add elbows to get around the blocker.
  assert.ok(routed.points.length > plain.points.length, 'the detour adds bends');
});

// ─── Fix 10: SKU / region / cost metadata sub-line ────────────────────────────

test('fix 10: service metadata is read and formatted as a sub-line', () => {
  const nodes = [
    {
      id: 'svc',
      type: 'azureNode',
      position: { x: 0, y: 0 },
      data: {
        label: 'API',
        sku: 'P1v3',
        region: 'eastus',
        pricing: { estimatedCost: 73.2, quantity: 2 },
      },
    },
  ] as unknown as Node[];

  const box0 = collectExportBoxes(nodes).get('svc')!;
  const line = metaSubline(box0);
  assert.match(line, /P1v3/);
  assert.match(line, /eastus/);
  assert.match(line, /146\.40\/mo/, 'cost is estimatedCost × quantity');
  assert.match(line, / · /, 'parts are dot-separated');
});

test('fix 10: a service with no metadata yields an empty sub-line', () => {
  const box0 = collectExportBoxes([
    { id: 's', type: 'azureNode', position: { x: 0, y: 0 }, data: { label: 'Bare' } },
  ] as unknown as Node[]).get('s')!;
  assert.equal(metaSubline(box0), '');
});

// ─── Fix 14: one wide-character-aware truncation policy ───────────────────────

test('fix 14: truncation counts CJK glyphs as two cells', () => {
  assert.equal(textWidth('あ'), 2, 'a full-width kana is two cells');
  assert.equal(textWidth('A'), 1);

  const jp = 'アプリケーションゲートウェイサービス';
  const cut = truncateLabel(jp, 10);
  assert.ok(cut.endsWith('…'), 'an over-long CJK label is ellipsised');
  assert.ok(textWidth(cut) <= 10, 'the truncated width respects the cell budget');

  // Short labels are returned unchanged.
  assert.equal(truncateLabel('Web App', 48), 'Web App');
});

test('fix 14: wrapping hard-breaks a space-less CJK token', () => {
  const lines = wrapLabel('データベースサーバー', 6, 3);
  assert.ok(lines.length >= 2, 'a long space-less token is broken across lines');
  for (const line of lines) assert.ok(textWidth(line) <= 6);
});

// ─── Fix 19: one canonical category palette shared by all exporters ───────────

test('fix 19: category styling is normalised and shared', () => {
  const a = categoryStyle('AI + Machine Learning');
  const b = categoryStyle('ai + machine learning');
  assert.deepEqual(a, b, 'category keys are case-insensitive');
  assert.match(a.border, /^#[0-9a-fA-F]{6}$/);
  assert.match(a.bg, /^#[0-9a-fA-F]{6}$/);
  // An unknown category falls back to the shared "other" style, never throws.
  assert.match(categoryStyle('totally-made-up').border, /^#[0-9a-fA-F]{6}$/);
});
