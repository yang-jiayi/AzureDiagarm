// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.
import test from 'node:test';
import assert from 'node:assert/strict';
import {
  planSerpentineWrap,
  wrapPositionedLayout,
  WRAP_TRIGGER_RATIO,
  type WrapBox,
} from '../src/utils/serpentineWrap';

function strip(count: number, pitch = 300, w = 180, h = 100): WrapBox[] {
  return Array.from({ length: count }, (_, i) => ({ id: `n${i}`, x: i * pitch, y: 0, width: w, height: h }));
}

function placed(boxes: WrapBox[], offsets: Map<string, { dx: number; dy: number }>): WrapBox[] {
  return boxes.map((box) => {
    const offset = offsets.get(box.id) ?? { dx: 0, dy: 0 };
    return { ...box, x: box.x + offset.dx, y: box.y + offset.dy };
  });
}

function bbox(boxes: WrapBox[]) {
  const minX = Math.min(...boxes.map((b) => b.x));
  const maxX = Math.max(...boxes.map((b) => b.x + b.width));
  const minY = Math.min(...boxes.map((b) => b.y));
  const maxY = Math.max(...boxes.map((b) => b.y + b.height));
  return { minX, maxX, minY, maxY, w: maxX - minX, h: maxY - minY };
}

test('a linear flow is folded until it fits the shape of a page', () => {
  const boxes = strip(12);
  const before = bbox(boxes);
  assert.ok(before.w / before.h > 20, `expected a strip, got ${(before.w / before.h).toFixed(1)}:1`);

  const offsets = planSerpentineWrap(boxes, { direction: 'LR', bandGap: 200 });
  assert.equal(offsets.size, boxes.length, 'every box must be repositioned');

  const after = bbox(placed(boxes, offsets));
  assert.ok(
    after.w / after.h <= WRAP_TRIGGER_RATIO,
    `wrapped drawing is still ${(after.w / after.h).toFixed(2)}:1`,
  );
});

test('a layout that is already a reasonable shape is left exactly as it was', () => {
  const boxes: WrapBox[] = [
    { id: 'a', x: 0, y: 0, width: 180, height: 100 },
    { id: 'b', x: 300, y: 0, width: 180, height: 100 },
    { id: 'c', x: 0, y: 200, width: 180, height: 100 },
    { id: 'd', x: 300, y: 200, width: 180, height: 100 },
  ];
  assert.equal(planSerpentineWrap(boxes, { direction: 'LR' }).size, 0);
});

test('a drawing too small to have a shape problem is never touched', () => {
  const boxes = strip(2, 4000);
  assert.equal(planSerpentineWrap(boxes, { direction: 'LR' }).size, 0);
});

test('no service is lost, duplicated or left overlapping another', () => {
  const boxes = strip(18);
  const offsets = planSerpentineWrap(boxes, { direction: 'LR', bandGap: 200 });
  const out = placed(boxes, offsets);
  assert.equal(new Set(out.map((b) => b.id)).size, boxes.length);

  for (let i = 0; i < out.length; i += 1) {
    for (let j = i + 1; j < out.length; j += 1) {
      const a = out[i];
      const b = out[j];
      const overlapX = Math.min(a.x + a.width, b.x + b.width) - Math.max(a.x, b.x);
      const overlapY = Math.min(a.y + a.height, b.y + b.height) - Math.max(a.y, b.y);
      assert.ok(
        overlapX <= 0.001 || overlapY <= 0.001,
        `${a.id} and ${b.id} overlap after wrapping`,
      );
    }
  }
});

test('the hop between two bands is short instead of sweeping back across the drawing', () => {
  const boxes = strip(12);
  const offsets = planSerpentineWrap(boxes, { direction: 'LR', bandGap: 200 });
  const out = placed(boxes, offsets);
  const byId = new Map(out.map((b) => [b.id, b]));

  // The seams are wherever two consecutive services ended up on different rows.
  const seams: number[] = [];
  for (let i = 1; i < boxes.length; i += 1) {
    const prev = byId.get(`n${i - 1}`)!;
    const next = byId.get(`n${i}`)!;
    if (Math.abs(next.y - prev.y) > 1) seams.push(i);
  }
  assert.ok(seams.length >= 1, 'expected at least one band seam');

  const bandWidth = bbox(out).w;
  for (const i of seams) {
    const prev = byId.get(`n${i - 1}`)!;
    const next = byId.get(`n${i}`)!;
    const hop = Math.abs(next.x - prev.x);
    assert.ok(
      hop < bandWidth * 0.5,
      `seam ${i - 1}->${i} sweeps ${hop.toFixed(0)}px across a ${bandWidth.toFixed(0)}px band`,
    );
  }
});

test('a top-down flow that is too tall is folded into columns instead', () => {
  const boxes: WrapBox[] = Array.from({ length: 10 }, (_, i) => ({
    id: `n${i}`, x: 0, y: i * 300, width: 180, height: 100,
  }));
  const offsets = planSerpentineWrap(boxes, { direction: 'TB', bandGap: 200 });
  assert.equal(offsets.size, boxes.length);
  const after = bbox(placed(boxes, offsets));
  assert.ok(after.h / after.w <= WRAP_TRIGGER_RATIO, `still ${(after.h / after.w).toFixed(2)}:1 tall`);
});

test('a zone carries its members with it, and members keep their relative places', () => {
  const groups = Array.from({ length: 6 }, (_, i) => ({
    id: `z${i}`, position: { x: i * 900, y: 0 }, width: 700, height: 300,
  }));
  const services = groups.flatMap((group, i) => ([
    { id: `s${i}a`, groupId: group.id, position: { x: 40, y: 60 }, width: 180, height: 100 },
    { id: `s${i}b`, groupId: group.id, position: { x: 400, y: 60 }, width: 180, height: 100 },
  ]));

  const wrapped = wrapPositionedLayout(services, groups, {
    direction: 'LR', bandGap: 200, nodeWidth: 180, nodeHeight: 100,
  });

  assert.ok(wrapped.bands > 1, 'a six-zone strip must wrap');
  // Members are stored relative to their zone, so wrapping must not move them.
  for (const service of wrapped.services) {
    const original = services.find((s) => s.id === service.id)!;
    assert.deepEqual(service.position, original.position, `${service.id} moved out of its zone`);
  }
  const zoneRows = new Set(wrapped.groups.map((g) => g.position.y));
  assert.ok(zoneRows.size > 1, 'zones must end up on more than one row');
});

test('an ungrouped service standing next to zones is wrapped along with them', () => {
  const groups = Array.from({ length: 5 }, (_, i) => ({
    id: `z${i}`, position: { x: i * 900, y: 0 }, width: 700, height: 300,
  }));
  const services = [
    ...groups.map((group, i) => ({ id: `s${i}`, groupId: group.id, position: { x: 40, y: 60 }, width: 180, height: 100 })),
    { id: 'loose', position: { x: 4600, y: 0 }, width: 180, height: 100 },
  ];

  const wrapped = wrapPositionedLayout(services, groups, {
    direction: 'LR', bandGap: 200, nodeWidth: 180, nodeHeight: 100,
  });
  const loose = wrapped.services.find((s) => s.id === 'loose')!;
  assert.notDeepEqual(loose.position, { x: 4600, y: 0 }, 'the loose service was left behind');
});

// --- Downstream of the wrap: the two paths that the fold invalidated ---

async function wrappedChain(emphasize: boolean) {
  const { applyLayoutPreset } = await import('../src/utils/layoutPresets.ts');
  const nodes = Array.from({ length: 12 }, (_, i) => ({
    id: `c${i}`,
    type: 'azureNode',
    position: { x: 0, y: 0 },
    width: 180,
    height: 100,
    data: { label: `Service ${i}` },
  })) as any[];
  const edges = Array.from({ length: 11 }, (_, i) => ({
    id: `ce${i}`,
    source: `c${i}`,
    target: `c${i + 1}`,
    sourceHandle: 'right',
    targetHandle: 'left',
    label: 'Invoke',
  })) as any[];
  return applyLayoutPreset(nodes, edges, {
    preset: 'flow-lr',
    spacing: 'comfortable',
    edgeStyle: 'smooth',
    emphasizePrimaryPath: emphasize,
  });
}

test('emphasising the primary path straightens each band, it does not stack them', async () => {
  const { nodes } = await wrappedChain(true);
  const tiles = nodes.filter((n: any) => n.type === 'azureNode');
  const rows = new Set(tiles.map((n: any) => Math.round(n.position.y)));

  // Snapping the whole chain to one median y would flatten the fold back into
  // the strip it was folded out of.
  assert.ok(rows.size > 1, `all ${tiles.length} tiles collapsed onto y=${[...rows][0]}`);

  const overlaps: string[] = [];
  for (let i = 0; i < tiles.length; i += 1) {
    for (let j = i + 1; j < tiles.length; j += 1) {
      const a = tiles[i] as any;
      const b = tiles[j] as any;
      const dx = Math.min(a.position.x + a.width, b.position.x + b.width) - Math.max(a.position.x, b.position.x);
      const dy = Math.min(a.position.y + a.height, b.position.y + b.height) - Math.max(a.position.y, b.position.y);
      if (dx > 1 && dy > 1) overlaps.push(`${a.id}/${b.id}`);
    }
  }
  assert.deepEqual(overlaps, [], `tiles overlap after straightening: ${overlaps.join(', ')}`);
});

test('a reversed band gets connectors that leave the face they actually travel towards', async () => {
  const { nodes, edges } = await wrappedChain(false);
  const at = new Map(nodes.map((n: any) => [n.id, n.position.x]));

  const backwards = edges.filter((e: any) => {
    const from = at.get(e.source);
    const to = at.get(e.target);
    if (from === undefined || to === undefined) return false;
    // Leaving the right face towards a target that sits to the left draws a
    // loop back across both tiles — the opposite of what the badge asserts.
    return e.sourceHandle === 'right' && to < from;
  });
  assert.deepEqual(
    backwards.map((e: any) => e.id),
    [],
    `${backwards.length}/${edges.length} connectors point backwards`,
  );

  const reversed = edges.filter((e: any) => (at.get(e.target) ?? 0) < (at.get(e.source) ?? 0));
  assert.ok(reversed.length > 0, 'the fixture no longer produces a reversed band');
  for (const edge of reversed) {
    assert.equal((edge as any).sourceHandle, 'left-source');
    assert.equal((edge as any).targetHandle, 'right-target');
  }
});

test('a re-arrange keeps the handles the layout realigned', async () => {
  const { mergeLayoutEdges } = await import('../src/utils/layoutResultMerge.ts');
  const before = [{ id: 'e1', source: 'a', target: 'b', sourceHandle: 'right', targetHandle: 'left' }] as any[];
  const laidOut = [{ id: 'e1', source: 'a', target: 'b', sourceHandle: 'left-source', targetHandle: 'right-target' }] as any[];

  // `applyLayout` feeds the layout's edges through this merge before setting
  // state; handles are top-level Edge fields, so a merge that only rebuilds
  // `data` silently restores the pre-layout direction.
  const merged = mergeLayoutEdges(before, before, laidOut);
  assert.equal(merged[0].sourceHandle, 'left-source', 'realigned source handle was discarded');
  assert.equal(merged[0].targetHandle, 'right-target', 'realigned target handle was discarded');

  const handAttached = [{ id: 'e1', source: 'a', target: 'b', sourceHandle: 'top', targetHandle: 'bottom' }] as any[];
  const kept = mergeLayoutEdges(handAttached, before, laidOut);
  assert.equal(kept[0].sourceHandle, 'top', 'a hand-attached handle must survive a re-arrange');
});

test('bands of different widths are still straightened separately', async () => {
  const { straightenPrimaryPath } = await import('../src/utils/layoutPresets.ts');
  // Two bands centred on the widest, so the seam makes a large FORWARD jump on
  // the major axis: the reversal test cannot see it and a seam test that reads
  // the major step cannot either.
  const xs = [0, 400, 800, 1200, 1600, 2000, 2400, 2800, 3800, 3400, 3000, 2600];
  const ys = [0, 0, 0, 0, 0, 0, 0, 0, 400, 400, 400, 400];
  const nodes = xs.map((x, i) => ({
    id: `n${i}`, type: 'azureNode', position: { x, y: ys[i] }, width: 180, height: 100, data: { label: `S${i}` },
  })) as any[];
  const edges = xs.slice(1).map((_, i) => ({ id: `e${i}`, source: `n${i}`, target: `n${i + 1}` })) as any[];

  const result = straightenPrimaryPath(nodes, edges, 'LR');
  const placed = new Map(result.nodes.map((n: any) => [n.id, Math.round(n.position.y)]));
  const strayed = nodes
    .map((n: any, i: number) => ({ id: n.id, want: ys[i], got: placed.get(n.id) }))
    .filter((n) => n.got !== n.want);
  assert.deepEqual(
    strayed,
    [],
    `nodes pulled into the wrong band: ${strayed.map((n) => `${n.id} ${n.want}->${n.got}`).join(', ')}`,
  );
});
