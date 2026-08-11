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
