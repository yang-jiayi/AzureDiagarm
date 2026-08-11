// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.
import test from 'node:test';
import assert from 'node:assert/strict';
import { fitGroupsToMembers, GROUP_INNER_PAD_PX, GROUP_HEADER_PAD_PX } from '../src/utils/groupFit';

const OPTS = { nodeWidth: 180, nodeHeight: 100 };

function zone(id: string, w: number, h: number) {
  return { id, position: { x: 0, y: 0 }, width: w, height: h };
}
function member(id: string, groupId: string, x: number, y: number) {
  return { id, groupId, position: { x, y }, width: 150, height: 75 };
}

test('a zone is refitted to hug the services inside it', () => {
  // What the engines used to produce: groupPadding 80 on every side of a
  // 150px tile, so the border was wider than the tile.
  const groups = [zone('z', 510, 915)];
  const services = [
    member('a', 'z', 80, 90),
    member('b', 'z', 80, 290),
    member('c', 'z', 260, 490),
  ];

  const fitted = fitGroupsToMembers(groups, services, OPTS);
  const box = fitted.groups[0];
  const members = fitted.services;

  const minX = Math.min(...members.map((m) => m.position.x));
  const minY = Math.min(...members.map((m) => m.position.y));
  const maxX = Math.max(...members.map((m) => m.position.x + m.width));
  const maxY = Math.max(...members.map((m) => m.position.y + m.height));

  assert.equal(minX, GROUP_INNER_PAD_PX, 'left inset');
  assert.equal(minY, GROUP_HEADER_PAD_PX, 'top inset leaves room for the title');
  assert.equal(box.width, maxX + GROUP_INNER_PAD_PX);
  assert.equal(box.height, maxY + GROUP_INNER_PAD_PX);
  assert.ok(box.width < 510 && box.height < 915, `zone did not shrink: ${box.width}x${box.height}`);

  const waste = 1 - ((maxX - minX) * (maxY - minY)) / (box.width * box.height);
  assert.ok(waste < 0.35, `zone is still ${(waste * 100).toFixed(0)}% empty border`);
});

test('members keep their relative arrangement when the zone is refitted', () => {
  const groups = [zone('z', 800, 700)];
  const services = [member('a', 'z', 100, 120), member('b', 'z', 400, 120), member('c', 'z', 100, 380)];
  const fitted = fitGroupsToMembers(groups, services, OPTS);

  const before = services.map((s) => s.position);
  const after = fitted.services.map((s) => s.position);
  const dx = after[0].x - before[0].x;
  const dy = after[0].y - before[0].y;
  for (let i = 1; i < after.length; i += 1) {
    assert.equal(after[i].x - before[i].x, dx, `member ${i} shifted differently in x`);
    assert.equal(after[i].y - before[i].y, dy, `member ${i} shifted differently in y`);
  }
});

test('an empty zone and an ungrouped service are left alone', () => {
  const groups = [zone('empty', 300, 200), zone('z', 800, 700)];
  const services = [member('a', 'z', 100, 120), { id: 'loose', position: { x: 9, y: 9 }, width: 150, height: 75 }];
  const fitted = fitGroupsToMembers(groups, services, OPTS);

  assert.equal(fitted.groups[0].width, 300, 'an empty zone has nothing to hug');
  assert.equal(fitted.groups[0].height, 200);
  assert.deepEqual(fitted.services[1].position, { x: 9, y: 9 }, 'an ungrouped service must not move');
});

test('a single-service zone stays above the minimum size and centres its tile', () => {
  const groups = [zone('z', 900, 600)];
  const services = [member('only', 'z', 300, 300)];
  const fitted = fitGroupsToMembers(groups, services, OPTS);
  const box = fitted.groups[0];
  const tile = fitted.services[0];

  assert.ok(box.width >= 220 && box.height >= 150, `zone below minimum: ${box.width}x${box.height}`);
  const leftGap = tile.position.x;
  const rightGap = box.width - (tile.position.x + tile.width);
  assert.ok(Math.abs(leftGap - rightGap) < 1, `tile not centred: ${leftGap} vs ${rightGap}`);
  assert.ok(tile.position.y >= GROUP_HEADER_PAD_PX, 'tile must clear the title bar');
});
