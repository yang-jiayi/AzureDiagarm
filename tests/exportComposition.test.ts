// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

import assert from 'node:assert/strict';
import test from 'node:test';
import {
  calculateContentCapturePlan,
  expandDiagramContentBounds,
  screenRectToDiagramBounds,
} from '../src/utils/exportComposition';

test('content capture plan tightly frames a wide diagram with balanced margins', () => {
  const bounds = { x: 120, y: 80, width: 1000, height: 500 };
  const plan = calculateContentCapturePlan(bounds, {
    hasHeader: true,
    legendItemCount: 5,
  });
  const left = plan.transformX + bounds.x * plan.scale;
  const right = plan.width - (plan.transformX + (bounds.x + bounds.width) * plan.scale);
  const top = plan.transformY + bounds.y * plan.scale;
  const bottom = plan.diagramHeight
    - (plan.transformY + (bounds.y + bounds.height) * plan.scale);

  assert.ok(Math.abs(left - right) < 1);
  assert.ok(Math.abs(top - bottom) < 1);
  assert.ok(left >= 70);
  assert.ok(top >= 70);
  assert.equal(plan.legendColumns, 3);
  assert.ok(plan.height > plan.diagramHeight);
});

test('content capture plan avoids a landscape viewport for a tall diagram', () => {
  const plan = calculateContentCapturePlan(
    { x: 0, y: 0, width: 420, height: 1400 },
    { legendItemCount: 5 },
  );

  assert.ok(plan.diagramHeight > plan.width);
  assert.ok(plan.width >= 720);
  assert.ok(plan.scale >= 0.65);
});

test('content capture plan omits optional composition regions', () => {
  const plan = calculateContentCapturePlan({ x: 0, y: 0, width: 600, height: 400 });
  assert.equal(plan.headerHeight, 0);
  assert.equal(plan.legendHeight, 0);
  assert.equal(plan.height, plan.diagramHeight);
});

test('screen-space label rectangles convert back to diagram coordinates', () => {
  const bounds = screenRectToDiagramBounds(
    { left: 540, top: 320, width: 180, height: 60 },
    { x: 240, y: 120 },
    1.5,
  );

  assert.deepEqual(bounds, {
    x: 200,
    y: 400 / 3,
    width: 120,
    height: 40,
  });
});

test('content bounds include offset edge labels with decoration padding', () => {
  const bounds = expandDiagramContentBounds(
    { x: 100, y: 80, width: 600, height: 400 },
    [{ x: 920, y: 220, width: 140, height: 44 }],
    16,
  );

  assert.deepEqual(bounds, {
    x: 100,
    y: 80,
    width: 976,
    height: 400,
  });
});
