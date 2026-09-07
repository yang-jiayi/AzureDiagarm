// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

import assert from 'node:assert/strict';
import test from 'node:test';
import {
  calculateContentCapturePlan,
  expandDiagramContentBounds,
  screenRectToDiagramBounds,
} from '../src/utils/exportComposition';
import { resolveCaptureLegendItems, type CaptureOptions } from '../src/utils/captureCanvas';

test('PNG composition keeps localized legend text while preserving authored paint', () => {
  const composition: NonNullable<CaptureOptions['composition']> = {
    bounds: { x: 0, y: 0, width: 600, height: 300 },
    legendItems: [{
      label: 'Security', description: 'Identity and trust', color: '#dc2626',
      lineStyle: 'dotted', connectionType: 'security',
    }],
    connectionEdges: [{
      id: 'authored', source: 'a', target: 'b', data: { connectionType: 'security' },
      style: { stroke: '#006d77', strokeDasharray: '10 2 3 2', opacity: 0.45 },
    }],
  };
  const before = structuredClone(composition);
  const [item] = resolveCaptureLegendItems(composition);
  assert.equal(item.label, 'Security');
  assert.equal(item.description, 'Identity and trust');
  assert.equal(item.color, '#006d77');
  assert.equal(item.dashPattern, '10, 2, 3, 2');
  assert.equal(item.opacity, 0.45);
  assert.deepEqual(composition, before);
  composition.connectionEdges?.push({
    id: 'default', source: 'a', target: 'b', data: { connectionType: 'security' },
  });
  const [mixed] = resolveCaptureLegendItems(composition);
  assert.equal(mixed.hasMixedStyles, true);
  assert.equal(mixed.label, 'Security (varied)');
  composition.legendVariedLabel = 'styles differ';
  composition.legendVariedDescription = 'Localized explanation';
  const [localized] = resolveCaptureLegendItems(composition);
  assert.equal(localized.label, 'Security (styles differ)');
  assert.equal(localized.description, 'Localized explanation');
});

test('PNG composition retains explicit zero opacity and solid intent', () => {
  const [item] = resolveCaptureLegendItems({
    bounds: { x: 0, y: 0, width: 600, height: 300 },
    legendItems: [{ label: 'Optional', description: '', color: '#64748b', connectionType: 'optional' }],
    connectionEdges: [{
      id: 'invisible', source: 'a', target: 'b', data: { connectionType: 'optional' },
      style: { strokeDasharray: 'none', opacity: 0 },
    }],
  });
  assert.equal(item.lineStyle, 'solid');
  assert.equal(item.dashPattern, undefined);
  assert.equal(item.opacity, 0);
});

test('PNG composition leaves generic legends alone and omits absent connection types', () => {
  const item = { label: 'Custom key', description: '', color: '#123456' };
  const bounds = { x: 0, y: 0, width: 100, height: 100 };
  assert.deepEqual(resolveCaptureLegendItems({ bounds, legendItems: [item] }), [item]);
  assert.deepEqual(resolveCaptureLegendItems({
    bounds, connectionEdges: [],
    legendItems: [{ ...item, connectionType: 'security' }],
  }), []);
});

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

test('content capture plan reserves enough height for an export overlay', () => {
  const plan = calculateContentCapturePlan(
    { x: 0, y: 0, width: 300, height: 80 },
    { minDiagramHeight: 420 },
  );

  assert.equal(plan.diagramHeight, 420);
  assert.equal(plan.height, 420);
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
