// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

/**
 * The export must paint a service the colour the user was looking at.
 *
 * These are not style preferences. The canvas colour is load-bearing: it is the
 * only thing telling a reader that two tiles are the same kind of service. When
 * the export carried its own map, networking was cyan on screen and orange in
 * the file, identity pink on screen and amber in the file, and — worse than any
 * wrong hue — integration and security arrived sharing one red, so a reader of
 * the deck could not tell apart two things the author had deliberately
 * separated. A zone was worse still: its colour came from its position in the
 * list, so dragging a zone repainted it.
 *
 * Both sides now read `canvasPalette`, and these tests are what stops them
 * drifting apart again.
 */

import test from 'node:test';
import assert from 'node:assert/strict';

import {
  CATEGORY_ACCENTS,
  DEFAULT_ACCENT,
  categoryAccent,
  matchZoneAccent,
} from '../src/utils/canvasPalette.ts';
import {
  categoryStyle,
  zoneStyleFor,
  type ExportBox,
} from '../src/services/diagramExportGeometry.ts';

function relativeLuminance(hex: string): number {
  const value = hex.replace('#', '');
  const channel = (pair: string) => {
    const c = parseInt(pair, 16) / 255;
    return c <= 0.03928 ? c / 12.92 : ((c + 0.055) / 1.055) ** 2.4;
  };
  const r = channel(value.slice(0, 2));
  const g = channel(value.slice(2, 4));
  const b = channel(value.slice(4, 6));
  return 0.2126 * r + 0.7152 * g + 0.0722 * b;
}

function contrast(a: string, b: string): number {
  const la = relativeLuminance(a);
  const lb = relativeLuminance(b);
  return (Math.max(la, lb) + 0.05) / (Math.min(la, lb) + 0.05);
}

function zone(label: string, id = 'z1'): ExportBox {
  return { id, kind: 'group', label, category: '', x: 0, y: 0, w: 400, h: 300 };
}

test('every category exports the accent the canvas draws it with', () => {
  for (const category of Object.keys(CATEGORY_ACCENTS)) {
    assert.equal(
      categoryStyle(category).border.toLowerCase(),
      categoryAccent(category).toLowerCase(),
      `${category} exports a different accent than the canvas shows`,
    );
  }
});

test('an unknown category falls back to the same neutral on both sides', () => {
  assert.equal(categoryStyle('no-such-category').border.toLowerCase(), DEFAULT_ACCENT);
  assert.equal(categoryStyle(undefined).border.toLowerCase(), DEFAULT_ACCENT);
});

test('category lookup ignores case and padding, as the canvas does', () => {
  assert.equal(categoryStyle('  Databases  ').border, categoryStyle('databases').border);
});

test('categories the canvas keeps apart never collapse into one export colour', () => {
  // The specific regression: integration and security both exported #C62828,
  // and ai/ml shared compute's blue. Any two categories the canvas gives
  // different accents must stay different in the file.
  const categories = Object.keys(CATEGORY_ACCENTS);
  for (const a of categories) {
    for (const b of categories) {
      if (a === b) continue;
      if (categoryAccent(a).toLowerCase() === categoryAccent(b).toLowerCase()) continue;
      assert.notEqual(
        categoryStyle(a).border.toLowerCase(),
        categoryStyle(b).border.toLowerCase(),
        `${a} and ${b} differ on the canvas but export the same colour`,
      );
    }
  }
});

test('tile text stays readable on the tile fill it is printed on', () => {
  for (const category of [...Object.keys(CATEGORY_ACCENTS), 'no-such-category']) {
    const style = categoryStyle(category);
    assert.ok(
      contrast(style.text, style.bg) >= 4.5,
      `${category}: ${style.text} on ${style.bg} is ${contrast(style.text, style.bg).toFixed(2)}:1`,
    );
  }
});

test('a zone takes its colour from its label, not its position in the list', () => {
  const labels = ['Data Layer', 'Security Perimeter', 'Compute Tier', 'Monitoring'];
  for (const label of labels) {
    const first = zoneStyleFor(zone(label), 0);
    const later = zoneStyleFor(zone(label), 5);
    assert.deepEqual(first, later, `${label} changed colour when it moved down the list`);
    assert.equal(
      first.border.toLowerCase(),
      matchZoneAccent(label)?.toLowerCase(),
      `${label} exports a different accent than the canvas shows`,
    );
  }
});

test('two differently named zones do not borrow each other colours', () => {
  const data = zoneStyleFor(zone('Data Layer', 'a'), 0);
  const security = zoneStyleFor(zone('Security Perimeter', 'b'), 1);
  assert.notEqual(data.border, security.border);
});

test('a picked zone colour still wins over the label', () => {
  const picked = zoneStyleFor(
    { ...zone('Data Layer'), customColor: { border: '#8B0000' } },
    0,
  );
  assert.equal(picked.border.toLowerCase(), '#8b0000');
});

test('an unlabelled zone still gets a colour rather than nothing', () => {
  const style = zoneStyleFor(zone(''), 2);
  assert.ok(/^#[0-9a-f]{6}$/i.test(style.border));
  assert.ok(contrast(style.text, style.bg) >= 4.5);
});

test('zone title text is readable on the zone panel', () => {
  for (const label of ['Ingress', 'AI Services', 'IoT Devices', 'Container Registry', '']) {
    const style = zoneStyleFor(zone(label), 0);
    assert.ok(
      contrast(style.text, style.bg) >= 4.5,
      `${label || '(unlabelled)'}: ${style.text} on ${style.bg}`,
    );
  }
});
