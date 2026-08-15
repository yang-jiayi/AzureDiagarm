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
import JSZip from 'jszip';
import type { Node } from 'reactflow';
import { buildDiagramSlidePptx } from '../src/services/pptxExporter.ts';
import { nativizeSlideXml } from '../src/services/pptxNativeShapes.ts';

const PIXEL_PNG =
  'data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg==';

import {
  CATEGORY_ACCENTS,
  DEFAULT_ACCENT,
  categoryAccent,
  matchZoneAccent,
  zoneAccent,
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

/** What `rgba(accent, alpha)` over `paper` actually resolves to on screen. */
function composite(accent: string, alpha: number, paper: string): string {
  const pair = (hex: string, i: number) => parseInt(hex.replace('#', '').slice(i * 2, i * 2 + 2), 16);
  const mixed = [0, 1, 2]
    .map((i) => Math.round(pair(accent, i) * alpha + pair(paper, i) * (1 - alpha)))
    .map((v) => v.toString(16).padStart(2, '0'))
    .join('');
  return `#${mixed}`;
}

/** The worst single-channel gap between two colours, in 0-255 terms. */
function channelDrift(a: string, b: string): number {
  const pair = (hex: string, i: number) => parseInt(hex.replace('#', '').slice(i * 2, i * 2 + 2), 16);
  return Math.max(...[0, 1, 2].map((i) => Math.abs(pair(a, i) - pair(b, i))));
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
  // `zoneAccent` is what GroupNode paints, keyword hit or not, so it is the
  // exact oracle. The unmatched labels below are the ones that used to come
  // out blue/green/orange in the file while the canvas drew them grey — and
  // that changed colour whenever the zone moved up or down the list.
  const labels = [
    'Data Layer',
    'Security Perimeter',
    'Compute Tier',
    'Monitoring',
    'Hub',
    'Shared Services',
    'DMZ',
    'Landing Zone',
    'On-Prem',
  ];
  for (const label of labels) {
    const style = zoneStyleFor(zone(label));
    assert.equal(
      style.border.toLowerCase(),
      zoneAccent(label).toLowerCase(),
      `${label} exports a different accent than the canvas shows`,
    );
  }
});

test('two zones the keyword table does not know stay the same colour as each other', () => {
  // They are indistinguishable to the canvas, so distinguishing them in the
  // file is an invention — and the only thing that could tell them apart is
  // the ordering that used to leak in.
  const hub = zoneStyleFor(zone('Hub', 'a'));
  const shared = zoneStyleFor(zone('Shared Services', 'b'));
  assert.equal(matchZoneAccent('Hub'), null);
  assert.equal(matchZoneAccent('Shared Services'), null);
  assert.deepEqual(hub, shared);
});

test('two differently named zones do not borrow each other colours', () => {
  const data = zoneStyleFor(zone('Data Layer', 'a'));
  const security = zoneStyleFor(zone('Security Perimeter', 'b'));
  assert.notEqual(data.border, security.border);
});

test('a picked zone colour still wins over the label', () => {
  const picked = zoneStyleFor({
    ...zone('Data Layer'),
    customColor: { border: '#8B0000' },
  });
  assert.equal(picked.border.toLowerCase(), '#8b0000');
});

test('an unlabelled zone still gets a colour rather than nothing', () => {
  const style = zoneStyleFor(zone(''));
  assert.ok(/^#[0-9a-f]{6}$/i.test(style.border));
  assert.ok(contrast(style.text, style.bg) >= 4.5);
});

test('zone title text is readable on the zone panel', () => {
  for (const label of ['Ingress', 'AI Services', 'IoT Devices', 'Container Registry', '']) {
    const style = zoneStyleFor(zone(label));
    assert.ok(
      contrast(style.text, style.bg) >= 4.5,
      `${label || '(unlabelled)'}: ${style.text} on ${style.bg}`,
    );
  }
});

test('a zone panel is the colour the canvas composites, on dark paper as well as light', () => {
  // GroupNode fills a zone with `rgba(accent, 0.10)` — matched — or `0.08`
  // when the keyword table does not know the label. That is *translucent*, so
  // what the reader sees is the accent over whatever paper the app is on. The
  // export used to mix toward white unconditionally, which on a light deck is
  // the same answer by luck, and on a dark deck is an inversion: a barely-there
  // tint on `#1e293b` came out as `F0F1F2`, a glaring white block on a dark
  // slide. Reading it back is not a matter of taste — a zone that reads as the
  // brightest thing on the page tells the reader it is the subject.
  const paper = [
    { theme: 'light', bg: '#ffffff' },
    { theme: 'dark', bg: '#1e293b' },
  ];
  const labels = ['Data Tier', 'Security', 'Compute', 'Networking', 'Hub', 'Shared Services'];
  for (const { theme, bg } of paper) {
    for (const label of labels) {
      const accent = zoneAccent(label);
      const alpha = matchZoneAccent(label) ? 0.1 : 0.08;
      const style = zoneStyleFor(zone(label), bg);
      const drift = channelDrift(style.bg, composite(accent, alpha, bg));
      assert.ok(
        drift <= 6,
        `${theme} ${label}: export ${style.bg} vs canvas ${composite(accent, alpha, bg)} `
        + `(off by ${drift}/255)`,
      );
      assert.ok(
        contrast(style.text, style.bg) >= 4.5,
        `${theme} ${label}: title ${style.text} on ${style.bg} is `
        + `${contrast(style.text, style.bg).toFixed(2)}:1`,
      );
    }
  }
});

test('a zone panel is never more than a slight step away from the paper it is on', () => {
  // The blunt version of the test above, and the one that would have caught the
  // original bug on its own without knowing any accent. A fill at 8-10% alpha
  // cannot be far from what is behind it — that is what low alpha *means*. The
  // broken dark export put an `F0F1F2` panel on a `1E293B` slide: 14:1 apart,
  // which is not a tint of anything, it is a new opaque object.
  for (const paper of ['#ffffff', '#1e293b']) {
    for (const label of ['Data Tier', 'Security', 'Hub']) {
      const style = zoneStyleFor(zone(label), paper);
      const step = contrast(style.bg, paper);
      assert.ok(
        step < 1.5,
        `${label} on ${paper}: panel ${style.bg} is ${step.toFixed(2)}:1 from the paper, `
        + 'which is an opaque block rather than a tint',
      );
    }
  }
});

test('the zone fill PowerPoint actually paints is the colour the canvas shows', async () => {
  // The two tests above characterise `zoneStyleFor`, which is a claim about a
  // helper, not about the file. The deck paints `fill: { color: bg,
  // transparency: 15 }` -- so the pixel a reader sees is that colour at 85%
  // over the slide, and every assertion above would stay green if
  // `addGroupShape` stopped calling `zoneStyleFor`, changed the transparency
  // byte, or the nativizer rewrote the fill. This reads the delivered slide.
  for (const [themeName, isDark, slidePaper] of [
    ['light', false, '#ffffff'],
    ['dark', true, '#1e293b'],
  ] as const) {
    const nodes = [
      { id: 'z', type: 'groupNode', position: { x: 0, y: 0 }, style: { width: 600, height: 400 },
        data: { label: 'Data Tier' } },
      { id: 's', type: 'azureNode', position: { x: 80, y: 80 }, width: 150, height: 75,
        data: { label: 'SQL', serviceName: 'SQL Database' } },
    ] as unknown as Node[];

    const pptx = await buildDiagramSlidePptx(PIXEL_PNG, {
      diagramName: 'Zone fill', author: 'Tester', date: '2026-08-10',
      isDarkMode: isDark, diagram: { nodes, edges: [] },
    });
    const zip = await JSZip.loadAsync((await pptx.write({ outputType: 'nodebuffer' })) as Buffer);
    const xml = nativizeSlideXml(await zip.file('ppt/slides/slide1.xml')!.async('string'));

    const shape = xml.split(/<p:sp>|<p:grpSp>/).find((part) => /name="zone-z"/.test(part));
    assert.ok(shape, `${themeName}: no zone shape in the delivered slide`);
    // The first solidFill in a shape can be the outline; drop <a:ln> first.
    const body = shape.replace(/<a:ln[\s\S]*?<\/a:ln>/g, '');
    const fill = /<a:solidFill><a:srgbClr val="([0-9A-Fa-f]{6})"(?:[^>]*)>?\s*(?:<a:alpha val="(\d+)"\/>)?/.exec(body);
    assert.ok(fill, `${themeName}: zone shape carries no solid fill`);

    const declared = `#${fill[1].toLowerCase()}`;
    const alpha = fill[2] === undefined ? 1 : Number(fill[2]) / 100000;
    const delivered = composite(declared, alpha, slidePaper);

    const canvas = composite(zoneAccent('Data Tier'), matchZoneAccent('Data Tier') ? 0.1 : 0.08, slidePaper);
    const drift = channelDrift(delivered, canvas);
    assert.ok(
      drift <= 6,
      `${themeName}: the slide paints ${declared} at ${(alpha * 100).toFixed(0)}% = ${delivered}, `
      + `but the canvas shows ${canvas} (off by ${drift}/255)`,
    );
    assert.ok(
      contrast(delivered, slidePaper) < 1.5,
      `${themeName}: the delivered zone reads as an opaque block (${contrast(delivered, slidePaper).toFixed(2)}:1 from the slide)`,
    );
  }
});