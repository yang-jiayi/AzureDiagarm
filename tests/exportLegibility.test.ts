import test from 'node:test';
import assert from 'node:assert/strict';
import JSZip from 'jszip';
import type { Edge, Node } from 'reactflow';
import { buildDiagramSlidePptx } from '../src/services/pptxExporter.ts';
import { buildVsdxPackage } from '../src/services/visioVsdxExporter.ts';

/**
 * Legibility regression guard for the native-shape exports.
 *
 * A wide architecture used to be squeezed onto one fixed 13.33" slide while
 * font sizes, chip widths and icon sizes stayed absolute, so service tiles fell
 * to 0.55" with 6 pt labels that wrapped one character per line and 2.4" edge
 * chips that covered the nodes. These tests measure the emitted shape XML so
 * that can never come back.
 */

const EMU_PER_INCH = 914400;
const PIXEL_PNG =
  'data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg==';

interface Shape { name: string; x: number; y: number; w: number; h: number; text: string; fontSize: number | null }

/** Approximate rendered text width in inches; CJK glyphs are full-width. */
function textWidthIn(text: string, fontSizePt: number): number {
  let units = 0;
  for (const character of text) {
    units += /[\u2e80-\u9fff\uff00-\uffef]/.test(character) ? 1 : 0.54;
  }
  return (units * fontSizePt) / 72;
}

function parseShapes(xml: string): Shape[] {
  const shapes: Shape[] = [];
  const shapeRe = /<p:(sp|pic)>([\s\S]*?)<\/p:\1>/g;
  let match: RegExpExecArray | null;
  while ((match = shapeRe.exec(xml))) {
    const body = match[2];
    const off = /<a:off x="(-?\d+)" y="(-?\d+)"\/>/.exec(body);
    const ext = /<a:ext cx="(\d+)" cy="(\d+)"\/>/.exec(body);
    if (!off || !ext) continue;
    const size = /sz="(\d+)"/.exec(body);
    shapes.push({
      name: /name="([^"]*)"/.exec(body)?.[1] ?? '',
      x: +off[1] / EMU_PER_INCH,
      y: +off[2] / EMU_PER_INCH,
      w: +ext[1] / EMU_PER_INCH,
      h: +ext[2] / EMU_PER_INCH,
      text: [...body.matchAll(/<a:t>([\s\S]*?)<\/a:t>/g)].map((t) => t[1]).join(''),
      fontSize: size ? +size[1] / 100 : null,
    });
  }
  return shapes;
}

function service(id: string, label: string, x: number, y: number, parent?: string): Node {
  return {
    id,
    type: 'azureNode',
    position: { x, y },
    width: 150,
    height: 75,
    ...(parent ? { parentNode: parent } : {}),
    data: {
      label,
      serviceName: label,
      iconPath: '/Azure_Public_Service_Icons/Icons/compute/10021-icon-service-Virtual-Machine.svg',
    },
  } as Node;
}

/** Four zones side by side — the shape a real AI-generated diagram takes. */
function wideDiagram(): { nodes: Node[]; edges: Edge[] } {
  const nodes: Node[] = [];
  const edges: Edge[] = [];
  const names = [
    'Copilot Studio', 'Key Vault', 'Azure OpenAI Service', 'Azure AI Search',
    'Azure Kubernetes Service', 'Azure SQL Database', 'Application Gateway', 'Azure Front Door',
    'Azure Functions', 'Azure Service Bus', 'Azure Data Factory', 'Azure Synapse Analytics',
  ];
  ['Ingress zone', 'Application zone', 'Data zone', 'Integration zone'].forEach((zone, z) => {
    nodes.push({
      id: `zone-${z}`,
      type: 'groupNode',
      position: { x: z * 900, y: 0 },
      style: { width: 820, height: 560 },
      data: { label: zone },
    } as Node);
    for (let i = 0; i < 3; i += 1) {
      const index = z * 3 + i;
      nodes.push(service(`svc-${index}`, names[index], 60 + (i % 2) * 380, 90 + Math.floor(i / 2) * 200, `zone-${z}`));
    }
  });
  for (let i = 0; i < 11; i += 1) {
    edges.push({
      id: `e-${i}`,
      source: `svc-${i}`,
      target: `svc-${i + 1}`,
      label: i % 2 === 0 ? 'HTTPS 経由でトークン検証を実施' : 'Managed identity authentication',
    } as Edge);
  }
  return { nodes, edges };
}

async function slideShapes(diagram: { nodes: Node[]; edges: Edge[] }): Promise<{ shapes: Shape[]; slideW: number; parts: Shape[][] }> {
  const pptx = await buildDiagramSlidePptx(PIXEL_PNG, {
    diagramName: 'Contoso Platform',
    author: 'Tester',
    date: '2026-08-10',
    isDarkMode: false,
    diagram,
  });
  const buffer = (await pptx.write({ outputType: 'nodebuffer' })) as Buffer;
  const zip = await JSZip.loadAsync(buffer);
  const presentation = await zip.file('ppt/presentation.xml')!.async('string');
  const slideW = +(/<p:sldSz cx="(\d+)"/.exec(presentation)?.[1] ?? '0') / EMU_PER_INCH;
  const all = await Promise.all(
    Object.keys(zip.files)
      .filter((name) => /^ppt\/slides\/slide\d+\.xml$/.test(name))
      .sort((a, b) => +a.replace(/\D/g, '') - +b.replace(/\D/g, ''))
      .map((name) => zip.file(name)!.async('string')),
  );
  // A drawing too big for one readable page is now continued across ordinary
  // slides, and such a deck opens with the whole thing shown small on purpose.
  // Legibility is a property of the parts, so measure those; reading slide 1
  // blind would grade the export on the one slide that is meant to be a
  // thumbnail.
  const overviewAt = all.findIndex((xml) => xml.includes('(Overview)'));
  const parts = (overviewAt < 0 ? all : all.slice(overviewAt + 1)).map(parseShapes);
  const withTiles = parts.find((shapes) => shapes.some((s) => /^service-[^l]/.test(s.name) && !s.name.startsWith('service-label') && !s.name.startsWith('service-meta')));
  return { shapes: withTiles ?? parseShapes(all[0]), slideW, parts };
}

test('a wide diagram keeps its shapes full size instead of shrinking them', async () => {
  const { shapes, slideW, parts } = await slideShapes(wideDiagram());
  assert.ok(slideW <= 56.01, `page must stay within PowerPoint's 56in limit, got ${slideW.toFixed(2)}in`);

  const isTile = (s: Shape): boolean => /^service-[^l]/.test(s.name) && !s.name.startsWith('service-label') && !s.name.startsWith('service-meta');
  for (const tile of shapes.filter(isTile)) {
    // 7pt is the product's legibility floor and a 75px tile renders its label
    // at h * 12, so 7/12in is the smallest tile the exporter may emit. A
    // continued drawing sits near that floor by design; anything under it is
    // the old "squeeze it onto one sheet" defect.
    assert.ok(tile.w > 1.15, `tile ${tile.name} shrank to ${tile.w.toFixed(2)}in`);
  }
  // Whether the page grew or the drawing was continued across slides, all
  // twelve services must reach the deck exactly once.
  const drawn = new Map<string, number>();
  for (const part of parts) {
    for (const tile of part.filter(isTile)) drawn.set(tile.name, (drawn.get(tile.name) ?? 0) + 1);
  }
  assert.equal(drawn.size, 12, `expected 12 services across the deck, got ${drawn.size}`);
  const duplicated = [...drawn].filter(([, count]) => count !== 1);
  assert.equal(duplicated.length, 0, `services drawn more than once: ${duplicated.map(([name]) => name).join(', ')}`);
});

test('node labels never wrap to a few characters per line', async () => {
  const { shapes } = await slideShapes(wideDiagram());
  const labels = shapes.filter((s) => s.name.startsWith('service-label-'));
  assert.ok(labels.length > 0);
  for (const label of labels) {
    const font = label.fontSize ?? 11;
    // The exporter's own floor. A drawing continued across slides is rendered
    // at whatever size keeps it here, so this is the number to hold, not a
    // larger one that only the old grow-the-page behaviour could reach.
    assert.ok(font >= 7, `label "${label.text}" fell to ${font}pt`);
    const charsPerLine = label.w / (font / 72);
    assert.ok(charsPerLine >= 8, `label "${label.text}" fits only ${charsPerLine.toFixed(1)} chars per line`);
    const lines = Math.ceil(textWidthIn(label.text, font) / Math.max(label.w, 0.01));
    assert.ok(
      lines * (font * 1.25) / 72 <= label.h + 0.03,
      `label "${label.text}" needs ${lines} lines but only has ${label.h.toFixed(2)}in`,
    );
  }
});

test('edge label chips stay smaller than a service tile and clear of it', async () => {
  const { shapes } = await slideShapes(wideDiagram());
  const tiles = shapes.filter((s) => /^service-svc-\d+$/.test(s.name));
  const chips = shapes.filter((s) => s.name.startsWith('connector-label-'));
  const smallestTile = Math.min(...tiles.map((t) => t.w));
  assert.ok(chips.length > 0);
  for (const chip of chips) {
    assert.ok(chip.w <= smallestTile, `chip "${chip.text}" is ${chip.w.toFixed(2)}in vs a ${smallestTile.toFixed(2)}in tile`);
    for (const tile of tiles) {
      const overlapW = Math.min(chip.x + chip.w, tile.x + tile.w) - Math.max(chip.x, tile.x);
      const overlapH = Math.min(chip.y + chip.h, tile.y + tile.h) - Math.max(chip.y, tile.y);
      const area = overlapW > 0 && overlapH > 0 ? overlapW * overlapH : 0;
      assert.ok(
        area <= 0.02 * tile.w * tile.h,
        `chip "${chip.text}" covers ${((area / (tile.w * tile.h)) * 100).toFixed(0)}% of ${tile.name}`,
      );
    }
  }
});

test('a diagram beyond the page limit degrades proportionally, not typographically', async () => {
  const nodes: Node[] = [];
  const edges: Edge[] = [];
  for (let i = 0; i < 40; i += 1) {
    nodes.push(service(`n-${i}`, i % 2 ? 'Azure Kubernetes Service' : 'Copilot Studio', i * 260, (i % 4) * 220));
    if (i > 0) edges.push({ id: `x-${i}`, source: `n-${i - 1}`, target: `n-${i}`, label: 'Managed identity authentication' } as Edge);
  }
  const { shapes, slideW } = await slideShapes({ nodes, edges });
  assert.ok(slideW <= 56.01, `page must stay within the 56in limit, got ${slideW.toFixed(2)}in`);

  const tiles = shapes.filter((s) => /^service-n-\d+$/.test(s.name));
  const chips = shapes.filter((s) => s.name.startsWith('connector-label-'));
  const smallestTile = Math.min(...tiles.map((t) => t.w));
  for (const chip of chips) {
    assert.ok(chip.w <= smallestTile, `chip "${chip.text}" outgrew the tiles when scaled down`);
  }
  const labels = shapes.filter((s) => s.name.startsWith('service-label-'));
  for (const label of labels) {
    // The ratio of glyph size to box width is what keeps the wrap count sane.
    const charsPerLine = label.w / ((label.fontSize ?? 11) / 72);
    assert.ok(charsPerLine >= 8, `scaled-down label "${label.text}" fits only ${charsPerLine.toFixed(1)} chars per line`);
  }
});

test('an arrow chip is never written smaller than the names beside it', async () => {
  // A chip carried its own 4pt floor while a tile name stopped at 7, so a
  // drawing scaled down far enough wrote its arrow labels at 6.74pt next to
  // tile names held at 7.04 — the one piece of text on the slide that says
  // *why* two services are joined, and the only one small enough to be mush.
  // Dropping the chip is the designed answer: the workflow list on the slide
  // still carries the sentence against the same step number.
  const nodes: Node[] = [];
  const edges: Edge[] = [];
  const names = ['Azure Front Door', 'Application Gateway', 'Azure App Service', 'Azure Functions'];
  for (let i = 0; i < 27; i += 1) {
    nodes.push(service(`d-${i}`, names[i % names.length], i * 900, i * 620));
    if (i > 0) edges.push({ id: `h-${i}`, source: `d-${i - 1}`, target: `d-${i}`, label: 'Hands off' } as Edge);
  }
  const { parts } = await slideShapes({ nodes, edges });

  const chips = parts.flat().filter((s) => s.name.startsWith('connector-label-') && s.text.trim() !== '');
  assert.ok(chips.length > 0, 'the diagram drew no connector labels at all');
  const under = chips.filter((c) => (c.fontSize ?? 99) < 7);
  assert.equal(
    under.length, 0,
    `${under.length} chip(s) below the 7pt floor, smallest ${Math.min(...under.map((c) => c.fontSize ?? 99))}pt`,
  );
});

test('a far-placed outlier node still lands on the slide', async () => {
  const nodes: Node[] = [];
  for (let i = 0; i < 8; i += 1) nodes.push(service(`c-${i}`, 'Azure Functions', (i % 4) * 200, Math.floor(i / 4) * 160));
  nodes.push(service('outlier', 'Copilot Studio', 9000, 4000));

  const { shapes, slideW } = await slideShapes({ nodes, edges: [] });
  const presentationH = 7.5;
  const outlier = shapes.find((s) => s.name === 'service-outlier');
  assert.ok(outlier, 'the outlier tile must be emitted');
  assert.ok(outlier.x >= 0 && outlier.x + outlier.w <= slideW + 0.01, `outlier drawn off-slide at x=${outlier.x.toFixed(2)}in on a ${slideW.toFixed(2)}in page`);
  assert.ok(outlier.y >= 0, `outlier drawn above the slide at y=${outlier.y.toFixed(2)}in`);
  assert.ok(outlier.y + outlier.h <= Math.max(presentationH, outlier.y + outlier.h) + 0.01);

  for (const shape of shapes.filter((s) => /^service-[a-z0-9-]+$/.test(s.name))) {
    assert.ok(
      shape.x >= -0.01 && shape.x + shape.w <= slideW + 0.01,
      `${shape.name} is outside the page horizontally (x=${shape.x.toFixed(2)}, w=${shape.w.toFixed(2)}, page=${slideW.toFixed(2)})`,
    );
  }
});

test('the Visio page never exceeds what Visio can open', async () => {
  const { nodes, edges } = wideDiagram();
  const pkg = await buildVsdxPackage(nodes, edges, 'Contoso Platform');
  assert.ok(pkg.pageWidthIn <= 200 && pkg.pageHeightIn <= 200, `page ${pkg.pageWidthIn}x${pkg.pageHeightIn}in is too large for Visio`);
  assert.ok(pkg.pageWidthIn >= 8.5 && pkg.pageHeightIn >= 8.5);
});

test('Visio text stays readable and a stray node does not blow up the sheet', async () => {
  const { nodes, edges } = wideDiagram();
  const pkg = await buildVsdxPackage(nodes, edges, 'Contoso Platform');
  const page = pkg.parts.find((p) => /page1\.xml$/i.test(p.path));
  const xml = typeof page?.data === 'string' ? page.data : '';
  assert.ok(xml.length > 0, 'page1.xml must be emitted');

  // Visio font sizes are inches; anything under 5.5pt is unreadable in print.
  const sizes = [...xml.matchAll(/<Cell N="Size" V="([\d.]+)"\/>/g)].map((m) => +m[1] * 72);
  assert.ok(sizes.length > 0, 'the page must declare character sizes');
  const minPt = Math.min(...sizes);
  assert.ok(minPt >= 5.5, `smallest Visio font is ${minPt.toFixed(2)}pt`);

  // A far-placed node must not turn the drawing into a near-empty plotter sheet.
  const outlier = [...nodes, {
    id: 'stray',
    type: 'azureNode',
    position: { x: 9000, y: 4000 },
    width: 150,
    height: 75,
    data: { label: 'Copilot Studio', serviceName: 'Copilot Studio' },
  } as Node];
  const strayPkg = await buildVsdxPackage(outlier, edges, 'Contoso Platform');
  assert.ok(
    strayPkg.pageWidthIn < 90 && strayPkg.pageHeightIn < 45,
    `a single stray node grew the page to ${strayPkg.pageWidthIn.toFixed(1)}x${strayPkg.pageHeightIn.toFixed(1)}in`,
  );

  // ...and it must still be on the page, not silently discarded.
  const strayPage = strayPkg.parts.find((p) => /page1\.xml$/i.test(p.path));
  const strayXml = typeof strayPage?.data === 'string' ? strayPage.data : '';
  const groups = [...strayXml.matchAll(
    /NameU="Service\.\d+"[\s\S]*?<Cell N="PinX" V="([\d.-]+)"\/>\s*<Cell N="PinY" V="([\d.-]+)"\/>\s*<Cell N="Width" V="([\d.-]+)"\/>\s*<Cell N="Height" V="([\d.-]+)"\/>/g,
  )];
  assert.equal(groups.length, outlier.filter((n) => n.type !== 'groupNode').length, 'every service must be emitted');
  for (const [, pinX, pinY, w, h] of groups) {
    const left = +pinX - +w / 2;
    const bottom = +pinY - +h / 2;
    assert.ok(left >= -0.01 && left + +w <= strayPkg.pageWidthIn + 0.01, `service off-page horizontally at ${left.toFixed(2)}in`);
    assert.ok(bottom >= -0.01 && bottom + +h <= strayPkg.pageHeightIn + 0.01, `service off-page vertically at ${bottom.toFixed(2)}in`);
  }
});
