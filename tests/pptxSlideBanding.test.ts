import test from 'node:test';
import assert from 'node:assert/strict';
import JSZip from 'jszip';
import type { Edge, Node } from 'reactflow';
import { buildDiagramSlidePptx } from '../src/services/pptxExporter.ts';

/**
 * An architecture that is too wide for one readable PowerPoint page used to be
 * squeezed onto a single 56" slide, which drove the tile labels down to under
 * 5pt — legally "exported" but unusable without redrawing the deck by hand.
 * The Azure Architecture Center convention is to keep the drawing at a readable
 * size and continue it across pages, so these tests pin the split.
 */

const PIXEL_PNG =
  'data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg==';

function service(id: string, label: string, x: number, y: number): Node {
  return {
    id,
    type: 'azureNode',
    position: { x, y },
    width: 150,
    height: 75,
    data: { label, serviceName: label },
  } as Node;
}

interface Deck {
  slides: string[];
}

async function buildDeck(nodes: Node[], edges: Edge[]): Promise<Deck> {
  const pptx = await buildDiagramSlidePptx(PIXEL_PNG, {
    diagramName: 'Contoso Platform',
    author: 'Tester',
    date: '2026-08-10',
    isDarkMode: false,
    diagram: { nodes, edges },
  });
  const buffer = (await pptx.write({ outputType: 'nodebuffer' })) as Buffer;
  const zip = await JSZip.loadAsync(buffer);
  const names = Object.keys(zip.files)
    .filter((name) => /^ppt\/slides\/slide\d+\.xml$/.test(name))
    .sort((a, b) => +a.replace(/\D/g, '') - +b.replace(/\D/g, ''));
  const slides = await Promise.all(names.map((name) => zip.file(name)!.async('string')));
  return { slides };
}

/** Font sizes carried by the service label text runs, in points. */
function labelFontSizes(slideXml: string): number[] {
  const sizes: number[] = [];
  const shapeRe = /<p:sp>[\s\S]*?<\/p:sp>/g;
  let match: RegExpExecArray | null;
  while ((match = shapeRe.exec(slideXml)) !== null) {
    if (!/name="service-label-/.test(match[0])) continue;
    for (const size of match[0].matchAll(/sz="(\d+)"/g)) sizes.push(+size[1] / 100);
  }
  return sizes;
}

function wideDiagram(count: number): { nodes: Node[]; edges: Edge[] } {
  const nodes: Node[] = [];
  const edges: Edge[] = [];
  for (let i = 0; i < count; i += 1) {
    nodes.push(service(`n-${i}`, i % 2 ? 'Azure Kubernetes Service' : 'Copilot Studio', i * 260, (i % 4) * 220));
    if (i > 0) edges.push({ id: `x-${i}`, source: `n-${i - 1}`, target: `n-${i}`, label: 'Managed identity' } as Edge);
  }
  return { nodes, edges };
}

test('a diagram that fits stays on exactly one slide', async () => {
  const nodes = [service('a', 'API Management', 0, 0), service('b', 'Azure Functions', 320, 0)];
  const deck = await buildDeck(nodes, [{ id: 'e', source: 'a', target: 'b', label: 'Invoke' } as Edge]);
  assert.equal(deck.slides.length, 1, 'a small architecture must not be split');
  assert.ok(!deck.slides[0].includes('(1 / 1)'), 'a single-slide export carries no part marker');
});

test('an architecture too wide to stay readable is continued across slides', async () => {
  const { nodes, edges } = wideDiagram(40);
  const deck = await buildDeck(nodes, edges);
  assert.ok(deck.slides.length > 1, `expected a multi-slide split, got ${deck.slides.length}`);
  deck.slides.forEach((xml, index) => {
    assert.ok(
      xml.includes(`(${index + 1} / ${deck.slides.length})`),
      `slide ${index + 1} must say which part of the drawing it shows`,
    );
  });
});

test('every tile label on every slide stays above the 7pt legibility floor', async () => {
  const { nodes, edges } = wideDiagram(40);
  const deck = await buildDeck(nodes, edges);
  const sizes = deck.slides.flatMap(labelFontSizes);
  assert.ok(sizes.length > 0, 'the deck must contain service labels');
  const smallest = Math.min(...sizes);
  assert.ok(smallest >= 7, `smallest label font is ${smallest}pt, below the 7pt legibility floor`);
});

test('splitting loses no service and duplicates none', async () => {
  const { nodes, edges } = wideDiagram(40);
  const deck = await buildDeck(nodes, edges);
  const drawn = new Map<string, number>();
  for (const xml of deck.slides) {
    for (const match of xml.matchAll(/name="service-(n-\d+)"/g)) {
      drawn.set(match[1], (drawn.get(match[1]) ?? 0) + 1);
    }
  }
  assert.equal(drawn.size, nodes.length, `${drawn.size} of ${nodes.length} services were drawn`);
  const repeated = [...drawn.entries()].filter(([, times]) => times !== 1);
  assert.deepEqual(repeated, [], 'a service must be drawn on exactly one slide');
});

test('a split deck keeps every shape inside the page', async () => {
  const { nodes, edges } = wideDiagram(40);
  const pptx = await buildDiagramSlidePptx(PIXEL_PNG, {
    diagramName: 'Contoso Platform',
    author: 'Tester',
    date: '2026-08-10',
    isDarkMode: false,
    diagram: { nodes, edges },
  });
  const buffer = (await pptx.write({ outputType: 'nodebuffer' })) as Buffer;
  const zip = await JSZip.loadAsync(buffer);
  const presentation = await zip.file('ppt/presentation.xml')!.async('string');
  const size = /<p:sldSz[^>]*cx="(\d+)"[^>]*cy="(\d+)"/.exec(presentation);
  assert.ok(size, 'the package must declare a slide size');
  const EMU = 914400;
  const pageW = +size[1] / EMU;
  const pageH = +size[2] / EMU;
  const names = Object.keys(zip.files).filter((n) => /^ppt\/slides\/slide\d+\.xml$/.test(n));
  for (const name of names) {
    const xml = await zip.file(name)!.async('string');
    for (const off of xml.matchAll(/<a:off x="(-?\d+)" y="(-?\d+)"\/><a:ext cx="(\d+)" cy="(\d+)"\/>/g)) {
      const x = +off[1] / EMU;
      const y = +off[2] / EMU;
      assert.ok(x >= -0.01 && y >= -0.01, `${name}: shape starts off-page at (${x}, ${y})`);
      assert.ok(
        x + +off[3] / EMU <= pageW + 0.01 && y + +off[4] / EMU <= pageH + 0.01,
        `${name}: shape runs past the ${pageW}x${pageH}in page`,
      );
    }
  }
});
