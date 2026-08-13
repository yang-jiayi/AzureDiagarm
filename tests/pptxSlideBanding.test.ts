import test from 'node:test';
import assert from 'node:assert/strict';
import JSZip from 'jszip';
import type { Edge, Node } from 'reactflow';
import { buildDiagramSlidePptx, fitTableRows, legibleScaleFor, tableRowHeightIn, wrappedLineCount } from '../src/services/pptxExporter.ts';
import { buildExportRoutes, collectExportBoxes } from '../src/services/diagramExportGeometry.ts';
import { buildVsdxPackage } from '../src/services/visioVsdxExporter.ts';

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
  /**
   * A tiled deck opens with the whole drawing shown small, so the legible
   * parts are the slides after it. Rules about legibility and about a shape
   * belonging to exactly one window apply to those, not to the overview.
   */
  parts: string[];
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
  const overview = slides.findIndex((xml) => xml.includes('(Overview)'));
  return { slides, parts: overview < 0 ? slides : slides.slice(overview + 1) };
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
  assert.ok(deck.parts.length > 1, `expected a multi-slide split, got ${deck.parts.length}`);
  assert.ok(deck.slides[0].includes('(Overview)'), 'a tiled deck must open with the whole drawing');
  deck.parts.forEach((xml, index) => {
    assert.ok(
      xml.includes(`(${index + 1} / ${deck.parts.length})`),
      `slide ${index + 1} must say which part of the drawing it shows`,
    );
  });
});

test('every tile label on every slide stays above the 7pt legibility floor', async () => {
  const { nodes, edges } = wideDiagram(40);
  const deck = await buildDeck(nodes, edges);
  const sizes = deck.parts.flatMap(labelFontSizes);
  assert.ok(sizes.length > 0, 'the deck must contain service labels');
  const smallest = Math.min(...sizes);
  assert.ok(smallest >= 7, `smallest label font is ${smallest}pt, below the 7pt legibility floor`);
});

test('splitting loses no service and duplicates none', async () => {
  const { nodes, edges } = wideDiagram(40);
  const deck = await buildDeck(nodes, edges);
  const drawn = new Map<string, number>();
  for (const xml of deck.parts) {
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

test('the numbered badges are paired with a numbered workflow list', async () => {
  const nodes = [service('a', 'API Management', 0, 0), service('b', 'Azure Functions', 320, 0), service('c', 'Azure SQL Database', 640, 0)];
  const edges = [
    { id: 'e1', source: 'a', target: 'b', data: { stepNumber: 1, stepDescription: 'API Management がリクエストを受け取ります' } },
    { id: 'e2', source: 'b', target: 'c', data: { stepNumber: 2, stepDescription: 'Functions が SQL Database に書き込みます' } },
  ] as Edge[];
  const deck = await buildDeck(nodes, edges);
  const all = deck.slides.join('');
  for (const step of [1, 2]) {
    assert.ok(all.includes(`name="workflow-step-${step}"`), `step ${step} is missing its numbered badge in the list`);
    assert.ok(all.includes(`name="workflow-text-${step}"`), `step ${step} is missing its narration row`);
  }
  assert.ok(all.includes('API Management がリクエストを受け取ります'), 'the step prose must be carried into the deck');
});

test('an unnumbered diagram gets no workflow list', async () => {
  const nodes = [service('a', 'API Management', 0, 0), service('b', 'Azure Functions', 320, 0)];
  const deck = await buildDeck(nodes, [{ id: 'e', source: 'a', target: 'b', label: 'Invoke' } as Edge]);
  assert.equal(deck.slides.length, 1, 'no workflow means no extra slide');
  assert.ok(!deck.slides[0].includes('workflow-text-'), 'no workflow rows without workflow data');
});

test('the Visio sheet carries the same numbered workflow narration', async () => {
  const nodes = [service('a', 'API Management', 0, 0), service('b', 'Azure Functions', 320, 0)];
  const edges = [
    { id: 'e1', source: 'a', target: 'b', data: { stepNumber: 1, stepDescription: 'Front Door がトラフィックを転送します' } },
  ] as Edge[];
  const pkg = await buildVsdxPackage(nodes, edges, 'Contoso Platform');
  const page = pkg.parts.find((p) => /page1\.xml$/i.test(p.path));
  assert.ok(page, 'the package must contain page1.xml');
  const xml = String(page.data);
  assert.ok(/NameU="Workflow\.\d+"/.test(xml), 'the workflow panel shape is missing');
  assert.ok(xml.includes('Front Door がトラフィックを転送します'), 'the step prose must reach the Visio sheet');
  assert.ok(xml.includes('&gt;Workflow&lt;') || xml.includes('>Workflow<'), 'the panel must be titled');
});

test('a Visio sheet without a workflow gets no panel', async () => {
  const nodes = [service('a', 'API Management', 0, 0), service('b', 'Azure Functions', 320, 0)];
  const pkg = await buildVsdxPackage(nodes, [{ id: 'e', source: 'a', target: 'b' } as Edge], 'Contoso');
  const xml = String(pkg.parts.find((p) => /page1\.xml$/i.test(p.path))!.data);
  assert.ok(!/NameU="Workflow\.\d+"/.test(xml), 'no workflow data must mean no panel');
});

/**
 * The four defects the second review found, all in the banding and workflow
 * code. Each one produced a deck that looked fine and was quietly wrong: a
 * missing service, a service drawn twice, a tile hidden under an opaque panel,
 * and a badge on the drawing whose sentence was nowhere in the deck.
 */

test('a stray service outside the trimmed bounds is still drawn on a banded deck', async () => {
  const { nodes, edges } = wideDiagram(40);
  nodes.push(service('stray', 'Copilot Studio', 100000, 40000));
  const deck = await buildDeck(nodes, edges);
  assert.ok(deck.parts.length > 1, 'this diagram must band');
  const drawn = deck.parts.filter((xml) => xml.includes('name="service-stray"')).length;
  assert.equal(drawn, 1, 'the outlier must be clamped onto exactly one band, not dropped');
});

test('a service straddling a band seam is drawn once, not on both slides', async () => {
  const { nodes, edges } = wideDiagram(40);
  // 40 tiles at a 260px pitch put the midpoint seam near x = 5145.
  nodes.push(service('straddler', 'Azure SQL Database', 5100, 700));
  const deck = await buildDeck(nodes, edges);
  assert.ok(deck.parts.length > 1, 'this diagram must band');
  const drawn = deck.parts.filter((xml) => xml.includes('name="service-straddler"')).length;
  assert.equal(drawn, 1, 'a seam-crossing service must belong to exactly one band');
});

test('a numbered arrow crossing a seam carries exactly one badge and one chip', async () => {
  const nodes: Node[] = [];
  const edges: Edge[] = [];
  for (let i = 0; i < 40; i += 1) {
    nodes.push(service(`n-${i}`, i % 2 ? 'Azure Kubernetes Service' : 'Copilot Studio', i * 260, (i % 4) * 220));
    if (i > 0) {
      edges.push({
        id: `x-${i}`,
        source: `n-${i - 1}`,
        target: `n-${i}`,
        label: 'Managed identity',
        ...(i <= 24 ? { data: { stepNumber: i, stepDescription: `Step ${i}` } } : {}),
      } as Edge);
    }
  }
  const deck = await buildDeck(nodes, edges);
  assert.ok(deck.parts.length > 1, 'this diagram must band');

  const count = (needle: string) => deck.parts.filter((xml) => xml.includes(needle)).length;
  for (let i = 1; i <= 24; i += 1) {
    assert.equal(count(`name="connector-step-x-${i}"`), 1, `badge ${i} must appear once`);
    assert.equal(count(`name="connector-label-x-${i}"`), 1, `chip for step ${i} must appear once`);
  }
});

test('a workflow too long for one slide is continued, never truncated', async () => {
  const nodes: Node[] = [];
  const edges: Edge[] = [];
  for (let i = 0; i < 21; i += 1) {
    nodes.push(service(`w-${i}`, 'Azure Functions', (i % 7) * 300, Math.floor(i / 7) * 240));
    if (i > 0) {
      edges.push({
        id: `w-${i}`,
        source: `w-${i - 1}`,
        target: `w-${i}`,
        label: 'Private Link',
        data: { stepNumber: i, stepDescription: `Step ${i} moves the message onward` },
      } as Edge);
    }
  }
  const deck = await buildDeck(nodes, edges);
  for (let step = 1; step <= 20; step += 1) {
    const rows = deck.slides.filter((xml) => xml.includes(`name="workflow-text-${step}"`)).length;
    assert.equal(rows, 1, `step ${step} must have exactly one row somewhere in the deck`);
  }
});

test('the Visio workflow panel never covers a clamped stray service', async () => {
  const nodes: Node[] = [];
  for (let i = 0; i < 8; i += 1) {
    nodes.push(service(`c-${i}`, 'Azure Functions', (i % 4) * 220, Math.floor(i / 4) * 180));
  }
  // Above and to the left of the cluster: the clamp used to pull it into the
  // band the opaque workflow panel occupies.
  nodes.push(service('stray', 'Copilot Studio', -9000, -4000));
  const edges: Edge[] = [
    { id: 'e-out', source: 'c-0', target: 'stray', label: 'Reach the stray', data: { stepNumber: 1, stepDescription: 'Connect to the far service' } } as Edge,
    { id: 'e-in', source: 'c-1', target: 'c-2', label: 'Managed identity', data: { stepNumber: 2, stepDescription: 'Authenticate with a managed identity' } } as Edge,
  ];

  const pkg = await buildVsdxPackage(nodes, edges, 'Contoso Platform');
  const xml = String(pkg.parts.find((p) => /page1\.xml$/i.test(p.path))!.data);

  const boxes = (namePattern: string) => {
    const re = new RegExp(
      `NameU="(${namePattern})"[\\s\\S]*?<Cell N="PinX" V="([\\d.-]+)"\\/>\\s*<Cell N="PinY" V="([\\d.-]+)"\\/>\\s*<Cell N="Width" V="([\\d.-]+)"\\/>\\s*<Cell N="Height" V="([\\d.-]+)"\\/>`,
      'g',
    );
    return [...xml.matchAll(re)].map((m) => ({
      name: m[1], x: +m[2] - +m[4] / 2, y: +m[3] - +m[5] / 2, w: +m[4], h: +m[5],
    }));
  };

  const panels = boxes('Workflow\\.\\d+');
  assert.equal(panels.length, 1, 'the sheet must carry exactly one workflow panel');
  const panel = panels[0];
  const tiles = boxes('Service\\.\\d+');
  assert.ok(tiles.length > 0, 'the sheet must carry service tiles');
  for (const tile of tiles) {
    const ox = Math.min(tile.x + tile.w, panel.x + panel.w) - Math.max(tile.x, panel.x);
    const oy = Math.min(tile.y + tile.h, panel.y + panel.h) - Math.max(tile.y, panel.y);
    const overlap = Math.max(0, ox) * Math.max(0, oy);
    assert.ok(
      overlap <= 0.02 * tile.w * tile.h,
      `${tile.name} is ${((overlap / (tile.w * tile.h)) * 100).toFixed(0)}% hidden behind the workflow panel`,
    );
  }
});

test('a generated linear flow lands on a page a document can actually use', async () => {
  const { applyLayoutPreset } = await import('../src/utils/layoutPresets.ts');
  const names = [
    'Azure Front Door', 'Application Gateway', 'Azure Kubernetes Service', 'Azure Service Bus',
    'Azure Functions', 'Azure Cosmos DB', 'Azure Data Factory', 'Azure Synapse Analytics',
    'Azure OpenAI Service', 'Azure AI Search', 'Key Vault', 'Azure Monitor',
  ];
  const nodes = names.map((name, i) => service(`g-${i}`, name, 0, 0));
  const edges = names.slice(1).map((_, i) => ({
    id: `g-e-${i}`, source: `g-${i}`, target: `g-${i + 1}`, label: 'Invoke',
  } as Edge));

  const laidOut = await applyLayoutPreset(nodes, edges, {
    preset: 'flow-lr', spacing: 'comfortable', edgeStyle: 'smooth', emphasizePrimaryPath: false,
  });

  const pptx = await buildDiagramSlidePptx(PIXEL_PNG, {
    diagramName: 'Contoso Platform',
    author: 'Tester',
    date: '2026-08-10',
    isDarkMode: false,
    diagram: { nodes: laidOut.nodes, edges: laidOut.edges },
  });
  const buffer = (await pptx.write({ outputType: 'nodebuffer' })) as Buffer;
  const zip = await JSZip.loadAsync(buffer);
  const presentation = await zip.file('ppt/presentation.xml')!.async('string');
  const size = /<p:sldSz[^>]*cx="(\d+)"[^>]*cy="(\d+)"/.exec(presentation)!;
  const pageW = +size[1] / 914400;
  const pageH = +size[2] / 914400;

  // Unwrapped, dagre returns one rank per service and the same twelve services
  // need a 42.6in page — over three standard slides side by side, printed
  // across four sheets.
  assert.ok(
    pageW <= 13.333 * 2,
    `a twelve-service flow needs a ${pageW.toFixed(1)}in page — it was not folded`,
  );
  assert.ok(
    pageW / pageH <= 2.6,
    `page is ${(pageW / pageH).toFixed(2)}:1, not a shape a document can use`,
  );
});

test('a zone that spans a band seam keeps its boundary on every slide', async () => {
  const nodes: Node[] = [
    { id: 'zone1', type: 'groupNode', position: { x: -60, y: -60 }, style: { width: 40 * 260 + 120, height: 400 }, data: { label: 'Production VNet' } } as Node,
  ];
  for (let i = 0; i < 40; i += 1) {
    nodes.push({
      id: `z-${i}`,
      type: 'azureNode',
      parentNode: 'zone1',
      position: { x: 60 + i * 260, y: 60 },
      width: 150,
      height: 75,
      data: { label: 'Azure Functions', serviceName: 'Azure Functions' },
    } as Node);
  }

  const pptx = await buildDiagramSlidePptx(PIXEL_PNG, {
    diagramName: 'Contoso Platform',
    author: 'Tester',
    date: '2026-08-10',
    isDarkMode: false,
    diagram: { nodes, edges: [] },
  });
  const buffer = (await pptx.write({ outputType: 'nodebuffer' })) as Buffer;
  const zip = await JSZip.loadAsync(buffer);
  const slides = Object.keys(zip.files)
    .filter((n) => /^ppt\/slides\/slide\d+\.xml$/.test(n))
    .sort((a, b) => (+a.replace(/\D/g, '')) - (+b.replace(/\D/g, '')));

  const withServices: string[] = [];
  const withZone: string[] = [];
  for (const name of slides) {
    const xml = await zip.file(name)!.async('string');
    if (/name="service-/.test(xml)) withServices.push(name);
    if (/name="(group|zone)-?[^"]*zone1"/.test(xml) || /Production VNet/.test(xml)) withZone.push(name);
  }

  assert.ok(withServices.length > 1, 'the fixture no longer bands across slides');
  // A zone is wider than a whole band, so centre-ownership would print the
  // boundary on one slide and leave the services on the others with no
  // container and no name.
  assert.deepEqual(
    withServices.filter((s) => !withZone.includes(s)),
    [],
    `slides show services with no zone boundary: ${withServices.filter((s) => !withZone.includes(s)).join(', ')}`,
  );
});

test('a zone wider than a band is cut at the page edge, not slid across it', async () => {
  const nodes: Node[] = [
    { id: 'zoneA', type: 'groupNode', position: { x: -60, y: -60 }, style: { width: 40 * 260 + 120, height: 400 }, data: { label: 'Production VNet' } } as Node,
  ];
  for (let i = 0; i < 40; i += 1) {
    nodes.push({
      id: `zz-${i}`,
      type: 'azureNode',
      parentNode: 'zoneA',
      position: { x: 60 + i * 260, y: 60 },
      width: 150,
      height: 75,
      data: { label: 'Azure Functions', serviceName: 'Azure Functions' },
    } as Node);
  }

  const pptx = await buildDiagramSlidePptx(PIXEL_PNG, {
    diagramName: 'Contoso Platform',
    author: 'Tester',
    date: '2026-08-10',
    isDarkMode: false,
    diagram: { nodes, edges: [] },
  });
  const buffer = (await pptx.write({ outputType: 'nodebuffer' })) as Buffer;
  const zip = await JSZip.loadAsync(buffer);
  const slides = Object.keys(zip.files)
    .filter((n) => /^ppt\/slides\/slide\d+\.xml$/.test(n))
    .sort((a, b) => (+a.replace(/\D/g, '')) - (+b.replace(/\D/g, '')));

  const EMU = 914400;
  let checked = 0;
  for (const name of slides) {
    const xml = await zip.file(name)!.async('string');
    const page = /<p:sldSz[^>]*cx="(\d+)"/.exec(await zip.file('ppt/presentation.xml')!.async('string'));
    const pageW = page ? +page[1] / EMU : 13.333;
    for (const m of xml.matchAll(/name="zone-zoneA"[\s\S]{0,400}?<a:off x="(-?[\d.]+)" y="(-?[\d.]+)"\s*\/>\s*<a:ext cx="([\d.]+)" cy="([\d.]+)"\s*\/>/g)) {
      const x = +m[1] / EMU;
      const w = +m[3] / EMU;
      checked += 1;
      // A zone clamped by its origin alone keeps its full width, so it is
      // painted right across the band and boxes in tiles that are not inside
      // it -- and PowerPoint's writer, handed a width larger than a slide,
      // emits the raw inch count as EMU and the boundary vanishes entirely.
      assert.ok(x + w <= pageW + 0.01, `zone runs ${(x + w).toFixed(2)}in past a ${pageW.toFixed(2)}in page on ${name}`);
      assert.ok(w >= pageW * 0.5, `zone spanning every tile shrank to ${w.toFixed(3)}in on a ${pageW.toFixed(2)}in ${name}`);
    }
  }
  assert.ok(checked > 0, 'no zone rectangle was found to check');
});

test('a trimmed outlier zone is clamped back onto the page, not cut to a hairline', async () => {
  // `chooseExportBounds` trims a far outlier out of the drawing bounds, so its
  // zone lands entirely outside the frame. Cutting a box that misses the frame
  // leaves a hairline at an off-page coordinate -- and pptxgenjs writes any
  // number above 100 as a raw EMU count, so the reader gets a service tile
  // with no container at all.
  const nodes: Node[] = [];
  for (let i = 0; i < 20; i += 1) {
    nodes.push({
      id: `near-${i}`, type: 'azureNode', position: { x: (i % 5) * 220, y: Math.floor(i / 5) * 140 },
      width: 150, height: 75, data: { label: 'Azure Functions', serviceName: 'Azure Functions' },
    } as Node);
  }
  nodes.push({
    id: 'zFar', type: 'groupNode', position: { x: 20000, y: 9000 },
    style: { width: 600, height: 400 }, data: { label: 'Remote Region' },
  } as Node);
  nodes.push({
    id: 'sFar', type: 'azureNode', parentNode: 'zFar', position: { x: 100, y: 120 },
    width: 150, height: 75, data: { label: 'Azure Functions', serviceName: 'Azure Functions' },
  } as Node);

  const pptx = await buildDiagramSlidePptx(PIXEL_PNG, {
    diagramName: 'Contoso Platform', author: 'Tester', date: '2026-08-10',
    isDarkMode: false, diagram: { nodes, edges: [] },
  });
  const zip = await JSZip.loadAsync((await pptx.write({ outputType: 'nodebuffer' })) as Buffer);
  const EMU = 914400;
  const pres = await zip.file('ppt/presentation.xml')!.async('string');
  const pageW = +/<p:sldSz[^>]*cx="(\d+)"/.exec(pres)![1] / EMU;
  const pageH = +/<p:sldSz[^>]*cy="(\d+)"/.exec(pres)![1] / EMU;

  let seen = 0;
  let onWindows = 0;
  for (const name of Object.keys(zip.files).filter((n) => /^ppt\/slides\/slide\d+\.xml$/.test(n))) {
    const xml = await zip.file(name)!.async('string');
    const at = xml.indexOf('name="zone-zFar"');
    if (at < 0) continue;
    const m = /<a:off x="(-?[\d.]+)" y="(-?[\d.]+)"\s*\/>\s*<a:ext cx="([\d.]+)" cy="([\d.]+)"/.exec(xml.slice(at, at + 400))!;
    const [x, y, w, h] = [+m[1] / EMU, +m[2] / EMU, +m[3] / EMU, +m[4] / EMU];
    seen += 1;
    if (!xml.includes('(Overview)')) onWindows += 1;
    assert.ok(w > 0.2 && h > 0.2, `trimmed zone collapsed to ${w.toFixed(3)}x${h.toFixed(3)}in on ${name}`);
    assert.ok(
      x >= -0.01 && y >= -0.01 && x + w <= pageW + 0.01 && y + h <= pageH + 0.01,
      `trimmed zone sits at ${x.toFixed(2)},${y.toFixed(2)} ${w.toFixed(2)}x${h.toFixed(2)}in off a ${pageW.toFixed(2)}x${pageH.toFixed(2)}in ${name}`,
    );
  }
  // Bringing a 6.25in remote zone back beside a 10.7in drawing makes a 17.6in
  // drawing, which cannot be shown on one standard slide above 6.7pt — so this
  // deck is tiled, and a tiled deck opens with an overview that redraws
  // everything. What must hold is the audit's own rule: exactly one *window*
  // carries the zone, so it is neither dropped nor drawn twice.
  assert.ok(seen >= 1, 'the outlier zone was not drawn at all');
  assert.equal(onWindows, 1, 'the outlier zone must belong to exactly one window');
});

/** Slide-space rectangles for every shape whose objectName matches a prefix. */
function shapeBoxes(slideXml: string, prefix: string): { name: string; x: number; y: number; w: number; h: number }[] {
  const boxes: { name: string; x: number; y: number; w: number; h: number }[] = [];
  const EMU = 914400;
  for (const match of slideXml.matchAll(/<p:sp>[\s\S]*?<\/p:sp>/g)) {
    const name = /name="([^"]*)"/.exec(match[0])?.[1] ?? '';
    if (!name.startsWith(prefix)) continue;
    const off = /<a:off x="(-?\d+)" y="(-?\d+)"\/>/.exec(match[0]);
    const ext = /<a:ext cx="(-?[\d.]+)" cy="(-?[\d.]+)"\/>/.exec(match[0]);
    if (!off || !ext) continue;
    boxes.push({ name, x: +off[1] / EMU, y: +off[2] / EMU, w: +ext[1] / EMU, h: +ext[2] / EMU });
  }
  return boxes;
}

test('a drawing large in both directions is tiled onto ordinary slides, not a plotter sheet', async () => {
  // Wide *and* tall: splitting on one axis alone left the page growing without
  // limit, which is how a generated architecture ended up as a 27 x 16in sheet
  // that PowerPoint will open but nobody can present.
  const nodes: Node[] = [];
  const edges: Edge[] = [];
  for (let row = 0; row < 8; row += 1) {
    for (let col = 0; col < 9; col += 1) {
      const id = `g-${row}-${col}`;
      nodes.push(service(id, 'Azure Kubernetes Service', col * 300, row * 240));
      if (col > 0) edges.push({ id: `e-${row}-${col}`, source: `g-${row}-${col - 1}`, target: id } as Edge);
    }
  }
  const deck = await buildDeck(nodes, edges);
  const pageIn = /<p:sldSz cx="(\d+)" cy="(\d+)"/.exec(
    await (async () => {
      const pptx = await buildDiagramSlidePptx(PIXEL_PNG, {
        diagramName: 'Contoso Platform',
        author: 'Tester',
        date: '2026-08-10',
        isDarkMode: false,
        diagram: { nodes, edges },
      });
      const buffer = (await pptx.write({ outputType: 'nodebuffer' })) as Buffer;
      const zip = await JSZip.loadAsync(buffer);
      return zip.file('ppt/presentation.xml')!.async('string');
    })(),
  );
  assert.ok(pageIn, 'the deck must declare a slide size');
  const pageW = +pageIn[1] / 914400;
  const pageH = +pageIn[2] / 914400;
  assert.ok(
    Math.abs(pageW - 13.333) < 0.05 && Math.abs(pageH - 7.5) < 0.05,
    `expected standard 13.33x7.5in slides, got ${pageW.toFixed(2)}x${pageH.toFixed(2)}in`,
  );

  // Tiling in one axis only would put every part in a single row or column, so
  // check the windows really form a grid: at least two parts must repeat the
  // same horizontal slice of the drawing at different heights.
  const columnKeys = deck.parts.map((xml) => {
    const tiles = shapeBoxes(xml, 'service-').filter((b) => !b.name.includes('label') && !b.name.includes('meta'));
    return tiles.map((t) => t.name.replace(/^service-g-\d+-/, '')).sort().join(',');
  });
  const repeated = columnKeys.filter((key, index) => key && columnKeys.indexOf(key) !== index);
  assert.ok(repeated.length > 0, `expected a 2-D grid of windows, got column sets ${JSON.stringify(columnKeys)}`);

  const drawn = new Map<string, number>();
  for (const xml of deck.parts) {
    for (const tile of shapeBoxes(xml, 'service-')) {
      if (tile.name.includes('label') || tile.name.includes('meta')) continue;
      drawn.set(tile.name, (drawn.get(tile.name) ?? 0) + 1);
    }
  }
  const wrong = nodes.filter((n) => drawn.get(`service-${n.id}`) !== 1);
  assert.equal(wrong.length, 0, `services drawn on the wrong number of slides: ${wrong.slice(0, 4).map((n) => `${n.id}=${drawn.get(`service-${n.id}`) ?? 0}`).join(', ')}`);
});

test('a label chip is pushed clear of the services it sits between', async () => {
  // Two tiles a hair apart with a long label: centred on the arrow the chip
  // lands squarely on both of them, which is the "export needs redrawing by
  // hand" complaint in its smallest form.
  const nodes = [service('a', 'API Management', 0, 0), service('b', 'Azure Functions', 170, 0)];
  const edges = [{
    id: 'e',
    source: 'a',
    target: 'b',
    label: 'HTTPS 経由でトークン検証を実施',
  } as Edge];
  const deck = await buildDeck(nodes, edges);
  const xml = deck.parts[0] ?? deck.slides[0];
  const tiles = shapeBoxes(xml, 'service-').filter((b) => !b.name.includes('label') && !b.name.includes('meta'));
  const chips = shapeBoxes(xml, 'connector-label-');
  assert.equal(chips.length, 1, 'the labelled arrow must carry exactly one chip');
  assert.ok(tiles.length >= 2, 'both services must be drawn');
  for (const tile of tiles) {
    const dx = Math.min(chips[0].x + chips[0].w, tile.x + tile.w) - Math.max(chips[0].x, tile.x);
    const dy = Math.min(chips[0].y + chips[0].h, tile.y + tile.h) - Math.max(chips[0].y, tile.y);
    const area = dx > 0 && dy > 0 ? dx * dy : 0;
    assert.ok(
      area <= 0.02 * tile.w * tile.h,
      `chip covers ${((area / (tile.w * tile.h)) * 100).toFixed(0)}% of ${tile.name}`,
    );
  }
});

test('parallel connector labels and callouts stay on the slide', () => Promise.resolve().then(async () => {
  // Ten edges between the same close-together pair: each label sat on the
  // one below, and the cap that kept a chip inside the gap between the tiles
  // forced a long label into a tall narrow ribbon. The stagger then walked
  // that ribbon straight off the bottom of the page, taking callout 2 with it.
  // Three edges is one short of the failure — the routes are already fanned
  // apart by a fraction of a rung, so adding the stagger on top of each
  // route's own anchor puts the chips off the lattice and half inside each
  // other, which only shows up once the ladder is four or more rungs deep.
  const nodes = [service('a', 'Azure Front Door', 0, 0), service('b', 'Azure Kubernetes Service', 190, 0)];
  const label = 'アプリケーション ゲートウェイ経由の HTTPS トラフィック';
  const edges = Array.from({ length: 10 }, (_, i) => `p${i + 1}`).map((id, i) => ({
    id,
    source: 'a',
    target: 'b',
    label: `${label} ${i + 1}`,
    data: { stepNumber: i + 1 },
  })) as Edge[];

  const deck = await buildDeck(nodes, edges);
  const page = { w: 13.333, h: 7.5 };
  const offenders: string[] = [];
  for (const slide of deck.slides) {
    for (const prefix of ['connector-label-', 'connector-step-']) {
      for (const box of shapeBoxes(slide, prefix)) {
        if (box.x < -0.01 || box.y < -0.01 || box.x + box.w > page.w + 0.01 || box.y + box.h > page.h + 0.01) {
          offenders.push(`${box.name} at ${box.x.toFixed(2)},${box.y.toFixed(2)} ${box.w.toFixed(2)}x${box.h.toFixed(2)}in`);
        }
      }
    }
  }
  assert.deepEqual(offenders, [], `off the page: ${offenders.join('; ')}`);

  const badges = deck.slides.flatMap((slide) => shapeBoxes(slide, 'connector-step-').map((b) => b.name));
  assert.equal(new Set(badges).size, 10, `expected all 10 callouts, got ${[...new Set(badges)].join(', ')}`);

  // A chip squeezed into the hop between two close tiles became a narrow
  // ribbon several inches tall, which is unreadable and drives the stagger.
  const tall = deck.slides
    .flatMap((slide) => shapeBoxes(slide, 'connector-label-'))
    .filter((b) => b.h > b.w);
  assert.deepEqual(tall.map((b) => `${b.name} ${b.w.toFixed(2)}x${b.h.toFixed(2)}in`), [],
    'a label chip was squeezed into a vertical ribbon');

  // Parallel edges are staggered, but the walk that moves a chip off a tile
  // could drag two of them back onto the same spot, because a chip could not
  // see the chips already placed.
  const marks = deck.slides.flatMap((slide) => [
    ...shapeBoxes(slide, 'connector-label-'),
    ...shapeBoxes(slide, 'connector-step-'),
  ]);
  const collisions: string[] = [];
  for (let i = 0; i < marks.length; i += 1) {
    for (let j = i + 1; j < marks.length; j += 1) {
      const a = marks[i];
      const b = marks[j];
      const dx = Math.min(a.x + a.w, b.x + b.w) - Math.max(a.x, b.x);
      const dy = Math.min(a.y + a.h, b.y + b.h) - Math.max(a.y, b.y);
      if (dx > 0.01 && dy > 0.01) collisions.push(`${a.name}/${b.name}`);
    }
  }
  assert.deepEqual(collisions, [], `annotations sit on each other: ${collisions.join(' ')}`);
}));

test('every numbered part of a tiled deck carries at least one service', () => Promise.resolve().then(async () => {
  // A grid cell can cover a stretch of drawing that holds only zone outlines
  // and pass-through arrows. That was still shipped as "part 1 of 6" — a slide
  // with nothing on it that a reader could name.
  const nodes: Node[] = [];
  const edges: Edge[] = [];
  // An L: services down the left and along the bottom, so the top-right of the
  // bounding box is empty and its cells own nothing.
  for (let i = 0; i < 9; i += 1) nodes.push(service(`down-${i}`, `Down Service ${i}`, 0, i * 240));
  for (let i = 1; i < 10; i += 1) nodes.push(service(`across-${i}`, `Across Service ${i}`, i * 300, 8 * 240));
  for (let i = 1; i < nodes.length; i += 1) {
    edges.push({ id: `e${i}`, source: nodes[i - 1].id, target: nodes[i].id } as Edge);
  }

  const deck = await buildDeck(nodes, edges);
  // The numbered workflow list continues onto its own slides, which carry a
  // `(n / m)` footer of their own and legitimately hold no tiles.
  const diagramParts = deck.parts.filter((slide) => !/Workflow \(/.test(slide));
  assert.ok(diagramParts.length > 1, `expected a tiled deck, got ${diagramParts.length} part(s)`);
  const empty = diagramParts
    .map((slide, i) => ({ i: i + 1, tiles: shapeBoxes(slide, 'service-').filter((b) => !b.name.startsWith('service-label-')).length }))
    .filter((part) => part.tiles === 0);
  assert.deepEqual(empty, [], `blank part slide(s): ${empty.map((p) => `part ${p.i}`).join(', ')}`);
}));

test('a numbered hop across an empty stretch keeps its label and its callout', () => Promise.resolve().then(async () => {
  // A barbell: two dense clusters with one long bridge between them. The grid
  // cell the bridge's label sits in holds no service, so it was dropped from
  // the deck — and the half-open ownership test only tiles the drawing while
  // every cell survives. With a hole punched in the middle, the cell to the
  // left of the label saw it as past its right edge and the cell to the right
  // saw it as before its left edge, so neither drew it. The arrow was still
  // drawn (it overlaps both parts) and the Workflow slide still listed the
  // step, leaving an unlabelled, unnumbered connector.
  const nodes: Node[] = [];
  const edges: Edge[] = [];
  for (let i = 0; i < 6; i += 1) nodes.push(service(`l${i}`, `Left Service ${i}`, (i % 2) * 220, Math.floor(i / 2) * 200));
  for (let i = 0; i < 6; i += 1) nodes.push(service(`r${i}`, `Right Service ${i}`, 3200 + (i % 2) * 220, Math.floor(i / 2) * 200));
  for (let i = 1; i < 6; i += 1) edges.push({ id: `le${i}`, source: `l${i - 1}`, target: `l${i}`, label: `left hop ${i}`, data: { stepNumber: i } } as Edge);
  edges.push({ id: 'bridge', source: 'l5', target: 'r0', label: 'private peering', data: { stepNumber: 6 } } as Edge);
  for (let i = 1; i < 6; i += 1) edges.push({ id: `re${i}`, source: `r${i - 1}`, target: `r${i}`, label: `right hop ${i}`, data: { stepNumber: i + 6 } } as Edge);

  const deck = await buildDeck(nodes, edges);
  const parts = deck.parts.filter((slide) => !/Workflow \(/.test(slide));
  assert.ok(parts.length > 1, `expected a tiled deck, got ${parts.length} part(s)`);

  const orphans: string[] = [];
  for (const edge of edges) {
    const labels = parts.filter((slide) => shapeBoxes(slide, `connector-label-${edge.id}`).length > 0).length;
    const badges = parts.filter((slide) => shapeBoxes(slide, `connector-step-${edge.id}`).length > 0).length;
    if (labels === 0 || badges === 0) orphans.push(`${edge.id}: label on ${labels} part(s), callout on ${badges}`);
  }
  assert.deepEqual(orphans, [], `numbered hops drawn on no part: ${orphans.join('; ')}`);
}));

test('a drawing too dense for one slide is grown, not squeezed to 4pt', () => Promise.resolve().then(async () => {
  // Every service inside a single zone lands in one grid cell, so the tiling
  // pass had nothing to split and reported the drawing as legible. The page
  // was then left at the standard 13.33 x 7.5in and the labels were scaled
  // down to fit — 4.2pt on the widest fixture, which no projector can render.
  const nodes: Node[] = [{ id: 'z', type: 'groupNode', position: { x: 0, y: 0 }, style: { width: 2800, height: 1000 }, data: { label: 'Landing zone' } } as Node];
  const edges: Edge[] = [];
  for (let i = 0; i < 4; i += 1) {
    nodes.push({ ...service(`c${i}`, `Clustered Service ${i}`, 40 + (i % 2) * 200, 40 + Math.floor(i / 2) * 130), parentNode: 'z' } as Node);
    if (i > 0) edges.push({ id: `e${i}`, source: `c${i - 1}`, target: `c${i}` } as Edge);
  }

  const deck = await buildDeck(nodes, edges);
  const fonts = deck.slides
    .flatMap((slide) => [...slide.matchAll(/name="service-label-[^"]*"[\s\S]{0,600}?sz="(\d+)"/g)])
    .map((m) => Number(m[1]) / 100);
  assert.ok(fonts.length > 0, 'expected service labels in the deck');
  const smallest = Math.min(...fonts);
  assert.ok(smallest >= 7, `service labels shrank to ${smallest.toFixed(2)}pt, below the 7pt legibility floor`);
}));

/** Connector annotations with the point size they are actually drawn at. */
function annotationBoxes(slideXml: string): { name: string; x: number; y: number; w: number; h: number; pt: number }[] {
  const boxes: { name: string; x: number; y: number; w: number; h: number; pt: number }[] = [];
  const EMU = 914400;
  for (const match of slideXml.matchAll(/<p:sp>[\s\S]*?<\/p:sp>/g)) {
    const name = /name="([^"]*)"/.exec(match[0])?.[1] ?? '';
    if (!/^connector-(label|step)-/.test(name)) continue;
    const off = /<a:off x="(-?\d+)" y="(-?\d+)"\/>/.exec(match[0]);
    const ext = /<a:ext cx="(-?[\d.]+)" cy="(-?[\d.]+)"\/>/.exec(match[0]);
    if (!off || !ext) continue;
    const sz = /sz="(\d+)"/.exec(match[0]);
    boxes.push({
      name, x: +off[1] / EMU, y: +off[2] / EMU, w: +ext[1] / EMU, h: +ext[2] / EMU, pt: sz ? +sz[1] / 100 : 0,
    });
  }
  return boxes;
}

/** Pairs of connector annotations that visibly sit on each other. */
function annotationClashes(slides: string[]): string[] {
  const clashes: string[] = [];
  for (const slide of slides) {
    const boxes = annotationBoxes(slide);
    for (let i = 0; i < boxes.length; i += 1) {
      for (let j = i + 1; j < boxes.length; j += 1) {
        const a = boxes[i];
        const b = boxes[j];
        const dx = Math.min(a.x + a.w, b.x + b.w) - Math.max(a.x, b.x);
        const dy = Math.min(a.y + a.h, b.y + b.h) - Math.max(a.y, b.y);
        if (dx > 0.01 && dy > 0.01) clashes.push(`${a.name}/${b.name}`);
      }
    }
  }
  return clashes;
}

const FAN_LABEL = 'マネージド ID を使用して注文ドキュメントを Cosmos DB に書き込みます';

function fan(count: number, label: (index: number) => string): { nodes: Node[]; edges: Edge[] } {
  return {
    nodes: [service('a', 'Azure Front Door', 0, 0), service('b', 'Azure Kubernetes Service', 190, 0)],
    edges: Array.from({ length: count }, (_, i) => ({
      id: `p${i + 1}`,
      source: 'a',
      target: 'b',
      label: label(i),
      data: { stepNumber: i + 1 },
    })) as Edge[],
  };
}

test('a shrunken fan is drawn at the size its chips were measured at', () => Promise.resolve().then(async () => {
  // A deep fan is laid out at a reduced point size so the rungs fit the frame.
  // Handing the draw call the ordinary size instead leaves the text spilling out
  // of its own chip and over its own numbered callout - invisible to any test
  // that only compares shape rectangles, because the rectangles do not move.
  for (const count of [7, 8, 14]) {
    const { nodes, edges } = fan(count, (i) => `${FAN_LABEL} ${i + 1}`);
    const deck = await buildDeck(nodes, edges);
    const chips = deck.slides.flatMap(annotationBoxes).filter((box) => box.name.startsWith('connector-label-'));
    assert.equal(chips.length, count, `fan of ${count} lost a label`);
    for (const chip of chips) {
      assert.ok(chip.pt >= 7, `${chip.name} is drawn at ${chip.pt}pt, below the legibility floor`);
      const line = (chip.pt * 1.3) / 72;
      const lines = Math.max(1, Math.round((chip.h - 0.06) / line));
      const needed = lines * line + 0.06;
      assert.ok(
        needed <= chip.h + 0.02,
        `${chip.name} is drawn at ${chip.pt}pt needing ${needed.toFixed(3)}in inside a ${chip.h.toFixed(3)}in chip`,
      );
    }
  }
}));

test('a fan whose first label is short still lays out on one lattice', () => Promise.resolve().then(async () => {
  // The rung height is the tallest chip in the bundle. Re-measuring only the
  // first route after a shrink reads a short label's height, the lattice
  // collapses below its own step, and the long siblings overlap.
  for (const count of [11, 12, 14, 20]) {
    const { nodes, edges } = fan(count, (i) => (i === 0 ? '短' : `${FAN_LABEL} ${i + 1}`));
    const deck = await buildDeck(nodes, edges);
    const clashes = annotationClashes(deck.slides);
    assert.deepEqual(clashes, [], `short-first fan of ${count} stacks annotations: ${clashes.slice(0, 4).join(' ')}`);
  }
}));

test('a fan of parallel edges does not displace the ordinary connectors around it', () => Promise.resolve().then(async () => {
  // A ladder is far larger than one chip, so on a busy slide it has to be the
  // thing that dodges. Placing it first, or scoring it against the service
  // tiles alone, pushed unrelated labels into each other instead.
  const nodes: Node[] = [];
  for (let row = 0; row < 4; row += 1) {
    for (let col = 0; col < 5; col += 1) nodes.push(service(`t${row}-${col}`, `Service ${row}${col}`, col * 290, row * 180));
  }
  for (const depth of [5, 7, 8]) {
    const edges: Edge[] = [];
    for (let i = 1; i < nodes.length; i += 1) {
      edges.push({
        id: `q${i}`, source: nodes[i - 1].id, target: nodes[i].id, label: `ホップ ${i}`, data: { stepNumber: i },
      } as Edge);
    }
    for (let i = 0; i < depth; i += 1) {
      edges.push({
        id: `bn${i}`,
        source: 't0-0',
        target: 't0-1',
        label: `マネージド ID で参照系を照会します ${i + 1}`,
        data: { stepNumber: nodes.length + i },
      } as Edge);
    }
    const deck = await buildDeck(nodes, edges);
    const clashes = annotationClashes(deck.slides);
    assert.deepEqual(clashes, [], `a ${depth}-edge ladder displaced neighbours: ${clashes.slice(0, 4).join(' ')}`);
  }
}));

test('a three-digit callout still fits inside its own bubble', () => Promise.resolve().then(async () => {
  // The callout circle is sized by the drawing scale, not by the label it hangs
  // off. Drawing the number at the label's point size runs a 3-digit step
  // straight out of its own bubble and over whatever the arrow passes through.
  const nodes: Node[] = [];
  const edges: Edge[] = [];
  for (let i = 0; i < 8; i += 1) nodes.push(service(`s${i}`, `Service ${i}`, (i % 4) * 300, Math.floor(i / 4) * 220));
  for (let i = 1; i < nodes.length; i += 1) {
    edges.push({
      id: `e${i}`, source: `s${i - 1}`, target: `s${i}`, label: `step ${i}`, data: { stepNumber: 99 + i },
    } as Edge);
  }
  const deck = await buildDeck(nodes, edges);
  const badges = deck.slides.flatMap(annotationBoxes).filter((box) => box.name.startsWith('connector-step-'));
  assert.ok(badges.length > 0, 'no callouts were drawn');
  for (const badge of badges) {
    const digits = 3;
    const wide = digits * 0.62 * (badge.pt / 72);
    const tall = (badge.pt * 1.3) / 72;
    assert.ok(
      wide <= badge.w + 0.005 && tall <= badge.h + 0.005,
      `${badge.name} draws 3 digits at ${badge.pt}pt needing ${wide.toFixed(3)}x${tall.toFixed(3)}in inside a ${badge.w.toFixed(3)}in circle`,
    );
  }
}));

test('a very large architecture is tiled onto standard slides, not a plotter page', () => Promise.resolve().then(async () => {
  // PowerPoint gives a deck exactly one page size, so growing the page for the
  // drawing drags the title and workflow slides onto the plotter sheet too.
  // Tiling a 56in sheet into four 56in parts is not an escape from that: every
  // part still inherits the page. Standard slides are the only usable answer
  // short of a drawing that genuinely cannot be tiled at all.
  const nodes: Node[] = [];
  const edges: Edge[] = [];
  for (let i = 0; i < 40; i += 1) {
    nodes.push(service(`n-${i}`, i % 2 ? 'Azure Kubernetes Service' : 'Copilot Studio', i * 260, (i % 4) * 220));
    if (i > 0) edges.push({ id: `x-${i}`, source: `n-${i - 1}`, target: `n-${i}`, label: 'managed identity' } as Edge);
  }

  const pptx = await buildDiagramSlidePptx(PIXEL_PNG, {
    diagramName: 'Contoso Platform',
    author: 'Tester',
    date: '2026-08-10',
    isDarkMode: false,
    diagram: { nodes, edges },
  });
  const presLayout = (pptx as unknown as { presLayout: { width: number; height: number } }).presLayout;
  const pageW = presLayout.width / 914400;
  const pageH = presLayout.height / 914400;
  assert.ok(
    Math.abs(pageW - 13.333) < 0.05 && Math.abs(pageH - 7.5) < 0.05,
    `the deck is a ${pageW.toFixed(2)}x${pageH.toFixed(2)}in page instead of standard 13.33x7.5in slides`,
  );

  const deck = await buildDeck(nodes, edges);
  const fonts = deck.parts
    .flatMap((slide) => [...slide.matchAll(/name="service-label-[^"]*"[\s\S]{0,600}?sz="(\d+)"/g)])
    .map((m) => Number(m[1]) / 100);
  assert.ok(fonts.length > 0, 'expected service labels in the deck');
  const smallest = Math.min(...fonts);
  assert.ok(smallest >= 7, `tiling dropped service labels to ${smallest.toFixed(2)}pt`);
}));

/** Every shape on a slide that is a drawn connector rather than an annotation. */
function arrowBoxes(slideXml: string): { name: string; x: number; y: number; w: number; h: number }[] {
  return shapeBoxes(slideXml, 'connector-').filter((box) => !/^connector-(label|step)-/.test(box.name));
}

/** Distance from a point to the nearest edge of a box, zero when inside it. */
function gap(box: { x: number; y: number; w: number; h: number }, at: { x: number; y: number }): number {
  return Math.hypot(
    at.x - Math.max(box.x, Math.min(at.x, box.x + box.w)),
    at.y - Math.max(box.y, Math.min(at.y, box.y + box.h)),
  );
}

/** A grid of services wired left to right, one long CJK label per hop. */
function wiredGrid(cols: number, rows: number): { nodes: Node[]; edges: Edge[] } {
  const nodes: Node[] = [];
  for (let row = 0; row < rows; row += 1) {
    for (let col = 0; col < cols; col += 1) {
      nodes.push(service(`g${row}-${col}`, `Azure Service ${row}${col}`, 120 + col * 290, 120 + row * 180));
    }
  }
  const edges: Edge[] = [];
  let step = 1;
  for (let row = 0; row < rows; row += 1) {
    for (let col = 0; col + 1 < cols; col += 1) {
      edges.push({
        id: `h${row}_${col}`, source: `g${row}-${col}`, target: `g${row}-${col + 1}`,
        label: FAN_LABEL, data: { stepNumber: step },
      } as Edge);
      step += 1;
    }
  }
  return { nodes, edges };
}

test('the ladder stacks its rungs in the same order the arrows fan out', () => Promise.resolve().then(async () => {
  // Parallel arrows are fanned alternately about the centre line - 0, +1, -1,
  // +2 - while a ladder runs straight down. Ranking the rungs by edge ordinal
  // therefore parks callout 1 beside the middle arrow and callout 3 beside the
  // top one: every chip is legible, every chip is on the wrong hop.
  for (const count of [3, 5, 7]) {
    const nodes = [service('a', 'Azure Front Door', 100, 400), service('b', 'Azure Kubernetes Service', 900, 400)];
    const edges = Array.from({ length: count }, (_, i) => ({
      id: `p${i + 1}`, source: 'a', target: 'b', label: `${FAN_LABEL} ${i + 1}`, data: { stepNumber: i + 1 },
    })) as Edge[];
    const fanned = new Map(buildExportRoutes(edges, collectExportBoxes(nodes)).map((r) => [r.id, r.fanOffset]));
    const deck = await buildDeck(nodes, edges);
    const chips = deck.slides
      .flatMap(annotationBoxes)
      .filter((box) => box.name.startsWith('connector-label-'))
      .map((box) => ({ ...box, id: box.name.replace('connector-label-', '') }));
    assert.equal(chips.length, count, `fan of ${count} lost a label`);
    chips.sort((l, r) => (fanned.get(l.id) ?? 0) - (fanned.get(r.id) ?? 0));
    const order = chips.map((chip) => chip.id).join(',');
    for (let i = 1; i < chips.length; i += 1) {
      assert.ok(
        chips[i].y >= chips[i - 1].y - 0.001,
        `rung order does not follow the arrows: ${chips[i - 1].id} sits below ${chips[i].id} (top to bottom by arrow: ${order})`,
      );
    }
  }
}));

test('repairing an overlap never carries a chip onto a different arrow', () => Promise.resolve().then(async () => {
  // The repair pass walks a chip until it stops overlapping. Unbounded, on a
  // dense grid the nearest clear spot is a row away, and the label ends up
  // perfectly legible beside somebody else's hop - worse than the overlap it
  // escaped, because the reader cannot tell it is on the wrong arrow.
  const nodes: Node[] = [];
  const edges: Edge[] = [];
  let step = 1;
  for (let row = 0; row < 6; row += 1) {
    for (let col = 0; col < 6; col += 1) nodes.push(service(`n${row}-${col}`, `Azure Service ${row}${col}`, col * 200, row * 140));
  }
  for (let row = 0; row < 6; row += 1) {
    for (let col = 0; col + 1 < 6; col += 1) {
      edges.push({
        id: `x${row}_${col}`, source: `n${row}-${col}`, target: `n${row}-${col + 1}`,
        label: `${FAN_LABEL} ${step}`, data: { stepNumber: step },
      } as Edge);
      step += 1;
    }
  }
  for (let col = 0; col < 6; col += 2) {
    for (let row = 0; row + 1 < 6; row += 1) {
      edges.push({
        id: `y${col}_${row}`, source: `n${row}-${col}`, target: `n${row + 1}-${col}`,
        label: `${FAN_LABEL} ${step}`, data: { stepNumber: step },
      } as Edge);
      step += 1;
    }
  }
  const deck = await buildDeck(nodes, edges);
  let checked = 0;
  let strayed = 0;
  let worst = { name: '', d: 0 };
  const strays: string[] = [];
  for (const slide of deck.slides) {
    const arrows = arrowBoxes(slide);
    if (arrows.length === 0) continue;
    for (const chip of shapeBoxes(slide, 'connector-label-')) {
      const own = arrows.find((arrow) => arrow.name === `connector-${chip.name.replace('connector-label-', '')}`);
      if (!own) continue;
      checked += 1;
      const at = { x: chip.x + chip.w / 2, y: chip.y + chip.h / 2 };
      const mine = gap(own, at);
      if (mine > worst.d) worst = { name: chip.name, d: mine };
      const nearest = arrows.reduce((best, arrow) => (gap(arrow, at) < gap(best, at) ? arrow : best), arrows[0]);
      if (nearest.name !== own.name && gap(nearest, at) < mine - 0.05) {
        strayed += 1;
        strays.push(`${chip.name} is ${gap(nearest, at).toFixed(2)}in from ${nearest.name} but ${mine.toFixed(2)}in from its own arrow`);
      }
    }
  }
  assert.ok(checked > 30, `only ${checked} chips were checked, the fixture did not build`);
  // Arrows genuinely cross on a grid this tight, so the nearest-arrow reading
  // is not exact - but a handful is a crossing and half the deck is a bug.
  assert.ok(
    strayed <= checked * 0.05,
    `${strayed} of ${checked} chips sit closer to a foreign arrow: ${strays.slice(0, 3).join('; ')}`,
  );
  assert.ok(worst.d <= 1.5, `${worst.name} was carried ${worst.d.toFixed(2)}in away from the hop it describes`);
}));

test('two ladders on one slide each find their own lane', () => Promise.resolve().then(async () => {
  // One bundle dodges the tiles and the other dodges the tiles, and both pick
  // the same clear band: a bundle has to be re-scored against the annotations
  // already placed, not only against the drawing.
  for (const [first, second] of [[4, 4], [3, 10]]) {
    const { nodes, edges } = wiredGrid(4, 5);
    for (let i = 0; i < first; i += 1) {
      edges.push({
        id: `f${i}`, source: 'g1-0', target: 'g1-1', label: `${FAN_LABEL} ${i + 1}`, data: { stepNumber: 100 + i },
      } as Edge);
    }
    for (let i = 0; i < second; i += 1) {
      edges.push({
        id: `k${i}`, source: 'g2-0', target: 'g2-1', label: `${FAN_LABEL} ${i + 1}`, data: { stepNumber: 200 + i },
      } as Edge);
    }
    const deck = await buildDeck(nodes, edges);
    const clashes = annotationClashes(deck.slides);
    assert.deepEqual(clashes, [], `fans of ${first} and ${second} collide: ${clashes.slice(0, 4).join(' ')}`);
  }
}));

test('a fan on a vertical arrow is spaced by the width of its chips', () => Promise.resolve().then(async () => {
  // A vertical hop stacks its chips side by side, so the rung is a width. Using
  // the height there - the horizontal case - spaces a column of 2in labels
  // 0.3in apart and every one of them lands on its neighbour.
  for (const count of [3, 5, 7]) {
    const nodes = [service('a', 'Azure Front Door', 700, 120), service('b', 'Azure Kubernetes Service', 700, 420)];
    const edges = Array.from({ length: count }, (_, i) => ({
      id: `v${i + 1}`,
      source: 'a',
      target: 'b',
      label: `カスタム ドメインの TLS 終端とヘッダー書き換えを行います ${i + 1}`,
      data: { stepNumber: i + 1 },
    })) as Edge[];
    const deck = await buildDeck(nodes, edges);
    const clashes = annotationClashes(deck.slides);
    assert.deepEqual(clashes, [], `a vertical fan of ${count} stacks annotations: ${clashes.slice(0, 4).join(' ')}`);
    const chips = deck.slides
      .flatMap(annotationBoxes)
      .filter((box) => box.name.startsWith('connector-label-'))
      .sort((l, r) => l.x - r.x);
    assert.equal(chips.length, count, `a vertical fan of ${count} lost labels: only ${chips.length} survived`);
    for (let i = 1; i < chips.length; i += 1) {
      const pitch = chips[i].x - chips[i - 1].x;
      assert.ok(
        pitch >= chips[i - 1].w * 0.95,
        `vertical rungs are ${pitch.toFixed(2)}in apart for ${chips[i - 1].w.toFixed(2)}in chips, so the ladder is a pile`,
      );
      assert.ok(
        Math.abs(chips[i].y - chips[0].y) <= chips[i].h,
        `${chips[i].name} sits ${Math.abs(chips[i].y - chips[0].y).toFixed(2)}in off the row its siblings are on`,
      );
    }
  }
}));


/** Step numbers narrated by the workflow slides, from the deck itself. */
function narratedSteps(slides: string[]): Set<number> {
  const steps = new Set<number>();
  for (const slide of slides) {
    for (const match of slide.matchAll(/name="workflow-text-(\d+)"/g)) steps.add(+match[1]);
  }
  return steps;
}

/** Step numbers drawn as callouts on the drawing, per connector. */
function drawnCallouts(slides: string[]): Map<string, number> {
  const drawn = new Map<string, number>();
  for (const slide of slides) {
    for (const match of slide.matchAll(/<p:sp>[\s\S]*?<\/p:sp>/g)) {
      const name = /name="connector-step-([^"]*)"/.exec(match[0])?.[1];
      if (!name) continue;
      const text = /<a:t>([^<]*)<\/a:t>/.exec(match[0])?.[1] ?? '';
      const step = Number.parseInt(text, 10);
      if (Number.isFinite(step)) drawn.set(name, step);
    }
  }
  return drawn;
}

/** Connector ids that still carry readable chip text. */
function chippedRoutes(slides: string[]): Set<string> {
  const carried = new Set<string>();
  for (const slide of slides) {
    for (const match of slide.matchAll(/<p:sp>[\s\S]*?<\/p:sp>/g)) {
      const name = /name="connector-label-([^"]*)"/.exec(match[0])?.[1];
      if (!name) continue;
      if (/<a:t>[^<]+<\/a:t>/.test(match[0])) carried.add(name);
    }
  }
  return carried;
}

/** A grid of services wired left-to-right and top-to-bottom, `fanned` times over. */
function fannedGrid(cols: number, rows: number, fanned: number): { nodes: Node[]; edges: Edge[] } {
  const nodes: Node[] = [];
  const edges: Edge[] = [];
  let step = 1;
  for (let r = 0; r < rows; r += 1) {
    for (let c = 0; c < cols; c += 1) nodes.push(service(`n${r}-${c}`, `Azure Service ${r}${c}`, c * 260, r * 170));
  }
  const wire = (id: string, source: string, target: string, label: string): void => {
    edges.push({ id, source, target, label, data: { stepNumber: step } } as Edge);
    step += 1;
  };
  for (let r = 0; r < rows; r += 1) {
    for (let c = 0; c + 1 < cols; c += 1) {
      for (let f = 0; f < fanned; f += 1) wire(`h${r}_${c}_${f}`, `n${r}-${c}`, `n${r}-${c + 1}`, `${FAN_LABEL} ${f + 1}`);
    }
  }
  for (let c = 0; c < cols; c += 1) {
    for (let r = 0; r + 1 < rows; r += 1) {
      for (let f = 0; f < fanned; f += 1) wire(`v${c}_${r}_${f}`, `n${r}-${c}`, `n${r + 1}-${c}`, `${FAN_LABEL} ${f + 1}`);
    }
  }
  return { nodes, edges };
}

test('a labelled edge never loses both its chip and its narration', () => Promise.resolve().then(async () => {
  // The deep-fan fallback drops the chips and numbers the arrows instead. Only
  // the pairs the workflow slide narrates carry a step number, so every other
  // member of the fan lost its wording entirely: no chip, no callout, no row.
  // That is a silent deletion of the one thing the arrow was drawn to say.
  const { nodes, edges } = fannedGrid(4, 3, 5);
  const deck = await buildDeck(nodes, edges);
  const chips = chippedRoutes(deck.slides);
  const callouts = drawnCallouts(deck.slides);
  const narrated = narratedSteps(deck.slides);
  const lost = edges
    .filter((edge) => {
      const step = callouts.get(edge.id);
      return !chips.has(edge.id) && (step === undefined || !narrated.has(step));
    })
    .map((edge) => edge.id);
  assert.deepEqual(lost, [], `${lost.length} of ${edges.length} labelled edges say nothing at all: ${lost.slice(0, 6).join(' ')}`);
}));

test('every callout the drawing shows is a row the workflow explains', () => Promise.resolve().then(async () => {
  // A numbered bubble with no matching row sends the reader to a list that does
  // not mention it. Backfilling the fan members has to backfill the narration
  // too, or the fallback trades one silent failure for another.
  const { nodes, edges } = fannedGrid(3, 3, 4);
  const deck = await buildDeck(nodes, edges);
  const narrated = narratedSteps(deck.slides);
  const orphans = [...drawnCallouts(deck.slides)].filter(([, step]) => !narrated.has(step));
  assert.deepEqual(orphans.map(([id, step]) => `${id}#${step}`), [], 'callouts point at workflow rows that were never written');
}));

test('a deep fan of arrows stays inside its own row', () => Promise.resolve().then(async () => {
  // The fan spreads its arrows by a fixed step per ordinal, so a nine-way fan
  // reached most of a row's height in each direction and the outer arrows ran
  // straight through the services above and below.
  const nodes = [service('a', 'Azure Front Door', 0, 300), service('b', 'Azure Kubernetes Service', 400, 300)];
  const edges = Array.from({ length: 9 }, (_, i) => ({
    id: `p${i + 1}`, source: 'a', target: 'b', label: `${FAN_LABEL} ${i + 1}`, data: { stepNumber: i + 1 },
  })) as Edge[];
  const routes = buildExportRoutes(edges, collectExportBoxes(nodes));
  const centre = 300 + 75 / 2;
  const spread = Math.max(...routes.flatMap((route) => route.points.map((point) => Math.abs(point.y - centre))));
  assert.ok(
    spread <= 75 / 2 + 0.01,
    `the fan reaches ${spread.toFixed(1)}px off centre, past the ${(75 / 2).toFixed(1)}px half-row it is allowed`,
  );
}));

test('a dense diagram of parallel flows exports without hanging', () => Promise.resolve().then(async () => {
  // The ladder search scored every candidate offset by re-measuring and
  // re-wrapping every rung, over a lattice that reached the far side of the
  // page: a thirty-service grid of three-way fans took 42 seconds, which in a
  // browser is a frozen tab, not an export.
  const { nodes, edges } = fannedGrid(6, 5, 3);
  assert.ok(edges.length >= 140, `the fixture must stay dense, got ${edges.length} edges`);
  const started = Date.now();
  await buildDeck(nodes, edges);
  const elapsed = Date.now() - started;
  assert.ok(elapsed < 15000, `exporting ${edges.length} labelled edges took ${(elapsed / 1000).toFixed(1)}s`);
}));

/** Every shape on a slide with its page rectangle in inches. */
function boxesOn(slide: string): { name: string; x: number; y: number; w: number; h: number }[] {
  const out: { name: string; x: number; y: number; w: number; h: number }[] = [];
  for (const match of slide.matchAll(/<p:(?:sp|cxnSp)>[\s\S]*?<\/p:(?:sp|cxnSp)>/g)) {
    const name = /name="([^"]*)"/.exec(match[0])?.[1];
    const off = /<a:off x="(-?\d+)" y="(-?\d+)"\/><a:ext cx="(\d+)" cy="(\d+)"/.exec(match[0]);
    if (!name || !off) continue;
    out.push({ name, x: +off[1] / 914400, y: +off[2] / 914400, w: +off[3] / 914400, h: +off[4] / 914400 });
  }
  return out;
}

/** Distance from a point to a rectangle's outline-as-diagonal, as the audit measures arrows. */
function gapToArrow(arrow: { x: number; y: number; w: number; h: number }, px: number, py: number): number {
  const dx = Math.max(arrow.x - px, 0, px - (arrow.x + arrow.w));
  const dy = Math.max(arrow.y - py, 0, py - (arrow.y + arrow.h));
  return Math.hypot(dx, dy);
}

test('a fan of labels never parks itself beside somebody else s arrow', () => Promise.resolve().then(async () => {
  // A ladder was scored only on what it covered, so on a roomy grid it walked
  // to whatever clear air it could find - which is beside a DIFFERENT hop. The
  // placement is spotless by every collision rule and the reader still credits
  // every rung to the wrong arrow, which is worse than clipping a tile because
  // nothing about it looks wrong.
  const nodes: Node[] = [];
  for (let r = 0; r < 3; r += 1) {
    for (let c = 0; c < 4; c += 1) nodes.push(service(`s${r}-${c}`, `Service ${r}${c}`, c * 260, r * 170));
  }
  const edges: Edge[] = [];
  let step = 0;
  for (let r = 0; r < 3; r += 1) {
    for (let c = 0; c + 1 < 4; c += 1) {
      step += 1;
      edges.push({ id: `h${r}_${c}`, source: `s${r}-${c}`, target: `s${r}-${c + 1}`, label: 'マネージド ID で参照系を照会します', data: { stepNumber: step } } as Edge);
    }
  }
  for (let i = 0; i < 8; i += 1) {
    step += 1;
    edges.push({ id: `fan${i}`, source: 's1-0', target: 's1-1', label: `注文ドキュメントを Cosmos DB に書き込みます ${i + 1}`, data: { stepNumber: step } } as Edge);
  }
  const deck = await buildDeck(nodes, edges);
  const strays: string[] = [];
  for (const slide of deck.slides) {
    const shapes = boxesOn(slide);
    const arrows = shapes.filter((s) => /^connector-/.test(s.name) && !/^connector-(label|step)-/.test(s.name));
    if (arrows.length === 0) continue;
    for (const chip of shapes.filter((s) => s.name.startsWith('connector-label-'))) {
      const own = arrows.find((a) => a.name === `connector-${chip.name.slice('connector-label-'.length)}`);
      if (!own) continue;
      const cx = chip.x + chip.w / 2;
      const cy = chip.y + chip.h / 2;
      const mine = gapToArrow(own, cx, cy);
      for (const rival of arrows) {
        if (rival.name === own.name) continue;
        if (gapToArrow(rival, cx, cy) < mine - 0.25) {
          strays.push(`${chip.name} is ${mine.toFixed(2)}in from its own arrow but ${gapToArrow(rival, cx, cy).toFixed(2)}in from ${rival.name}`);
          break;
        }
      }
    }
  }
  assert.deepEqual(strays, [], `${strays.length} label(s) sit nearer a foreign arrow: ${strays.slice(0, 4).join('; ')}`);
}));

test('the colour key never stands on a tile or a numbered callout', () => Promise.resolve().then(async () => {
  // The key used to be stamped as a card over the drawing, so on a full grid it
  // buried a service tile by 92% and hid a numbered callout outright. It is a
  // strip in a band the tiler reserves, which is why the two agree on how much
  // height the drawing actually gets.
  const nodes: Node[] = [];
  for (let r = 0; r < 4; r += 1) {
    for (let c = 0; c < 6; c += 1) nodes.push(service(`g${r}${c}`, `Azure Service ${r}${c}`, c * 260, r * 190));
  }
  const kinds = ['sync', 'async', 'telemetry', 'data'];
  const edges: Edge[] = [];
  let step = 0;
  for (let r = 0; r < 4; r += 1) {
    for (let c = 0; c + 1 < 6; c += 1) {
      step += 1;
      edges.push({
        id: `h${r}${c}`, source: `g${r}${c}`, target: `g${r}${c + 1}`,
        label: 'マネージド ID で注文ドキュメントを書き込みます',
        data: { connectionType: kinds[r % 4], stepNumber: step, stepDescription: `手順 ${step}` },
      } as Edge);
    }
  }
  const deck = await buildDeck(nodes, edges);
  const buried: string[] = [];
  let keys = 0;
  for (const slide of deck.slides) {
    const shapes = boxesOn(slide);
    const key = shapes.find((s) => s.name === 'connection-legend');
    if (!key) continue;
    keys += 1;
    for (const s of shapes) {
      if (s.name === 'connection-legend') continue;
      const isTile = s.name.startsWith('service-') && !s.name.includes('label') && !s.name.includes('meta');
      if (!isTile && !s.name.startsWith('connector-step-') && !s.name.startsWith('connector-label-')) continue;
      const w = Math.max(0, Math.min(key.x + key.w, s.x + s.w) - Math.max(key.x, s.x));
      const h = Math.max(0, Math.min(key.y + key.h, s.y + s.h) - Math.max(key.y, s.y));
      const own = Math.max(0.0001, s.w * s.h);
      if (w * h > 0.02 * own) buried.push(`${s.name} is ${((w * h) / own * 100).toFixed(0)}% under the colour key`);
    }
  }
  assert.ok(keys > 0, 'the fixture must draw a colour key for this to test anything');
  assert.deepEqual(buried, [], `${buried.length} shape(s) sit under the key: ${buried.slice(0, 4).join('; ')}`);
}));

test('a muted fan hands its wording to the workflow row that replaces it', async () => {
  // A stuck fan trades its chips for numbered callouts and the workflow slide.
  // That is only a trade if the slide says what the chips said: an author who
  // writes a terse description as well as a real label used to lose the label
  // outright, leaving the deck reading "13. Step 13" with the sentence gone.
  const nodes: Node[] = [];
  for (let r = 0; r < 3; r += 1) {
    for (let c = 0; c < 3; c += 1) nodes.push(service(`g${r}${c}`, `Azure Service ${r}${c}`, c * 300, r * 200));
  }
  const edges: Edge[] = [];
  let step = 0;
  for (let r = 0; r < 3; r += 1) {
    for (let c = 0; c + 1 < 3; c += 1) {
      step += 1;
      edges.push({ id: `h${r}${c}`, source: `g${r}${c}`, target: `g${r}${c + 1}`, label: 'writes the order document', data: { stepNumber: step, stepDescription: `Step ${step}` } } as Edge);
    }
  }
  for (let r = 0; r + 1 < 3; r += 1) {
    for (let c = 0; c < 3; c += 1) {
      step += 1;
      edges.push({ id: `v${r}${c}`, source: `g${r}${c}`, target: `g${r + 1}${c}`, label: 'writes the order document', data: { stepNumber: step, stepDescription: `Step ${step}` } } as Edge);
    }
  }
  const fanLabels: string[] = [];
  for (let i = 0; i < 5; i += 1) {
    step += 1;
    const label = `replicates the ledger ${i}`;
    fanLabels.push(label);
    edges.push({ id: `f${i}`, source: 'g11', target: 'g12', label, data: { stepNumber: step, stepDescription: `Step ${step}` } } as Edge);
  }

  const deck = await buildDeck(nodes, edges);
  const drawn = deck.slides.join(' ');
  const chips = new Set(
    [...drawn.matchAll(/name="connector-label-([^"]+)"/g)].map((m) => m[1]),
  );
  const muted = fanLabels.filter((_, i) => !chips.has(`f${i}`));
  assert.ok(muted.length > 0, 'the fixture must mute at least one chip for this to test anything');
  const lost = muted.filter((label) => !drawn.includes(label));
  assert.deepEqual(lost, [], `${lost.length} muted label(s) appear nowhere in the deck: ${lost.join('; ')}`);
});
test('the overview thumbnail never draws type the reader cannot resolve', async () => {
  // The overview is a map, not a reading surface, so it is exempt from the
  // legibility floor the detail slides are held to. It is not exempt from
  // being drawn at all: at 40 services the tile labels clamp to 4pt, which is
  // grey ink rather than small words. Every one of those strings is legible on
  // the slide that follows, so the thumbnail shows the shape and leaves the
  // wording to them.
  const { nodes, edges } = wideDiagram(40);
  const deck = await buildDeck(nodes, edges);
  assert.ok(deck.parts.length > 1, 'the fixture must tile for there to be an overview at all');
  const overview = deck.slides[deck.slides.length - deck.parts.length - 1];
  assert.ok(overview.includes('(Overview)'), 'located the overview slide');

  const RESOLVABLE_PT = 6;
  const tooSmall: string[] = [];
  for (const shape of overview.matchAll(/<p:sp>[\s\S]*?<\/p:sp>/g)) {
    const text = [...shape[0].matchAll(/<a:t>([^<]*)<\/a:t>/g)].map((m) => m[1]).join('');
    if (text.trim() === '') continue;
    for (const size of shape[0].matchAll(/sz="(\d+)"/g)) {
      if (+size[1] / 100 < RESOLVABLE_PT) tooSmall.push(`${text} @ ${+size[1] / 100}pt`);
    }
  }
  assert.deepEqual(tooSmall, [], `${tooSmall.length} unresolvable run(s) on the overview: ${tooSmall.slice(0, 3).join('; ')}`);

  // and the names it gave up are on the readable slides
  assert.ok(deck.parts.join(' ').includes('Azure Kubernetes Service'), 'the detail slides carry the names');
});

test('a tile that gives up its name still says something', async () => {
  // Dropping unreadable type is only an improvement while the tile keeps
  // meaning something. Under Node no icon ever rasterises, so a tile that has
  // given up its name and has no icon to fall back on would be drawn as an
  // empty grey box — which says strictly less than type that is merely small,
  // and which the font-size rule scores perfectly because there is no type
  // left to measure.
  const { nodes, edges } = wideDiagram(72);
  const deck = await buildDeck(nodes, edges);
  assert.ok(deck.parts.length > 1, 'the fixture must tile for there to be an overview at all');
  const overview = deck.slides[deck.slides.length - deck.parts.length - 1];

  const named = new Set(
    [...overview.matchAll(/name="service-label-([^"]+)"/g)].map((m) => m[1]),
  );
  const iconed = new Set([...overview.matchAll(/name="icon-([^"]+)"/g)].map((m) => m[1]));
  const tiles = [...overview.matchAll(/name="service-((?!label-|meta-)[^"]+)"/g)].map((m) => m[1]);
  assert.ok(tiles.length > 0, 'the overview draws tiles');
  const blank = tiles.filter((id) => !named.has(id) && !iconed.has(id));
  assert.deepEqual(blank, [], `${blank.length} of ${tiles.length} overview tiles carry neither a name nor an icon`);

  // and what it does say is still above the floor
  for (const shape of overview.matchAll(/<p:sp>[\s\S]*?<\/p:sp>/g)) {
    const text = [...shape[0].matchAll(/<a:t>([^<]*)<\/a:t>/g)].map((m) => m[1]).join('');
    if (text.trim() === '') continue;
    for (const size of shape[0].matchAll(/sz="(\d+)"/g)) {
      assert.ok(+size[1] / 100 >= 6, `"${text}" drawn at ${+size[1] / 100}pt on the overview`);
    }
  }
});


/**
 * A fan of parallel hops is drawn as a rigid ladder: one offset moves every
 * rung together. On a crowded drawing there is sometimes no offset at all that
 * puts every rung nearer its own arrow than anybody else's — not a position the
 * search can improve on, but proof that no honest position exists for an object
 * that shape. The escape is the one the Architecture Center itself uses for a
 * bundle of parallel flows: number the arrows and let the workflow list carry
 * the sentences.
 *
 * This used to be gated on the fan being at least five deep, so a shallow fan
 * with nowhere honest to stand simply shipped a rung parked beside a stranger's
 * hop — read, believed, and wrong.
 */
test('a ladder with nowhere honest to stand becomes numbers rather than lie', async () => {
  const nodes: Node[] = [];
  for (let r = 0; r < 3; r += 1) {
    for (let c = 0; c < 3; c += 1) nodes.push(service(`g${r}${c}`, `Azure Service ${r}${c}`, c * 300, r * 200));
  }
  const wording = '注文ドキュメントを Cosmos DB に書き込みます';
  const edges: Edge[] = [];
  let step = 0;
  for (let r = 0; r < 3; r += 1) {
    for (let c = 0; c + 1 < 3; c += 1) {
      step += 1;
      edges.push({ id: `h${r}${c}`, source: `g${r}${c}`, target: `g${r}${c + 1}`, label: wording, data: { stepNumber: step, stepDescription: `手順 ${step}` } } as Edge);
    }
  }
  for (let r = 0; r + 1 < 3; r += 1) {
    for (let c = 0; c < 3; c += 1) {
      step += 1;
      edges.push({ id: `v${r}${c}`, source: `g${r}${c}`, target: `g${r + 1}${c}`, label: wording, data: { stepNumber: step, stepDescription: `手順 ${step}` } } as Edge);
    }
  }
  for (let i = 0; i < 3; i += 1) {
    step += 1;
    edges.push({ id: `f${i}`, source: 'g11', target: 'g12', label: `${wording} ${i}`, data: { stepNumber: step, stepDescription: `手順 ${step}` } } as Edge);
  }

  const deck = await buildDeck(nodes, edges);
  const slide = deck.slides[0];

  // Every drawn chip has to read as belonging to the arrow it names, measured
  // the same way a reader would: against the nearest arrow of any other hop.
  const shapes = [...slide.matchAll(/<p:sp>[\s\S]*?<\/p:sp>/g)].map((m) => m[0]);
  const box = (xml: string): { x: number; y: number; w: number; h: number } | null => {
    const off = /<a:off x="(-?\d+)" y="(-?\d+)"\/>/.exec(xml);
    const ext = /<a:ext cx="(\d+)" cy="(\d+)"\/>/.exec(xml);
    return off && ext
      ? { x: +off[1] / 914400, y: +off[2] / 914400, w: +ext[1] / 914400, h: +ext[2] / 914400 }
      : null;
  };
  const named = shapes
    .map((xml) => ({ name: /name="([^"]+)"/.exec(xml)?.[1] ?? '', rect: box(xml) }))
    .filter((s): s is { name: string; rect: NonNullable<ReturnType<typeof box>> } => s.rect !== null);
  const arrows = named.filter((s) => s.name.startsWith('connector-') && !/^connector-(label|step)-/.test(s.name));
  assert.ok(arrows.length > 0, 'the fixture draws arrows');

  // Distance from a point to an arrow's bounding segment, which is how the
  // exporter and the audit both measure it.
  const gap = (r: { x: number; y: number; w: number; h: number }, p: { x: number; y: number }): number => {
    const ax = r.x;
    const ay = r.y;
    const bx = r.x + r.w;
    const by = r.y + r.h;
    const dx = bx - ax;
    const dy = by - ay;
    const len = dx * dx + dy * dy;
    const t = len === 0 ? 0 : Math.max(0, Math.min(1, ((p.x - ax) * dx + (p.y - ay) * dy) / len));
    return Math.hypot(p.x - (ax + t * dx), p.y - (ay + t * dy));
  };

  let checked = 0;
  for (const chip of named.filter((s) => s.name.startsWith('connector-label-'))) {
    const own = arrows.find((a) => a.name === `connector-${chip.name.replace('connector-label-', '')}`);
    if (!own) continue;
    const at = { x: chip.rect.x + chip.rect.w / 2, y: chip.rect.y + chip.rect.h / 2 };
    const mine = gap(own.rect, at);
    const nearest = arrows.reduce((best, a) => (gap(a.rect, at) < gap(best.rect, at) ? a : best), arrows[0]);
    checked += 1;
    assert.ok(
      nearest.name === own.name || gap(nearest.rect, at) >= mine - 0.25,
      `${chip.name} is ${gap(nearest.rect, at).toFixed(2)}in from ${nearest.name} but ${mine.toFixed(2)}in from its own arrow`,
    );
  }
  assert.ok(checked > 0, 'the fixture must draw chips for the rule to mean anything');

  // Muting is only honest if the sentence survives somewhere the reader can
  // reach, so whatever the fan gave up has to appear in the deck's text.
  const text = deck.slides.join('');
  for (let i = 0; i < 3; i += 1) {
    assert.ok(text.includes(`${wording} ${i}`.slice(0, 12)), `the wording of fan member ${i} left the deck entirely`);
  }
});

test('a name the tile can hold is not cut short', async () => {
  // The cap used to be a flat 40 cells regardless of the tile, so a roomy tile
  // showed "Azure Database for PostgreSQL フレキシ…" and the rest of the name
  // was written down nowhere at all. A tile is allowed to clip a name only
  // when the tile genuinely cannot hold it.
  const names = [
    'Azure Kubernetes Service 本番クラスター',
    'Azure Database for PostgreSQL フレキシブル サーバー',
    'Azure Container Registry プレミアム',
  ];
  const nodes = names.map((name, i) => service(`s${i}`, name, i * 260, 0));
  const edges = names.slice(1).map((_, i) => (
    { id: `e${i}`, source: `s${i}`, target: `s${i + 1}`, label: '書き込み' } as Edge
  ));
  const deck = await buildDeck(nodes, edges);
  const drawn = deck.slides
    .flatMap((xml) => [...xml.matchAll(/<p:sp>[\s\S]*?<\/p:sp>/g)].map((m) => m[0]))
    .filter((sp) => /name="service-label-/.test(sp))
    .map((sp) => [...sp.matchAll(/<a:t>([\s\S]*?)<\/a:t>/g)].map((t) => t[1]).join(''));
  assert.ok(drawn.length >= names.length, `expected a caption per service, saw ${drawn.length}`);
  for (const name of names) {
    assert.ok(
      drawn.some((text) => text === name),
      `"${name}" is not written in full anywhere; captions were ${JSON.stringify(drawn)}`,
    );
  }
});

/**
 * Splitting has to be able to stop.
 *
 * Both coarsening loops in `planDiagramWindows` break on
 * `scaleOf(c, r) >= legibleScale && scaleOf(next) < legibleScale`, so a
 * `legibleScale` no reachable grid ever attains makes the left conjunct false
 * at every step and the break never fires — the loop coarsens past every grid
 * that reads. That has produced two separate catastrophes from opposite
 * directions: a demand the frame could not meet gave 400 services 49 slides
 * reading "Azure…", and a demand the renderer would not grant gave 60
 * short-tile services 61 slides carrying one tile each.
 *
 * The end-to-end audit only sees the second one when the deck collapses all the
 * way to one tile per slide. Between 40 and 55 authored pixels the deck
 * over-tiles by 24-48% while every window still carries two tiles, and no
 * property of the emitted file separates that from a small correct deck — the
 * distinguishing fact is a counterfactual. So it is asserted here instead, on
 * the function, where it is true by construction or not at all.
 */
test('the legibility target never exceeds what the renderer will draw', () => {
  const NATURAL_PER_IN = 1 / 96;
  const frames = [
    { w: 12.63, h: 5.78 },   // the standard 16:9 window
    { w: 4, h: 2 },          // a frame small enough for its own ceiling to bind
    { w: 40, h: 20 },        // a grown page
  ];
  for (const frame of frames) {
    for (let target = 1; target <= 400; target += 1) {
      const scale = legibleScaleFor(target, frame);
      assert.ok(
        scale <= NATURAL_PER_IN + 1e-12,
        `target ${target}px in a ${frame.w}x${frame.h}in frame asks for ${scale} in/px, `
        + `but the renderer caps every window at ${NATURAL_PER_IN} in/px`,
      );
      assert.ok(scale > 0, `target ${target}px produced a non-positive scale ${scale}`);
    }
  }
});

/**
 * Wrapping is a *count*, not a ratio, and the difference is what puts a row off
 * the page.
 *
 * `ceil(width / box)` assumes text can break anywhere. It can, in CJK and in a
 * single over-long token, and for those the two agree. Latin prose breaks
 * between words and abandons the rest of a line when the next word will not
 * fit, so the ratio is only ever a lower bound — and a table sized from a lower
 * bound is a table measured onto the page that prints below it.
 *
 * Asserted here rather than through an export because no output statistic can
 * separate the two: the emitted file records the row heights the exporter
 * chose, so a rule reading them back agrees with whatever it did.
 */
test('a wrapped line count is never below the break-anywhere ratio', () => {
  const EM = 0.54; // the estimator's average Latin advance
  const samples = [
    'Azure Kubernetes Service aks001contosoplatformprodeastus2 nodepool001systemsurgeeastus2',
    'Azure Database for PostgreSQL Flexible Server (Production, Zone Redundant) 17',
    'contoso-platform-prod-eastus2-aks-system-nodepool-surge',
    '受信した注文イベントを検証し、重複を排除したうえで下流の在庫サービスに書き込みます',
    'Short name',
    '',
  ];
  for (const text of samples) {
    for (const box of [1, 2.5, 3.3, 5, 11.89]) {
      for (const pt of [7, 9, 10, 12]) {
        const lines = wrappedLineCount(text, box, pt);
        assert.ok(Number.isInteger(lines) && lines >= 1, `"${text.slice(0, 20)}" gave ${lines} lines`);
        let units = 0;
        for (const ch of text) {
          units += /[\u2e80-\u9fff\uac00-\ud7af\uff00-\uff60\uffe0-\uffe6]/.test(ch) ? 1 : EM;
        }
        const ratio = Math.max(1, Math.ceil((units * pt / 72) / box));
        assert.ok(
          lines >= ratio,
          `"${text.slice(0, 28)}" at ${pt}pt in ${box}in wraps to ${lines} lines, `
          + `below the ${ratio} that break-anywhere alone requires`,
        );
      }
    }
  }
  // The case the ratio actually gets wrong: three runs each over half the
  // column. Break-anywhere fits 38 characters into two 2in lines at 12pt;
  // breaking between words cannot, because the second run will not fit after
  // the first and the rest of that line is abandoned.
  const tokens = 'aaaaaaaaaaaa bbbbbbbbbbbb cccccccccccc';
  const ratio = Math.max(1, Math.ceil((tokens.length * EM * 12 / 72) / 2.0));
  assert.equal(ratio, 2, 'the fixture must be one the break-anywhere ratio calls two lines');
  assert.equal(
    wrappedLineCount(tokens, 2.0, 12), 3,
    'three runs each over half the column take three lines, not the two the ratio predicts',
  );
});

/**
 * A table row is never budgeted at less than its type plus the insets
 * PowerPoint charges on top of it.
 *
 * None of the deck's tables declares autofit, so `<a:tr h>` is a minimum and
 * every one of them grows to whatever its contents need — text, plus `marT` and
 * `marB`, which pptxgenjs emits at 0.05in each. Omitting the insets is 0.1in a
 * row, and a page of rows of that is a page and a half of table.
 */
test('a table row budgets its cell insets as well as its type', () => {
  const COL_W = [5.2, 3.9, 3.53];
  const INSET_V = 0.1;
  for (const pt of [7, 9, 10, 12]) {
    const single = tableRowHeightIn(['Api', 'Compute', '—'], COL_W, pt);
    assert.ok(
      single >= pt * 1.35 / 72 + INSET_V - 1e-9,
      `a one-line row at ${pt}pt was budgeted ${single}in, below its type plus insets`,
    );
    const long = 'Azure Database for PostgreSQL Flexible Server (Production, Zone Redundant) 17';
    const wrapped = tableRowHeightIn([long, 'Databases', 'Data tier'], COL_W, pt);
    const lines = wrappedLineCount(long, COL_W[0] - 0.2, pt);
    assert.ok(
      wrapped >= lines * pt * 1.35 / 72 + INSET_V - 1e-9,
      `a ${lines}-line row at ${pt}pt was budgeted ${wrapped}in, below its type plus insets`,
    );
    assert.ok(wrapped >= single, 'a wrapped row is never shorter than a single-line row');
  }
});

/**
 * The frame ceiling still has to bind when it is the tighter of the two, or the
 * first catastrophe comes back: `gridFor` returns null the moment the window
 * bleed alone fills the frame, and a null grid sends the planner to a fixed
 * fallback grid the coarsening loops then walk straight past.
 */
test('the legibility target never exceeds what the frame can deliver', () => {
  const BLEED_PX = 18;
  for (const frame of [{ w: 12.63, h: 5.78 }, { w: 2, h: 1 }, { w: 40, h: 20 }]) {
    for (let target = 1; target <= 400; target += 1) {
      const finest = Math.min(frame.w, frame.h) / (BLEED_PX * 2 + target);
      assert.ok(
        legibleScaleFor(target, frame) <= finest + 1e-12,
        `target ${target}px in a ${frame.w}x${frame.h}in frame asks for more than the `
        + `${finest} in/px the frame can hold`,
      );
    }
  }
});

/**
 * The WAF pillar and regional-comparison tables must never silently drop a row.
 *
 * `fitTableRows` shrinks the type first and only then discards rows, and on
 * these two the contents are closed sets — five fixed WAF pillar names against
 * a 2.91in budget, and the shipped region list against 5.32in — so the
 * row-dropping arm should be unreachable. "Should be unreachable" is worth
 * pinning: it is reachable the moment either budget is tightened or either
 * label grows, and the failure is silent by construction.
 */
test('the closed-set tables keep every row they are given', () => {
  const PILLARS = [
    'Reliability', 'Security', 'Cost Optimization',
    'Operational Excellence', 'Performance Efficiency',
  ];
  const pillarRows = PILLARS.map((p) => [p, 'Adequate, with gaps', '72 / 100']);
  const pillarFit = fitTableRows(
    pillarRows,
    ['Pillar', 'Maturity', 'Score'],
    [4.2, 6.6, 1.83],
    2.91,
    12,
  );
  assert.equal(pillarFit.rows, pillarRows.length, 'the pillar table dropped a WAF pillar');

  const regionRows = Array.from({ length: 9 }, (_, i) => [
    `East US ${i} (Zone redundant)`, '$12,345', '+4.2%', 'Comparable to baseline',
  ]);
  const regionFit = fitTableRows(
    regionRows,
    ['Region', 'Monthly', 'Delta', 'Notes'],
    [3.4, 2.4, 2.0, 4.8],
    5.32,
    12,
  );
  assert.equal(regionFit.rows, regionRows.length, 'the region table dropped a region');
});
