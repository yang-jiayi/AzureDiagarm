import test from 'node:test';
import assert from 'node:assert/strict';
import JSZip from 'jszip';
import type { Edge, Node } from 'reactflow';
import { buildDiagramSlidePptx } from '../src/services/pptxExporter.ts';
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
  for (const name of Object.keys(zip.files).filter((n) => /^ppt\/slides\/slide\d+\.xml$/.test(n))) {
    const xml = await zip.file(name)!.async('string');
    const at = xml.indexOf('name="zone-zFar"');
    if (at < 0) continue;
    const m = /<a:off x="(-?[\d.]+)" y="(-?[\d.]+)"\s*\/>\s*<a:ext cx="([\d.]+)" cy="([\d.]+)"/.exec(xml.slice(at, at + 400))!;
    const [x, y, w, h] = [+m[1] / EMU, +m[2] / EMU, +m[3] / EMU, +m[4] / EMU];
    seen += 1;
    assert.ok(w > 0.2 && h > 0.2, `trimmed zone collapsed to ${w.toFixed(3)}x${h.toFixed(3)}in on ${name}`);
    assert.ok(
      x >= -0.01 && y >= -0.01 && x + w <= pageW + 0.01 && y + h <= pageH + 0.01,
      `trimmed zone sits at ${x.toFixed(2)},${y.toFixed(2)} ${w.toFixed(2)}x${h.toFixed(2)}in off a ${pageW.toFixed(2)}x${pageH.toFixed(2)}in ${name}`,
    );
  }
  assert.equal(seen, 1, 'the outlier zone was not drawn at all');
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
