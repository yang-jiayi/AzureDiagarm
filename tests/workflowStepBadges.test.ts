import test from 'node:test';
import assert from 'node:assert/strict';
import JSZip from 'jszip';
import type { Edge, Node } from 'reactflow';
import { buildDiagramSlidePptx, buildArchitectureDeckPptx } from '../src/services/pptxExporter.ts';
import { buildVsdxPackage } from '../src/services/visioVsdxExporter.ts';
import { exportToDrawio } from '../src/services/drawioExporter.ts';
import { buildInteractiveDiagramHtml } from '../src/services/htmlDiagramExporter.ts';

/**
 * Azure Architecture Center reference diagrams number the arrows and repeat
 * those numbers in the "Workflow" prose. The app renders those badges on the
 * canvas, so an export that silently drops them breaks the correspondence
 * between the picture and the description the reader is given.
 *
 * These tests assert every export format emits a badge for a numbered edge and
 * — just as importantly — emits none for an unnumbered one, so the guard fails
 * if either the numbering or the suppression regresses.
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

const NODES: Node[] = [
  service('a', 'Azure Front Door', 0, 0),
  service('b', 'Azure Kubernetes Service', 500, 0),
  service('c', 'Azure SQL Database', 1000, 0),
];

/** Two hops; only the first carries a workflow step number. */
function edges(): Edge[] {
  return [
    { id: 'e1', source: 'a', target: 'b', label: 'HTTPS', data: { stepNumber: 1 } } as Edge,
    { id: 'e2', source: 'b', target: 'c', label: 'SQL' } as Edge,
  ];
}

async function pptxSlideXml(): Promise<string> {
  const pptx = await buildDiagramSlidePptx(PIXEL_PNG, {
    diagramName: 'Numbered flow',
    author: 'Tester',
    date: '2026-08-10',
    isDarkMode: false,
    diagram: { nodes: NODES, edges: edges() },
  });
  const buffer = (await pptx.write({ outputType: 'nodebuffer' })) as Buffer;
  const zip = await JSZip.loadAsync(buffer);
  return zip.file('ppt/slides/slide1.xml')!.async('string');
}

test('PowerPoint draws one editable badge shape per numbered connector', async () => {
  const xml = await pptxSlideXml();
  const badges = [...xml.matchAll(/name="connector-step-([^"]+)"/g)].map((m) => m[1]);
  assert.deepEqual(badges, ['e1'], 'exactly the numbered edge gets a badge');
  // The badge has to be a real shape, not a picture, or PowerPoint cannot edit it.
  const badgeShape = /<p:sp>(?:(?!<\/p:sp>)[\s\S])*name="connector-step-e1"[\s\S]*?<\/p:sp>/.exec(xml);
  assert.ok(badgeShape, 'the badge is a drawn shape');
  assert.match(badgeShape![0], /<a:t>1<\/a:t>/, 'the badge carries the step number as text');
});

test('Visio draws one badge shape per numbered connector', async () => {
  const { parts } = await buildVsdxPackage(NODES, edges(), 'Numbered flow');
  const page = parts.find((part) => part.path === 'visio/pages/page1.xml')!.data as string;
  const badges = [...page.matchAll(/NameU="StepBadge\.\d+"/g)];
  assert.equal(badges.length, 1, 'exactly the numbered edge gets a badge');
  const shape = /<Shape [^>]*NameU="StepBadge\.\d+"[\s\S]*?<\/Shape>/.exec(page)!;
  assert.match(shape[0], /<Text>1<\/Text>/, 'the badge carries the step number');
  assert.match(shape[0], /<Row T="Ellipse"/, 'the badge is a real ellipse, not a text box');
});

test('Draw.io emits an anchored badge cell per numbered connector', async () => {
  const xml = await exportToDrawio(NODES, edges(), 'Numbered flow');
  const badges = [...xml.matchAll(/id="[^"]*-step"/g)];
  assert.equal(badges.length, 1, 'exactly the numbered edge gets a badge');
  const cell = /<mxCell[^>]*id="[^"]*-step"[^>]*>/.exec(xml)!;
  assert.match(cell[0], /value="1"/, 'the badge carries the step number');
  assert.match(cell[0], /ellipse/, 'the badge is drawn as an ellipse');
});

test('the interactive HTML export carries workflow numbering into its renderer', async () => {
  const html = (await buildInteractiveDiagramHtml(NODES, edges(), 'Numbered flow'))!;
  assert.ok(html, 'the export produced a document');

  // The HTML draws its SVG at runtime, so the durable contract is the embedded
  // layout: the numbered edge must carry stepNumber and the plain one must not.
  const match = /const layout = (.+?);\n\nlet scale/s.exec(html);
  assert.ok(match, 'layout JSON is embedded in the HTML');
  const layout = JSON.parse(
    match[1].replace(/\\u003c/g, '<').replace(/\\u2028/g, '\u2028').replace(/\\u2029/g, '\u2029'),
  ) as { edges: Array<{ id: string; stepNumber?: number }> };
  const numbered = layout.edges.filter((edge) => edge.stepNumber !== undefined);
  assert.deepEqual(numbered.map((edge) => [edge.id, edge.stepNumber]), [['e1', 1]]);

  assert.match(html, /\.edge-step\s*\{/, 'the badge is styled');
  assert.match(html, /classList\.add\('edge-step'\)/, 'the renderer draws the badge');
});

test('no format invents badges when the diagram has no workflow numbering', async () => {
  const plain: Edge[] = [{ id: 'e1', source: 'a', target: 'b', label: 'HTTPS' } as Edge];

  const pptx = await buildDiagramSlidePptx(PIXEL_PNG, {
    diagramName: 'Plain', author: 'Tester', date: '2026-08-10', isDarkMode: false,
    diagram: { nodes: NODES, edges: plain },
  });
  const zip = await JSZip.loadAsync((await pptx.write({ outputType: 'nodebuffer' })) as Buffer);
  assert.doesNotMatch(await zip.file('ppt/slides/slide1.xml')!.async('string'), /connector-step-/);

  const { parts } = await buildVsdxPackage(NODES, plain, 'Plain');
  assert.doesNotMatch(parts.find((p) => p.path === 'visio/pages/page1.xml')!.data as string, /StepBadge\./);

  assert.doesNotMatch(await exportToDrawio(NODES, plain, 'Plain'), /id="[^"]*-step"/);
  assert.doesNotMatch((await buildInteractiveDiagramHtml(NODES, plain, 'Plain'))!, /"stepNumber"/);
});

/**
 * A numbered arrow that points at nothing is worse than no number at all, so
 * the deck must carry the matching numbered prose the Architecture Center
 * pairs with every reference diagram.
 */
async function deckSlides(workflow: unknown): Promise<string[]> {
  const pptx = await buildArchitectureDeckPptx(PIXEL_PNG, {
    diagramName: 'Numbered flow',
    author: 'Tester',
    date: '2026-08-10',
    isDarkMode: false,
    services: [{ name: 'Azure Front Door', category: 'Networking' }],
    workflow: workflow as never,
    diagram: { nodes: NODES, edges: edges() },
  });
  const zip = await JSZip.loadAsync((await pptx.write({ outputType: 'nodebuffer' })) as Buffer);
  const names = Object.keys(zip.files).filter((name) => /^ppt\/slides\/slide\d+\.xml$/.test(name));
  return Promise.all(names.sort().map((name) => zip.file(name)!.async('string')));
}

test('the deck carries a numbered workflow slide matching the arrow callouts', async () => {
  const slides = await deckSlides([
    { step: 1, description: 'The client calls Front Door over HTTPS.', services: ['a', 'b'] },
    { step: 2, description: 'Front Door forwards the request to AKS.', services: ['b', 'c'] },
  ]);
  const workflowSlide = slides.find((xml) => /<a:t>Workflow<\/a:t>/.test(xml));
  assert.ok(workflowSlide, 'the deck has a Workflow slide');
  assert.match(workflowSlide!, /<a:t>2 steps<\/a:t>/, 'the slide states how many steps there are');
  assert.match(workflowSlide!, /The client calls Front Door over HTTPS\./);
  assert.match(workflowSlide!, /Front Door forwards the request to AKS\./);
  // Each step needs its own numbered badge so it can be tied back to an arrow.
  for (const number of ['1', '2']) {
    assert.ok(new RegExp(`<a:t>${number}</a:t>`).test(workflowSlide!), `step ${number} is numbered`);
  }
});

test('the workflow slide sorts by step and is omitted when there is no workflow', async () => {
  const outOfOrder = await deckSlides([
    { step: 2, description: 'Second thing happens.', services: ['b', 'c'] },
    { step: 1, description: 'First thing happens.', services: ['a', 'b'] },
  ]);
  const slide = outOfOrder.find((xml) => /<a:t>Workflow<\/a:t>/.test(xml))!;
  assert.ok(
    slide.indexOf('First thing happens.') < slide.indexOf('Second thing happens.'),
    'steps are laid out in step order, not array order',
  );

  for (const empty of [null, [], [{ step: 1, description: '' }]]) {
    const slides = await deckSlides(empty);
    assert.ok(!slides.some((xml) => /<a:t>Workflow<\/a:t>/.test(xml)),
      `no empty Workflow slide for ${JSON.stringify(empty)}`);
  }
});

/**
 * A tall wrapped label used to be drawn straight through the badge, because
 * the badge offset was a fixed fraction of the badge diameter rather than the
 * chip's real height. Long CJK labels wrap to several lines, so this is the
 * common case, not an edge case.
 */
test('the PowerPoint badge clears even a tall wrapped label', async () => {
  const EMU = 914400;
  const long = 'HTTPS 経由でトークン検証を実施し、マネージド ID で認可します';
  const pptx = await buildDiagramSlidePptx(PIXEL_PNG, {
    diagramName: 'Wrapped label', author: 'Tester', date: '2026-08-10', isDarkMode: false,
    diagram: {
      nodes: NODES,
      edges: [{ id: 'e1', source: 'a', target: 'b', label: long, data: { stepNumber: 1 } } as Edge],
    },
  });
  const zip = await JSZip.loadAsync((await pptx.write({ outputType: 'nodebuffer' })) as Buffer);
  const xml = await zip.file('ppt/slides/slide1.xml')!.async('string');

  const boxOf = (name: string) => {
    const body = new RegExp(`<p:sp>(?:(?!</p:sp>)[\\s\\S])*name="${name}"[\\s\\S]*?</p:sp>`).exec(xml);
    assert.ok(body, `${name} exists`);
    const off = /<a:off x="(-?\d+)" y="(-?\d+)"\/>/.exec(body![0])!;
    const ext = /<a:ext cx="(\d+)" cy="(\d+)"\/>/.exec(body![0])!;
    return { x: +off[1] / EMU, y: +off[2] / EMU, w: +ext[1] / EMU, h: +ext[2] / EMU };
  };
  const chip = boxOf('connector-label-e1');
  const badge = boxOf('connector-step-e1');

  const overlapH = Math.min(chip.y + chip.h, badge.y + badge.h) - Math.max(chip.y, badge.y);
  const overlapW = Math.min(chip.x + chip.w, badge.x + badge.w) - Math.max(chip.x, badge.x);
  const overlap = overlapW > 0 && overlapH > 0 ? overlapW * overlapH : 0;
  assert.equal(overlap, 0, 'the badge does not overlap the label chip');
  assert.ok(badge.y >= chip.y + chip.h, 'the badge sits below the whole chip');
});
