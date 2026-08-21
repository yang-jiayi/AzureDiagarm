import test from 'node:test';
import assert from 'node:assert/strict';
import JSZip from 'jszip';
import type { Edge, Node } from 'reactflow';
import { buildDiagramSlidePptx } from '../src/services/pptxExporter.ts';
import { nativizeSlideXml, nativizePackage } from '../src/services/pptxNativeShapes.ts';
import { buildVsdxPackage } from '../src/services/visioVsdxExporter.ts';

/**
 * A drawing made of rectangles and lines carries no words about itself.
 *
 * The tiles say their own names because they hold text, but an arrow holds
 * none, a zone's title is a separate caption, and nothing in the file records
 * which zone a service sits in or what the numbered badges refer to. A reader
 * using a screen reader, PowerPoint's accessibility checker, or a Visio report
 * therefore gets a list of service names in z-order and nothing else — none of
 * the architecture. These tests pin the descriptions that fix that, and pin the
 * rule that a description may only state what the drawing actually shows.
 */

const PIXEL_PNG =
  'data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg==';

function service(id: string, label: string, x: number, y: number, parentNode?: string): Node {
  return {
    id,
    type: 'azureNode',
    position: { x, y },
    width: 150,
    height: 75,
    ...(parentNode ? { parentNode } : {}),
    data: { label, serviceName: label, sku: 'Premium', region: 'japaneast' },
  } as Node;
}

function zone(id: string, label: string, x: number, y: number, w: number, h: number): Node {
  return {
    id,
    type: 'groupNode',
    position: { x, y },
    style: { width: w, height: h },
    data: { label },
  } as Node;
}

const NODES: Node[] = [
  zone('z1', 'Landing zone', 0, 0, 700, 400),
  service('web', 'Azure Front Door', 60, 90, 'z1'),
  service('app', 'App Service', 400, 90, 'z1'),
  service('log', 'Azure Monitor', 900, 500),
];

const EDGES: Edge[] = [
  { id: 'e1', source: 'web', target: 'app', label: 'HTTPS', data: { stepNumber: 1 } } as Edge,
  { id: 'e2', source: 'app', target: 'log', label: 'Telemetry' } as Edge,
];

async function deck(): Promise<JSZip> {
  const pptx = await buildDiagramSlidePptx(PIXEL_PNG, {
    diagramName: 'Contoso Platform',
    author: 'Tester',
    date: '2026-08-10',
    isDarkMode: false,
    diagram: { nodes: NODES, edges: EDGES },
  });
  const zip = await JSZip.loadAsync((await pptx.write({ outputType: 'nodebuffer' })) as Buffer);
  return nativizePackage(zip);
}

async function slideXml(zip: JSZip): Promise<string[]> {
  const names = Object.keys(zip.files)
    .filter((name) => /^ppt\/slides\/slide\d+\.xml$/.test(name))
    .sort((a, b) => +a.replace(/\D/g, '') - +b.replace(/\D/g, ''));
  return Promise.all(names.map((name) => zip.file(name)!.async('string')));
}

function descriptions(xml: string): Map<string, string> {
  const found = new Map<string, string>();
  for (const match of xml.matchAll(/<p:cNvPr id="\d+" name="([^"]*)" descr="([^"]*)"/g)) {
    found.set(match[1], match[2]);
  }
  return found;
}

test('every tile and every arrow says what it is', async () => {
  const slides = await slideXml(await deck());
  let tiles = 0;
  let arrows = 0;
  for (const xml of slides) {
    const alt = descriptions(xml);
    for (const [name, text] of alt) {
      if (name.startsWith('node-') || name.startsWith('service-')) {
        tiles += 1;
        assert.match(text, /^Service: /, `${name} must describe itself as a service`);
      }
      if (name.startsWith('connector-')) {
        arrows += 1;
        assert.match(text, /Connection/, `${name} must describe itself as a connection`);
      }
    }
    for (const shape of xml.matchAll(/name="(connector-e\d+)"/g)) {
      assert.ok(alt.has(shape[1]), `${shape[1]} carries no description`);
    }
  }
  assert.ok(tiles > 0, 'the fixture must draw tiles');
  assert.ok(arrows > 0, 'the fixture must draw arrows');
});

test('an arrow names the services it actually joins, and never any other', async () => {
  const slides = await slideXml(await deck());
  const names = NODES.filter((n) => n.type === 'azureNode').map((n) => String((n.data as { label: string }).label));
  let named = 0;
  for (const xml of slides) {
    for (const [name, text] of descriptions(xml)) {
      if (!name.startsWith('connector-')) continue;
      const pair = /Connection from (.+?) to ([^:]+)/.exec(text);
      if (!pair) continue;
      named += 1;
      // A name in the sentence has to be a service in the diagram: the bounding
      // box of a bent hop can land on an unrelated tile, and naming that tile
      // would be an invented fact.
      assert.ok(names.includes(pair[1].trim()), `"${pair[1]}" is not a service in this diagram`);
      assert.ok(names.includes(pair[2].trim()), `"${pair[2]}" is not a service in this diagram`);
      assert.notEqual(pair[1].trim(), pair[2].trim(), 'an arrow cannot join a service to itself');
    }
  }
  assert.ok(named > 0, 'a straight hop between two tiles must resolve both endpoints');
});

test('a description that already exists is not overwritten', async () => {
  const slides = await slideXml(await deck());
  const icons = slides.flatMap((xml) => [...xml.matchAll(/name="(icon-[^"]*)" descr="([^"]*)"/g)]);
  for (const icon of icons) {
    assert.equal(/^Service: /.test(icon[2]), false, 'the icon keeps the alt text it was given');
  }
});

test('the notes page lists what the slide draws', async () => {
  const zip = await deck();
  const notes = Object.keys(zip.files).filter((name) => /^ppt\/notesSlides\/notesSlide\d+\.xml$/.test(name));
  assert.ok(notes.length > 0, 'the deck must carry notes pages');
  const text = (await Promise.all(notes.map((name) => zip.file(name)!.async('string')))).join('\n');
  assert.match(text, /Contoso Platform/);
  assert.match(text, /Services on this slide/);
  assert.match(text, /Azure Front Door/);
  assert.match(text, /Landing zone/);
});

test('a service tile links to the documentation without changing how it looks', async () => {
  const zip = await deck();
  const slides = await slideXml(zip);
  const withLink = slides.filter((xml) => /<a:hlinkClick/.test(xml));
  assert.ok(withLink.length > 0, 'tiles must carry a hyperlink');
  const rels = await Promise.all(
    Object.keys(zip.files)
      .filter((name) => /^ppt\/slides\/_rels\/slide\d+\.xml\.rels$/.test(name))
      .map((name) => zip.file(name)!.async('string')),
  );
  const targets = rels.join('\n');
  assert.match(targets, /learn\.microsoft\.com\/search\/\?terms=/);
  // Every relationship the slide points at has to exist, or PowerPoint refuses
  // to open the package.
  for (const [index, xml] of slides.entries()) {
    for (const link of xml.matchAll(/<a:hlinkClick r:id="(rId\d+)"/g)) {
      assert.ok(rels[index]?.includes(`Id="${link[1]}"`), `slide ${index + 1} points at a missing ${link[1]}`);
    }
  }
  // A hyperlink on a text run would underline it and repaint it in the theme's
  // link colour, which would change what the slide looks like.
  for (const xml of slides) {
    for (const run of xml.matchAll(/<a:rPr[^>]*>[\s\S]*?<\/a:rPr>/g)) {
      assert.equal(/<a:hlinkClick/.test(run[0]), false, 'no text run may carry a hyperlink');
    }
  }
});

test('the description survives being folded into a group', async () => {
  const zip = await deck();
  for (const xml of await slideXml(zip)) {
    for (const group of xml.matchAll(/<p:grpSp>[\s\S]*?<\/p:grpSp>/g)) {
      const alt = descriptions(group[0]);
      const owner = /name="(node-[^"]*)" descr="([^"]*)"/.exec(group[0]);
      assert.ok(owner, 'a grouped tile must describe itself');
      assert.match(owner[2], /^Service: /);
      assert.ok(alt.size > 0);
    }
  }
});

test('a Visio service records the zone it is drawn in', async () => {
  const pkg = await buildVsdxPackage(NODES, EDGES, 'Contoso');
  const page = pkg.parts.find((part) => part.path === 'visio/pages/page1.xml');
  assert.ok(page, 'page1.xml must be present');
  const xml = page.data as string;
  assert.ok(xml.includes('N="Zone"'), 'shape data must expose the zone');
  const rows = [...xml.matchAll(/<Row N="Zone"[^>]*>([\s\S]*?)<\/Row>/g)].map((m) => m[1]);
  assert.equal(rows.length, 2, 'only the two services inside the zone get the row');
  for (const row of rows) assert.match(row, /Landing zone/);
});

test('a slide with nothing to describe is returned untouched', () => {
  const xml = '<p:sld><p:cSld><p:spTree/></p:cSld></p:sld>';
  assert.equal(nativizeSlideXml(xml), xml);
});

test('a control character in a name never reaches the hyperlink attribute', async () => {
  // U+0001 is illegal in XML 1.0 and no escaping helps, so one in an attribute
  // makes the whole deck unopenable. Service names are pasted, imported and
  // model-generated, so this is not hypothetical.
  const pptx = await buildDiagramSlidePptx(PIXEL_PNG, {
    diagramName: 'Contoso',
    author: 'Tester',
    date: '2026-08-10',
    isDarkMode: false,
    diagram: { nodes: [service('web', 'Azure\u0001 Front\u000b Door', 60, 90)], edges: [] },
  });
  const zip = await JSZip.loadAsync((await pptx.write({ outputType: 'nodebuffer' })) as Buffer);
  for (const xml of await slideXml(await nativizePackage(zip))) {
    assert.equal(/[\u0000-\u0008\u000b\u000c\u000e-\u001f]/.test(xml), false, 'the slide carries an illegal code point');
    assert.match(xml, /tooltip="[^"]*Front Door/);
  }
});
