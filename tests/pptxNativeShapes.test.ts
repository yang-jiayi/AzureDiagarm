import test from 'node:test';
import assert from 'node:assert/strict';
import JSZip from 'jszip';
import type { Edge, Node } from 'reactflow';
import { buildDiagramSlidePptx } from '../src/services/pptxExporter.ts';
import { nativizeSlideXml } from '../src/services/pptxNativeShapes.ts';

/**
 * pptxgenjs can only emit `<p:sp>`, which is enough to draw an architecture
 * but not to edit one: every arrow is a loose line that stays behind when the
 * service it points at is dragged, and every service name is a separate text
 * box that stays behind when its tile is dragged. That is why an exported deck
 * had to be redrawn by hand to be usable. These tests pin the repair.
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

async function slidesOf(nodes: Node[], edges: Edge[]): Promise<string[]> {
  const pptx = await buildDiagramSlidePptx(PIXEL_PNG, {
    diagramName: 'Contoso Platform',
    author: 'Tester',
    date: '2026-08-10',
    isDarkMode: false,
    diagram: { nodes, edges },
  });
  const zip = await JSZip.loadAsync((await pptx.write({ outputType: 'nodebuffer' })) as Buffer);
  const names = Object.keys(zip.files)
    .filter((name) => /^ppt\/slides\/slide\d+\.xml$/.test(name))
    .sort((a, b) => +a.replace(/\D/g, '') - +b.replace(/\D/g, ''));
  return Promise.all(names.map((name) => zip.file(name)!.async('string')));
}

function chain(count: number): { nodes: Node[]; edges: Edge[] } {
  const nodes: Node[] = [];
  const edges: Edge[] = [];
  for (let i = 0; i < count; i += 1) {
    nodes.push(service(`n${i}`, `Azure Service ${i}`, (i % 4) * 300, Math.floor(i / 4) * 220));
    if (i > 0) edges.push({ id: `e${i}`, source: `n${i - 1}`, target: `n${i}`, label: 'HTTPS' } as Edge);
  }
  return { nodes, edges };
}

test('every arrow becomes a connector PowerPoint can keep attached', async () => {
  const { nodes, edges } = chain(8);
  const slides = await slidesOf(nodes, edges);
  let drawn = 0;
  let connectors = 0;
  let glued = 0;
  for (const xml of slides) {
    const out = nativizeSlideXml(xml);
    drawn += [...xml.matchAll(/<p:sp>[\s\S]*?<\/p:sp>/g)].filter((m) => /name="connector-e\d+"/.test(m[0])).length;
    connectors += (out.match(/<p:cxnSp>/g) ?? []).length;
    glued += (out.match(/<a:stCxn /g) ?? []).length;

    // Glue that points at a shape which is not on the slide makes PowerPoint
    // drop the arrow on open, which is worse than not gluing at all.
    const ids = new Set([...out.matchAll(/<p:cNvPr id="(\d+)"/g)].map((m) => m[1]));
    for (const cxn of out.matchAll(/<a:(?:st|end)Cxn id="(\d+)" idx="(\d+)"\/>/g)) {
      assert.ok(ids.has(cxn[1]), `glued to shape id ${cxn[1]}, which is not on the slide`);
      assert.ok(+cxn[2] <= 3, `glued to connection site ${cxn[2]}, which a rectangle does not have`);
    }
  }
  assert.ok(drawn > 0, 'the fixture must draw arrows');
  assert.equal(connectors, drawn, 'every drawn arrow converts to a connector');
  assert.ok(glued > 0, 'and the ones that meet a tile squarely are glued to it');
});

test('a service name moves into the tile it names, unchanged', async () => {
  const { nodes, edges } = chain(6);
  const slides = await slidesOf(nodes, edges);
  for (const xml of slides) {
    const out = nativizeSlideXml(xml);
    const names = [...xml.matchAll(/name="service-label-([^"]+)"/g)].map((m) => m[1]);
    if (names.length === 0) continue;
    assert.equal(
      (out.match(/name="service-label-/g) ?? []).length,
      0,
      'no service name is left floating beside its tile',
    );
    for (const id of names) {
      const tile = new RegExp(`<p:cNvPr id="\\d+" name="service-${id}">[\\s\\S]*?</p:sp>`).exec(out)?.[0];
      assert.ok(tile, `tile service-${id} survives`);
      assert.ok(/<p:txBody>/.test(tile), `tile service-${id} carries its own name`);
    }
    // and the wording itself is untouched
    for (const word of xml.matchAll(/<a:t>(Azure Service \d+)<\/a:t>/g)) {
      assert.ok(out.includes(word[0]), `"${word[1]}" survives the conversion`);
    }
  }
});

test('a tile is grouped with the icon drawn on it', () => {
  // Icons only rasterise in a browser, so the group path cannot be reached
  // from a generated deck under Node. It is the path that stops a dragged
  // service leaving its icon behind, so it is pinned directly.
  const slide =
    '<p:spTree>' +
    '<p:sp><p:nvSpPr><p:cNvPr id="7" name="service-a"></p:cNvPr><p:cNvSpPr/><p:nvPr/></p:nvSpPr>' +
    '<p:spPr><a:xfrm><a:off x="1000000" y="2000000"/><a:ext cx="900000" cy="600000"/></a:xfrm>' +
    '<a:prstGeom prst="roundRect"><a:avLst/></a:prstGeom></p:spPr></p:sp>' +
    '<p:sp><p:nvSpPr><p:cNvPr id="8" name="service-label-a"></p:cNvPr><p:cNvSpPr/><p:nvPr/></p:nvSpPr>' +
    '<p:spPr><a:xfrm><a:off x="1030000" y="2400000"/><a:ext cx="840000" cy="150000"/></a:xfrm>' +
    '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr>' +
    '<p:txBody><a:bodyPr wrap="square" lIns="0" tIns="0" rIns="0" bIns="0"/><a:lstStyle/>' +
    '<a:p><a:r><a:rPr sz="900"/><a:t>Azure Front Door</a:t></a:r></a:p></p:txBody></p:sp>' +
    '<p:pic><p:nvPicPr><p:cNvPr id="9" name="icon-a"></p:cNvPr><p:cNvPicPr/><p:nvPr/></p:nvPicPr>' +
    '<p:blipFill><a:blip r:embed="rId2"/></p:blipFill>' +
    '<p:spPr><a:xfrm><a:off x="1350000" y="2060000"/><a:ext cx="200000" cy="200000"/></a:xfrm></p:spPr></p:pic>' +
    '</p:spTree>';

  const out = nativizeSlideXml(slide);
  const group = /<p:grpSp>[\s\S]*<\/p:grpSp>/.exec(out)?.[0];
  assert.ok(group, 'the tile is wrapped in a group');
  assert.ok(group.includes('name="service-a"'), 'the group holds the tile');
  assert.ok(group.includes('name="icon-a"'), 'and the icon drawn on it');
  assert.equal((out.match(/name="icon-a"/g) ?? []).length, 1, 'the icon is moved, not copied');
  assert.ok(out.includes('<a:t>Azure Front Door</a:t>'), 'the name survives');
  assert.equal((out.match(/name="service-label-a"/g) ?? []).length, 0, 'and no longer floats');

  // Child coordinates stay absolute, which is what lets the shapes move in
  // with no rewriting at all.
  const frame = /<p:grpSpPr><a:xfrm>([\s\S]*?)<\/a:xfrm>/.exec(out)?.[1] ?? '';
  assert.ok(frame.includes('<a:off x="1000000" y="2000000"/>'), 'group origin matches the tile');
  assert.ok(frame.includes('<a:chOff x="1000000" y="2000000"/>'), 'child origin matches it exactly');
  assert.ok(frame.includes('<a:chExt cx="900000" cy="600000"/>'), 'child extent matches it exactly');
});

test('a caption that sits outside its shape is left where it is', () => {
  // A zone title is drawn above its box, not inside it. Folding it in would
  // move it, so the conversion must decline.
  const slide =
    '<p:spTree>' +
    '<p:sp><p:nvSpPr><p:cNvPr id="7" name="zone-z"></p:cNvPr><p:cNvSpPr/><p:nvPr/></p:nvSpPr>' +
    '<p:spPr><a:xfrm><a:off x="1000000" y="2000000"/><a:ext cx="900000" cy="600000"/></a:xfrm>' +
    '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr></p:sp>' +
    '<p:sp><p:nvSpPr><p:cNvPr id="8" name="zone-label-z"></p:cNvPr><p:cNvSpPr/><p:nvPr/></p:nvSpPr>' +
    '<p:spPr><a:xfrm><a:off x="1000000" y="1800000"/><a:ext cx="900000" cy="150000"/></a:xfrm>' +
    '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr>' +
    '<p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr sz="900"/><a:t>Landing zone</a:t></a:r></a:p></p:txBody></p:sp>' +
    '</p:spTree>';

  const out = nativizeSlideXml(slide);
  assert.ok(out.includes('name="zone-label-z"'), 'the title keeps its own box');
  assert.ok(!/name="zone-z"[\s\S]*?<p:txBody>[\s\S]*?<\/p:sp>/.test(out.split('name="zone-label-z"')[0]), 'and is not folded into the zone');
});

test('the converted package still opens and its tags balance', async () => {
  const { nodes, edges } = chain(24);
  const pptx = await buildDiagramSlidePptx(PIXEL_PNG, {
    diagramName: 'Contoso Platform',
    author: 'Tester',
    date: '2026-08-10',
    isDarkMode: false,
    diagram: { nodes, edges },
  });
  const zip = await JSZip.loadAsync((await pptx.write({ outputType: 'nodebuffer' })) as Buffer);
  const before = Object.keys(zip.files).length;
  for (const name of Object.keys(zip.files).filter((n) => /^ppt\/slides\/slide\d+\.xml$/.test(n))) {
    const out = nativizeSlideXml(await zip.file(name)!.async('string'));
    for (const tag of ['p:sp', 'p:cxnSp', 'p:grpSp', 'p:pic', 'p:txBody']) {
      assert.equal(
        (out.match(new RegExp(`<${tag}>`, 'g')) ?? []).length,
        (out.match(new RegExp(`</${tag}>`, 'g')) ?? []).length,
        `${name}: <${tag}> tags balance`,
      );
    }
    zip.file(name, out);
  }
  const reopened = await JSZip.loadAsync((await zip.generateAsync({ type: 'nodebuffer' })) as Buffer);
  assert.equal(Object.keys(reopened.files).length, before, 'no part is lost');
  assert.ok(await reopened.file('ppt/presentation.xml')!.async('string'), 'the presentation part reads back');
});
