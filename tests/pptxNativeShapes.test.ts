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

test('every arrow survives the conversion, and none of them corrupts the deck', async () => {
  const { nodes, edges } = chain(8);
  const slides = await slidesOf(nodes, edges);
  let drawn = 0;
  let connectors = 0;
  let kept = 0;
  let glued = 0;
  const named = (xml: string, tag: 'p:sp' | 'p:cxnSp') => [
    ...xml.matchAll(new RegExp(`<${tag}>[\\s\\S]*?</${tag}>`, 'g')),
  ].filter((match) => /name="connector-e\d+"/.test(match[0])).map((match) => match[0]);

  for (const xml of slides) {
    const out = nativizeSlideXml(xml);
    drawn += named(xml, 'p:sp').length;
    connectors += named(out, 'p:cxnSp').length;
    kept += named(out, 'p:sp').length;
    glued += (out.match(/<a:stCxn /g) ?? []).length;

    // PowerPoint refuses the whole package, not the shape, when a <p:cxnSp>
    // carries its own geometry. So a hop that bends stays a plain shape with
    // the route it was drawn with, rather than becoming a connector that
    // PowerPoint would re-route on open.
    for (const connector of named(out, 'p:cxnSp')) {
      assert.ok(
        !connector.includes('<a:custGeom>'),
        'a connector carrying custom geometry makes PowerPoint reject the file',
      );
      assert.match(
        connector,
        /<a:prstGeom prst="(straightConnector1|bentConnector3)"/,
        'a connector needs a preset PowerPoint can re-route',
      );
    }
    for (const arrow of named(out, 'p:sp')) {
      assert.ok(arrow.includes('<a:custGeom>'), 'an arrow left as a shape keeps its exact route');
    }

    // Glue that points at a shape which is not on the slide makes PowerPoint
    // drop the arrow on open, which is worse than not gluing at all.
    const ids = new Set([...out.matchAll(/<p:cNvPr id="(\d+)"/g)].map((m) => m[1]));
    for (const cxn of out.matchAll(/<a:(?:st|end)Cxn id="(\d+)" idx="(\d+)"\/>/g)) {
      assert.ok(ids.has(cxn[1]), `glued to shape id ${cxn[1]}, which is not on the slide`);
      assert.ok(+cxn[2] <= 3, `glued to connection site ${cxn[2]}, which a rectangle does not have`);
    }
  }
  assert.ok(drawn > 0, 'the fixture must draw arrows');
  assert.equal(connectors + kept, drawn, 'no arrow is lost in the conversion');
  assert.ok(connectors > 0, 'the hops that run straight become real connectors');
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

test('two tiles that touch do not get an arrow glued to itself', async () => {
  // Tiles 150px wide on a 150px pitch meet exactly, so the hop between them is
  // zero-length and both of its ends land on the same site of the same shape.
  // Gluing that tells PowerPoint the arrow starts and finishes in one place.
  const nodes = [service('a', 'A', 0, 0), service('b', 'B', 150, 0), service('c', 'C', 300, 0)];
  const edges = [
    { id: 'ab', source: 'a', target: 'b', label: 'calls' },
    { id: 'bc', source: 'b', target: 'c', label: 'calls' },
  ] as Edge[];
  for (const xml of await slidesOf(nodes, edges)) {
    const out = nativizeSlideXml(xml);
    for (const cxn of out.matchAll(/<p:cxnSp>[\s\S]*?<\/p:cxnSp>/g)) {
      const start = /<a:stCxn id="(\d+)" idx="(\d+)"\/>/.exec(cxn[0]);
      const end = /<a:endCxn id="(\d+)" idx="(\d+)"\/>/.exec(cxn[0]);
      if (!start || !end) continue;
      assert.ok(
        !(start[1] === end[1] && start[2] === end[2]),
        `connector glued to shape ${start[1]} site ${start[2]} at both ends`,
      );
    }
  }
});

test('a group holds everything it claims to hold, at the coordinates it was drawn at', async () => {
  // Nothing rasterizes under Node, so a generated deck has no icon to group
  // with. Synthesise the picture a browser would emit, which is the only
  // configuration a user ever exports, and check the frame that results.
  const { nodes, edges } = chain(6);
  const slides = await slidesOf(nodes, edges);
  let checked = 0;
  for (const xml of slides) {
    const tiles = [...xml.matchAll(/<p:sp>[\s\S]*?<\/p:sp>/g)]
      .map((m) => m[0])
      .filter((sp) => /name="service-[^"]*"/.test(sp) && !/name="service-(label|meta)-/.test(sp));
    let nextId = Math.max(...[...xml.matchAll(/<p:cNvPr id="(\d+)"/g)].map((m) => +m[1])) + 1;
    const pics = tiles.map((sp) => {
      const id = /name="service-([^"]*)"/.exec(sp)![1];
      const off = /<a:off x="(-?\d+)" y="(-?\d+)"\/>/.exec(sp)!;
      const ext = /<a:ext cx="(\d+)" cy="(\d+)"\/>/.exec(sp)!;
      const size = Math.round(Math.min(+ext[1] * 0.3, +ext[2] * 0.42));
      const pic = `<p:pic><p:nvPicPr><p:cNvPr id="${nextId}" name="icon-${id}"/>`
        + `<p:cNvPicPr><a:picLocks noChangeAspect="1"/></p:cNvPicPr><p:nvPr/></p:nvPicPr>`
        + `<p:blipFill><a:blip r:embed="rId2"/><a:stretch><a:fillRect/></a:stretch></p:blipFill>`
        + `<p:spPr><a:xfrm><a:off x="${+off[1] + Math.round((+ext[1] - size) / 2)}" y="${+off[2] + Math.round(+ext[2] * 0.06)}"/>`
        + `<a:ext cx="${size}" cy="${size}"/></a:xfrm>`
        + `<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr></p:pic>`;
      nextId += 1;
      return pic;
    });
    if (pics.length === 0) continue;
    const out = nativizeSlideXml(xml.replace('</p:spTree>', `${pics.join('')}</p:spTree>`));

    for (const group of out.matchAll(/<p:grpSp>[\s\S]*?<\/p:grpSp>/g)) {
      const frame = /<p:grpSpPr><a:xfrm><a:off x="(-?\d+)" y="(-?\d+)"\/><a:ext cx="(\d+)" cy="(\d+)"\/><a:chOff x="(-?\d+)" y="(-?\d+)"\/><a:chExt cx="(\d+)" cy="(\d+)"\/>/.exec(group[0]);
      assert.ok(frame, 'group has no readable frame');
      const [ox, oy, cx, cy, hx, hy, hcx, hcy] = frame!.slice(1).map(Number);
      // Child coordinates stay absolute only while these agree. Any drift
      // silently translates and rescales everything inside the group.
      assert.deepEqual([hx, hy, hcx, hcy], [ox, oy, cx, cy], 'group child frame differs from its own frame');

      for (const child of group[0].matchAll(/<p:(sp|pic)>[\s\S]*?<\/p:\1>/g)) {
        const off = /<a:off x="(-?\d+)" y="(-?\d+)"\/>/.exec(child[0]);
        const ext = /<a:ext cx="(\d+)" cy="(\d+)"\/>/.exec(child[0]);
        if (!off || !ext) continue;
        const slack = 9144; // 0.01in
        assert.ok(
          +off[1] >= ox - slack && +off[2] >= oy - slack
            && +off[1] + +ext[1] <= ox + cx + slack && +off[2] + +ext[2] <= oy + cy + slack,
          'group does not enclose a child, which clips it',
        );
        checked += 1;
      }
    }
  }
  assert.ok(checked > 0, 'no grouped children were checked, so this test proves nothing');
});
