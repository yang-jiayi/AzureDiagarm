import test from 'node:test';
import assert from 'node:assert/strict';
import JSZip from 'jszip';
import type { Node } from 'reactflow';
import { buildDiagramSlidePptx } from '../src/services/pptxExporter.ts';
import { nativizePackage } from '../src/services/pptxNativeShapes.ts';
import { embedVectorIcons } from '../src/services/pptxVectorIcons.ts';

/**
 * A slide can only carry an icon as a raster, so the deck drew them at 128px —
 * roughly 213 dpi at the largest size a tile ever gives an icon. That is soft
 * as soon as the deck is projected, printed, or a service is pulled out and
 * scaled up. OOXML lets a picture carry the vector original alongside its
 * raster, and PowerPoint 2016+ draws that instead, so these tests pin the
 * attachment: the SVG has to reach the package, its relationship has to
 * resolve, and the raster has to survive for the readers that need it.
 */

const PIXEL_PNG =
  'data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg==';
const PIXEL_BYTES = Uint8Array.from(Buffer.from(PIXEL_PNG.split(',')[1], 'base64'));
const SVG_A = '<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 18 18"><circle cx="9" cy="9" r="8"/></svg>';
const SVG_B = '<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 18 18"><rect width="18" height="18"/></svg>';

function service(id: string, label: string, x: number, iconPath: string): Node {
  return {
    id,
    type: 'azureNode',
    position: { x, y: 0 },
    width: 150,
    height: 100,
    data: { label, serviceName: label, iconPath, category: 'compute' },
  } as Node;
}

/** The delivered package: what `downloadNativePptx` actually hands the user. */
async function deliver(nodes: Node[]): Promise<JSZip> {
  const vectorIcons = new Map<string, string>();
  const presetIcons = new Map([
    ['/i/a.svg', { bytes: PIXEL_BYTES, dataUrl: PIXEL_PNG, sizePx: 128, svg: SVG_A }],
    ['/i/b.svg', { bytes: PIXEL_BYTES, dataUrl: PIXEL_PNG, sizePx: 128, svg: SVG_B }],
  ]);
  const pptx = await buildDiagramSlidePptx(PIXEL_PNG, {
    diagramName: 'Contoso Platform',
    author: 'Tester',
    date: '2026-08-10',
    isDarkMode: false,
    diagram: { nodes, edges: [] },
    presetIcons,
  }, vectorIcons);
  const zip = await nativizePackage(await JSZip.loadAsync((await pptx.write({ outputType: 'nodebuffer' })) as Buffer));
  await embedVectorIcons(zip, vectorIcons);
  return zip as JSZip;
}

async function slide1(zip: JSZip): Promise<{ xml: string; rels: Map<string, string> }> {
  const xml = await zip.file('ppt/slides/slide1.xml')!.async('string');
  const relsXml = await zip.file('ppt/slides/_rels/slide1.xml.rels')!.async('string');
  const rels = new Map([...relsXml.matchAll(/Id="([^"]+)"[^>]*Target="([^"]+)"/g)].map((m) => [m[1], m[2]]));
  return { xml, rels };
}

test('every icon picture carries its vector original, and it resolves', async () => {
  const zip = await deliver([
    service('a', 'App Service', 0, '/i/a.svg'),
    service('b', 'SQL Database', 300, '/i/b.svg'),
  ]);
  const { xml, rels } = await slide1(zip);

  const pics = (xml.match(/<p:pic>/g) ?? []).length;
  assert.ok(pics >= 2, `expected the icons to be drawn, saw ${pics} pictures`);

  const svgRefs = [...xml.matchAll(/<asvg:svgBlip[^>]*r:embed="([^"]+)"/g)].map((m) => m[1]);
  assert.equal(svgRefs.length, pics, 'every icon picture should carry a vector original');

  // A relationship id that points at nothing, or at the wrong part, is exactly
  // how a deck opens with a repair prompt, so follow the whole chain.
  for (const id of svgRefs) {
    const target = rels.get(id);
    assert.ok(target, `${id} is referenced by a picture but declared by no relationship`);
    assert.match(target, /\.svg$/, `${id} should point at an SVG, points at ${target}`);
    assert.ok(zip.file(`ppt/${target.replace('../', '')}`), `${target} is declared but not in the package`);
  }
});

test('the raster survives, so PowerPoint 2013 and older still see an icon', async () => {
  const zip = await deliver([service('a', 'App Service', 0, '/i/a.svg')]);
  const { xml, rels } = await slide1(zip);

  const blips = [...xml.matchAll(/<a:blip r:embed="([^"]+)"><a:extLst>/g)].map((m) => m[1]);
  assert.ok(blips.length > 0, 'the picture should keep its own raster blip');
  for (const id of blips) {
    assert.match(rels.get(id) ?? '', /\.png$/, 'the primary blip should still be the PNG');
  }
});

test('one media part per distinct icon, however often it is drawn', async () => {
  const zip = await deliver([
    service('a', 'App Service', 0, '/i/a.svg'),
    service('b', 'SQL Database', 300, '/i/b.svg'),
    service('c', 'Second App', 600, '/i/a.svg'),
  ]);
  const { xml } = await slide1(zip);

  const svgParts = Object.keys(zip.files).filter((name) => name.startsWith('ppt/media/') && name.endsWith('.svg'));
  const svgRefs = (xml.match(/<asvg:svgBlip/g) ?? []).length;
  assert.ok(svgRefs >= 3, `all three icons should be vectorised, saw ${svgRefs}`);
  assert.equal(svgParts.length, 2, `two distinct icons should store two parts, stored ${svgParts.length}`);
});

test('relationship ids are unique across the whole slide', async () => {
  const zip = await deliver([
    service('a', 'App Service', 0, '/i/a.svg'),
    service('b', 'SQL Database', 300, '/i/b.svg'),
  ]);
  const relsXml = await zip.file('ppt/slides/_rels/slide1.xml.rels')!.async('string');
  const ids = [...relsXml.matchAll(/Id="([^"]+)"/g)].map((m) => m[1]);
  assert.equal(new Set(ids).size, ids.length, `duplicate relationship id in ${ids.join(', ')}`);
});

/**
 * A synthetic package, because real output happens to number its relationships
 * contiguously and so cannot show the difference between counting them and
 * taking the maximum. Numbering from the count is the classic way to mint an
 * id that is already taken, which silently repoints an existing picture.
 */
function fakeZip(slideXml: string, relsXml: string) {
  const files: Record<string, string> = {
    'ppt/slides/slide1.xml': slideXml,
    'ppt/slides/_rels/slide1.xml.rels': relsXml,
  };
  return {
    files,
    file(path: string, data?: string) {
      if (data !== undefined) { files[path] = data; return undefined; }
      return files[path] === undefined ? null : { async: async () => files[path] };
    },
  } as never;
}

const PIC = (name: string, embed: string, extra = ''): string =>
  `<p:pic><p:nvPicPr><p:cNvPr id="2" name="${name}" descr="x"/></p:nvPicPr>`
  + `<p:blipFill><a:blip r:embed="${embed}">${extra}</a:blip></p:blipFill></p:pic>`;

test('a new relationship is numbered from the highest id, not the count', async () => {
  const zip = fakeZip(
    `<p:sld>${PIC('icon-a', 'rId7')}</p:sld>`,
    '<Relationships><Relationship Id="rId7" Target="../media/image1.png"/></Relationships>',
  );
  await embedVectorIcons(zip, new Map([['icon-a', SVG_A]]));

  const rels = (zip as unknown as { files: Record<string, string> }).files['ppt/slides/_rels/slide1.xml.rels'];
  const ids = [...rels.matchAll(/Id="([^"]+)"/g)].map((m) => m[1]);
  assert.equal(new Set(ids).size, ids.length, `minted an id that was already taken: ${ids.join(', ')}`);
  assert.ok(ids.includes('rId8'), `expected rId8 after rId7, got ${ids.join(', ')}`);
});

test('a picture the diagram did not draw is left exactly as it was', async () => {
  const before = `<p:sld>${PIC('logo', 'rId1')}</p:sld>`;
  const zip = fakeZip(before, '<Relationships><Relationship Id="rId1" Target="../media/image1.png"/></Relationships>');
  await embedVectorIcons(zip, new Map([['icon-a', SVG_A]]));

  const files = (zip as unknown as { files: Record<string, string> }).files;
  assert.equal(files['ppt/slides/slide1.xml'], before, 'an unrelated picture must not be rewritten');
  assert.ok(!Object.keys(files).some((n) => n.endsWith('.svg')), 'no media part should be written for it');
});

test('a picture that already carries an extension list is not given a second one', async () => {
  const existing = '<a:extLst><a:ext uri="{OTHER}"/></a:extLst>';
  const before = `<p:sld>${PIC('icon-a', 'rId1', existing)}</p:sld>`;
  const zip = fakeZip(before, '<Relationships><Relationship Id="rId1" Target="../media/image1.png"/></Relationships>');
  await embedVectorIcons(zip, new Map([['icon-a', SVG_A]]));

  const files = (zip as unknown as { files: Record<string, string> }).files;
  assert.equal(files['ppt/slides/slide1.xml'], before, 'an existing extension list must be left alone');
});

test('a blip with children still gets its vector original', async () => {
  // A blip legitimately carries children — `<a:alphaModFix/>` for transparency,
  // a duotone recolour. Matching only empty blips would leave those pictures
  // raster and say nothing about it.
  const zip = fakeZip(
    `<p:sld>${PIC('icon-a', 'rId1', '<a:alphaModFix amt="80000"/>')}</p:sld>`,
    '<Relationships><Relationship Id="rId1" Target="../media/image1.png"/></Relationships>',
  );
  await embedVectorIcons(zip, new Map([['icon-a', SVG_A]]));

  const xml = (zip as unknown as { files: Record<string, string> }).files['ppt/slides/slide1.xml'];
  assert.match(xml, /<asvg:svgBlip[^>]*r:embed="rId2"/, 'the picture should have been vectorised');
  assert.match(xml, /<a:alphaModFix amt="80000"\/><a:extLst>/, 'the blip\'s own children must survive');
});

test('the icon says what it is, so a screen reader announces a service', async () => {
  const zip = await deliver([service('a', 'App Service', 0, '/i/a.svg')]);
  const { xml } = await slide1(zip);
  // pptxgenjs falls back to the image's own name, so without alt text every
  // icon in the deck introduced itself to a screen reader as `preencoded.png`.
  assert.equal((xml.match(/descr="preencoded\.png"/g) ?? []).length, 0, 'no icon should be described by its file name');
  assert.match(xml, /<p:cNvPr[^>]*name="icon-a"[^>]*descr="App Service icon"/, 'the icon should be described by its service');
});
