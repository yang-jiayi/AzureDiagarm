import test from 'node:test';
import assert from 'node:assert/strict';
import JSZip from 'jszip';
import type { Edge, Node } from 'reactflow';
import { buildDiagramSlidePptx } from '../src/services/pptxExporter.ts';
import { nativizeSlideXml } from '../src/services/pptxNativeShapes.ts';

const PIXEL_PNG =
  'data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg==';

function service(id: string, label: string, x: number, y: number, data: Record<string, unknown> = {}): Node {
  return {
    id,
    type: 'azureNode',
    position: { x, y },
    width: 150,
    height: 75,
    data: { label, serviceName: label, ...data },
  } as Node;
}

async function slideXml(nodes: Node[], edges: Edge[]): Promise<string> {
  const pptx = await buildDiagramSlidePptx(PIXEL_PNG, {
    diagramName: 'Semantics',
    author: 'Tester',
    date: '2026-08-10',
    isDarkMode: false,
    diagram: { nodes, edges },
  });
  const buffer = (await pptx.write({ outputType: 'nodebuffer' })) as Buffer;
  const zip = await JSZip.loadAsync(buffer);
  return zip.file('ppt/slides/slide1.xml')!.async('string');
}

test('PPTX colour-codes a security connector and renders a connection legend (fix 4)', async () => {
  const nodes = [
    service('a', 'Firewall', 0, 0),
    service('b', 'Gateway', 600, 0),
  ];
  const edges = [
    { id: 'sec', source: 'a', target: 'b', data: { connectionType: 'security' } },
  ] as unknown as Edge[];

  const xml = (await slideXml(nodes, edges)).toLowerCase();
  // Security connectors are red (#dc2626) — not the old flat slate grey.
  assert.ok(xml.includes('dc2626'), 'security connector uses the canonical red');
  // Scoped to the connector's own line. `64748b` is also the footer's text
  // colour — it was lifted to slate-500 so the credit line clears WCAG AA —
  // so a whole-document search now proves nothing about the connector.
  const connector = /<p:(?:sp|cxnsp)>(?:(?!<\/p:(?:sp|cxnsp)>)[\s\S])*name="connector-sec"[\s\S]*?<\/p:(?:sp|cxnsp)>/.exec(xml);
  assert.ok(connector, 'the security connector shape is emitted');
  assert.ok(!connector![0].includes('64748b'), 'no leftover flat-grey connector colour');
  // The legend shape and its label are present.
  assert.ok(xml.includes('connection-legend'), 'a connection legend shape is emitted');
  assert.ok(xml.includes('security'), 'the legend names the security connection type');
});

test('PPTX renders a SKU · region · cost sub-line for a service (fix 10)', async () => {
  const nodes = [
    service('svc', 'API', 0, 0, {
      sku: 'P1v3',
      region: 'eastus',
      pricing: { estimatedCost: 50, quantity: 1 },
    }),
  ];
  const xml = await slideXml(nodes, []);
  assert.ok(xml.includes('service-meta-svc'), 'a metadata sub-line shape is emitted');
  assert.ok(xml.includes('P1v3'), 'the SKU is shown');
  assert.ok(xml.includes('eastus'), 'the region is shown');
  assert.ok(xml.includes('$50'), 'the monthly cost is shown');
});

test('PPTX keeps a telemetry connector purple and dash-dotted (fix 4)', async () => {
  const nodes = [
    service('a', 'App', 0, 0),
    service('b', 'Monitor', 600, 0),
  ];
  const edges = [
    { id: 'tel', source: 'a', target: 'b', data: { connectionType: 'telemetry' } },
  ] as unknown as Edge[];

  const xml = (await slideXml(nodes, edges)).toLowerCase();
  assert.ok(xml.includes('7c3aed'), 'telemetry connector uses the canonical purple');
  assert.ok(xml.includes('dashdot') || xml.includes('dash'), 'telemetry connector is dashed');
});

/**
 * The canvas draws up to two tag chips on a tile, then a `+N` counter for the
 * rest. The deck used to draw none of them, which silently dropped the only
 * thing on the tile saying `prod` rather than `dev` — and the tile was already
 * measured tall enough to hold them, so the room was being paid for and thrown
 * away. `tagged()` uses the height React Flow reports for a tile with a chip
 * strip; `service()` uses the plain 75 an untagged tile measures.
 */
function tagged(id: string, label: string, x: number, y: number, tags: string[]): Node {
  return { ...service(id, label, x, y, { tags }), height: 94 } as Node;
}

/** Every chip drawn on `id`, in the order the fitter placed them. */
function chipsFor(xml: string, id: string): string[] {
  const shapes = [...xml.matchAll(/<p:sp>[\s\S]*?<\/p:sp>/g)].map((m) => m[0]);
  return shapes
    .filter((s) => new RegExp(`name="tagtext-${id}-\\d+"`).test(s))
    .map((s) => [...s.matchAll(/<a:t>([^<]*)<\/a:t>/g)].map((m) => m[1]).join(''));
}

test('PPTX draws the same two tag chips and +N counter the canvas draws', async () => {
  const xml = await slideXml([
    tagged('one', 'Storage', 0, 0, ['prod']),
    tagged('two', 'Compute', 400, 0, ['prod', 'pci']),
    tagged('five', 'Database', 800, 0, ['prod', 'pci', 'hipaa', 'gdpr', 'soc2']),
  ], []);

  assert.deepEqual(chipsFor(xml, 'one'), ['prod'], 'a single tag is drawn as one chip');
  assert.deepEqual(chipsFor(xml, 'two'), ['prod', 'pci'], 'two tags are both drawn');
  // Five tags: two named plus a counter for the remaining three, exactly as
  // `AzureNode` renders it. A deck has no tooltip, so the counter is the only
  // honest way to say "there is more here".
  assert.deepEqual(chipsFor(xml, 'five'), ['prod', 'pci', '+3'], 'the rest collapse into +N');
});

test('PPTX reserves no tag room on a tile that has no tags', async () => {
  // The band is conditional. If it were unconditional every untagged tile in
  // every existing deck would lose a strip of caption height to nothing.
  const xml = await slideXml([service('bare', 'Storage Account', 0, 0)], []);
  assert.equal(chipsFor(xml, 'bare').length, 0, 'no chips are emitted');
  assert.ok(!/name="tag-bare-/.test(xml), 'no chip background is emitted either');
});

test('PPTX never cuts a tag chip in half', async () => {
  // Service names are ellipsized because the index slide spells them out
  // again. A tag has no such second home, so a cut chip would be a claim the
  // deck cannot back up: the chips degrade whole, dropping to the counter.
  const xml = await slideXml([
    tagged('long', 'API', 0, 0, ['production-payment-card-industry-scope', 'x']),
  ], []);
  for (const chip of chipsFor(xml, 'long')) {
    assert.ok(!chip.includes('\u2026') && !chip.endsWith('...'), `chip "${chip}" was truncated`);
  }
});

test('tag chips travel with the tile when it is dragged', async () => {
  // `nativizeSlideXml` folds a tile's parts into one group. Anything drawn on
  // a tile that is not in that filter is left behind on the slide when the
  // reader moves the tile — which is how the category stripe once stranded.
  // Grouping happens in `nativizePackage`, after pptxgenjs has written the
  // deck, so this has to read the repaired bytes rather than the raw output.
  const xml = nativizeSlideXml(await slideXml([
    tagged('a', 'App', 0, 0, ['prod']),
    tagged('a-1', 'App replica', 400, 0, ['dr']),
  ], []));
  const grouped = [...xml.matchAll(/<p:grpSp>[\s\S]*?<\/p:grpSp>/g)].map((m) => m[0]);
  const inGroups = grouped.join('');
  for (const id of ['a', 'a-1']) {
    for (const chip of [`tag-${id}-0`, `tagtext-${id}-0`]) {
      assert.ok(inGroups.includes(`name="${chip}"`), `${chip} is inside a group`);
    }
  }
  // `a-1`'s chips must land in `a-1`'s group, not get swept into `a`'s by a
  // bare prefix match — node ids are allowed to contain the separator.
  const groupA = grouped.find((g) => /name="service-a"/.test(g)) ?? '';
  assert.ok(groupA, "the tile group for `a` exists");
  assert.ok(!groupA.includes('name="tag-a-1-'), "a-1's chips did not leak into a's group");
});
test('a tile too narrow for even the +N counter reserves no strip for it', async () => {
  // The band was gated on the tile's HEIGHT while the chips fit on its WIDTH,
  // so a tile narrow enough that not even `+3` fitted still gave up a strip
  // along its bottom — and because the icon-squeeze's room is capped by width,
  // dropping the tags could never win that width back. The strip stayed empty
  // and the name lost a line to it. A tagged tile that draws nothing must lay
  // out exactly as the same tile with no tags at all.
  //
  // 24px against a far-away node scales the tile to 0.250in, which is the
  // widest tile that cannot fit the counter; the measurements come from a
  // width sweep rather than a guess.
  const narrow = (tags?: string[]): Node[] => [
    {
      id: 'n',
      type: 'azureNode',
      position: { x: 0, y: 0 },
      width: 24,
      height: 94,
      data: {
        label: 'Application Gateway Frontend',
        serviceName: 'Application Gateway Frontend',
        ...(tags ? { tags } : {}),
      },
    },
    {
      id: 'far',
      type: 'azureNode',
      position: { x: 9000, y: 4000 },
      width: 150,
      height: 75,
      data: { label: 'Edge', serviceName: 'Edge' },
    },
  ] as Node[];

  const labelBox = (xml: string): string => {
    const shape = /<p:sp>(?:(?!<\/p:sp>)[\s\S])*name="service-label-n"[\s\S]*?<\/p:sp>/.exec(xml);
    assert.ok(shape, 'the tile draws its name');
    const ext = /<a:ext cx="\d+" cy="(\d+)"/.exec(shape[0]);
    assert.ok(ext, 'the name box has a height');
    return ext[1];
  };

  const tagged = await slideXml(narrow(['pci', 'hipaa', 'soc2']), []);
  const plain = await slideXml(narrow(), []);

  assert.equal(
    (tagged.match(/name="tagtext-n-\d+"/g) ?? []).length,
    0,
    'the tile is too narrow to draw any chip',
  );
  assert.equal(
    labelBox(tagged),
    labelBox(plain),
    'a tile that draws no chip gives up no room to them',
  );
});
test('tag chips carry the canvas brand colours, not an invented neutral', async () => {
  // `.node-tags span` is `--azd-color-brand-subtle` on `--azd-color-brand-border`
  // with `--azd-color-text-secondary` ink. The export drew grey, which is the
  // same class of defect as the zone panels: a colour the canvas owns being
  // re-decided in the file. The border is the visible tell — the canvas chip is
  // ringed in Azure blue.
  for (const [dark, fill, line] of [
    [false, 'EFF6FF', 'BFDBFE'],
    [true, '22384A', '42657E'],
  ] as const) {
    const pptx = await buildDiagramSlidePptx(PIXEL_PNG, {
      diagramName: 'Chips',
      author: 'Tester',
      date: '2026-08-10',
      isDarkMode: dark,
      diagram: { nodes: [tagged('t', 'Storage', 0, 0, ['prod', 'pci'])], edges: [] },
    });
    const buffer = (await pptx.write({ outputType: 'nodebuffer' })) as Buffer;
    const zip = await JSZip.loadAsync(buffer);
    const xml = await zip.file('ppt/slides/slide1.xml')!.async('string');
    const chip = /<p:sp>(?:(?!<\/p:sp>)[\s\S])*name="tag-t-0"[\s\S]*?<\/p:sp>/.exec(xml);
    assert.ok(chip, `${dark ? 'dark' : 'light'}: the chip shape is emitted`);
    assert.ok(chip[0].includes(fill), `${dark ? 'dark' : 'light'}: chip fill is ${fill}`);
    assert.ok(chip[0].includes(line), `${dark ? 'dark' : 'light'}: chip border is ${line}`);
  }
});
test('a slide that draws no connector carries no colour key', async () => {
  // A tiled deck used to key every connection type in the *diagram* on every
  // slide, so a grid whose only two security hops sit in one corner put a
  // "Security" swatch on nine consecutive slides that drew no line at all --
  // a key explaining something the reader cannot see, in a strip of space
  // taken from the drawing to hold it.
  const nodes: Node[] = [];
  for (let r = 0; r < 9; r++) {
    for (let c = 0; c < 12; c++) nodes.push(service(`n${r}-${c}`, `Service ${r}-${c}`, c * 220, r * 130));
  }
  const edges = [
    { id: 'e1', source: 'n0-0', target: 'n0-1', data: { connectionType: 'security' } },
    { id: 'e2', source: 'n0-1', target: 'n1-1', data: { connectionType: 'security' } },
  ] as unknown as Edge[];

  const pptx = await buildDiagramSlidePptx(PIXEL_PNG, {
    diagramName: 'Legend tiling', author: 'Tester', date: '2026-08-10',
    isDarkMode: false, diagram: { nodes, edges },
  });
  const zip = await JSZip.loadAsync((await pptx.write({ outputType: 'nodebuffer' })) as Buffer);
  const names = Object.keys(zip.files)
    .filter((n) => /^ppt\/slides\/slide\d+\.xml$/.test(n))
    .sort((a, b) => +a.replace(/\D/g, '') - +b.replace(/\D/g, ''));
  assert.ok(names.length > 3, 'the grid should have tiled across several slides');

  let keyed = 0;
  let withConnectors = 0;
  for (const name of names) {
    const xml = await zip.file(name)!.async('string');
    const hasKey = /name="connection-legend"/.test(xml);
    const drawsConnector = /name="connector-/.test(xml);
    if (hasKey) keyed++;
    if (drawsConnector) withConnectors++;
    assert.equal(
      hasKey, drawsConnector,
      `${name}: colour key ${hasKey ? 'present' : 'absent'} but the slide `
      + `${drawsConnector ? 'draws' : 'draws no'} connector`,
    );
  }
  assert.ok(withConnectors >= 1, 'the hops must be drawn somewhere');
  assert.ok(keyed < names.length, 'not every slide should carry the key');
});