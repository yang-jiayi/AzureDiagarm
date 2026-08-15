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