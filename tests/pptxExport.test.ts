import assert from 'node:assert/strict';
import test from 'node:test';
import { posix } from 'node:path';
import { deflateSync } from 'node:zlib';
import JSZip from 'jszip';
import type { Edge, Node } from 'reactflow';
import {
  buildArchitectureDeckBlob, buildArchitectureDeckPptx,
  buildDiagramPptxBlob, buildDiagramSlidePptx,
  estimateTextWidthIn, wrappedLineCount, type ArchitectureDeckOptions,
} from '../src/services/pptxExporter';
import { nativizePackage } from '../src/services/pptxNativeShapes';
import { embedVectorIcons } from '../src/services/pptxVectorIcons';
import { GEOMETRY_EA_FONT, GEOMETRY_LATIN_FONT, singleLineName } from '../src/services/diagramExportGeometry';
import { nodesForExport } from '../src/utils/nodesForExport';
import type { ExportIcon, ExportIcons } from '../src/services/diagramExportIcons';

const options = { diagramName: 'Architecture & 日本語', author: 'Test author', date: '2026-09-05', isDarkMode: false };
const EMU = 914400;
const service = (id: string, x: number, y: number, data: Record<string, unknown> = {}): Node => ({
  id, type: 'azureNode', position: { x, y }, width: 300, height: 140, data: { label: id, ...data },
});
const nodes: Node[] = [
  { id: 'zone', type: 'groupNode', position: { x: -800, y: -450 }, style: { width: 1100, height: 650 }, data: { label: 'Region & Network' } },
  { ...service('a', 40, 100, { label: 'App Service', tags: ['prod', 'pci', 'regional'] }), parentNode: 'zone' },
  { ...service('b', 700, 100, { label: 'SQL Database' }), parentNode: 'zone' },
  { ...service('c', 700, 400, { label: 'Monitor' }), parentNode: 'zone' },
];
const edges: Edge[] = [
  { id: 'flow', source: 'a', target: 'b', label: 'HTTPS', data: { connectionType: 'security' } },
  { id: 'elbow', source: 'a', target: 'c', label: 'Telemetry', data: { connectionType: 'telemetry' } },
];

function attributes(tag: string): Record<string, string> {
  return Object.fromEntries([...tag.matchAll(/([\w:]+)="([^"]*)"/g)].map(match => [match[1], match[2]]));
}
function unescapeXml(value: string): string {
  return value.replace(/&lt;/g, '<').replace(/&gt;/g, '>').replace(/&quot;/g, '"')
    .replace(/&apos;/g, "'").replace(/&amp;/g, '&');
}
function allText(xml: string): string {
  return [...xml.matchAll(/<a:t\b[^>]*>([\s\S]*?)<\/a:t>/g)].map(match => unescapeXml(match[1])).join('');
}
function objects(xml: string, type = 'sp') {
  return [...xml.matchAll(new RegExp(`<p:${type}>[\\s\\S]*?<\\/p:${type}>`, 'g'))].map(match => {
    const attrs = attributes(match[0].match(/<p:cNvPr\b[^>]*>/)![0]);
    return { xml: match[0], ...attrs, name: attrs.name, id: attrs.id };
  });
}
async function slideXml(zip: JSZip): Promise<string[]> {
  return Promise.all(Object.keys(zip.files).filter(path => /^ppt\/slides\/slide\d+\.xml$/.test(path))
    .sort((a, b) => Number(a.match(/slide(\d+)/)![1]) - Number(b.match(/slide(\d+)/)![1]))
    .map(path => zip.file(path)!.async('string')));
}
async function packageFrom(blob: Blob) {
  const zip = await JSZip.loadAsync(await blob.arrayBuffer());
  return { zip, slides: await slideXml(zip) };
}
async function expectedPackage(pptx: Awaited<ReturnType<typeof buildDiagramSlidePptx>>, vectors: Map<string, string>) {
  const zip = await nativizePackage(
    await JSZip.loadAsync(await pptx.write({ outputType: 'arraybuffer' }) as ArrayBuffer),
    { latin: GEOMETRY_LATIN_FONT, ea: GEOMETRY_EA_FONT },
  );
  await embedVectorIcons(zip, vectors);
  return slideXml(zip);
}

function crc32(bytes: Uint8Array): number {
  let crc = 0xffffffff;
  for (const value of bytes) {
    crc ^= value;
    for (let bit = 0; bit < 8; bit++) crc = (crc >>> 1) ^ ((crc & 1) ? 0xedb88320 : 0);
  }
  return (crc ^ 0xffffffff) >>> 0;
}
function png(width: number, height: number): Buffer {
  const chunk = (name: string, bytes: Buffer) => {
    const length = Buffer.alloc(4), checksum = Buffer.alloc(4);
    length.writeUInt32BE(bytes.length);
    const content = Buffer.concat([Buffer.from(name), bytes]);
    checksum.writeUInt32BE(crc32(content));
    return Buffer.concat([length, content, checksum]);
  };
  const header = Buffer.alloc(13);
  header.writeUInt32BE(width, 0);
  header.writeUInt32BE(height, 4);
  header[8] = 8;
  header[9] = 6;
  return Buffer.concat([Buffer.from([137, 80, 78, 71, 13, 10, 26, 10]), chunk('IHDR', header),
    chunk('IDAT', deflateSync(Buffer.alloc((width * 4 + 1) * height))), chunk('IEND', Buffer.alloc(0))]);
}
const bytes = png(256, 256);
const icon: ExportIcon = {
  bytes, dataUrl: `data:image/png;base64,${bytes.toString('base64')}`, sizePx: 256,
  svg: '<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 64 32"><rect width="64" height="32" fill="#0078d4"/></svg>',
};

test('graph Blob export uses the mature geometry, native grouping, descriptions and connector repair', async () => {
  const before = structuredClone({ nodes, edges });
  const result = await packageFrom(await buildDiagramPptxBlob({ nodes, edges }, options, new Map()));
  const vectors = new Map<string, string>();
  const direct = await buildDiagramSlidePptx('', { ...options, diagram: { nodes, edges }, presetIcons: new Map() }, vectors);
  assert.deepEqual(result.slides, await expectedPackage(direct, vectors), 'no parallel legacy renderer may diverge from the audited export');
  const xml = result.slides[0];
  assert.ok(objects(xml, 'grpSp').some(group => group.name === 'node-a'));
  const group = objects(xml, 'grpSp').find(group => group.name === 'node-a')!;
  assert.match(group.xml, /name="tagtext-a-0"/, 'tag chips still move with their service');
  assert.match(group.xml, /descr="Service:/, 'native accessibility descriptions survive');
  assert.match(xml, /name="connection-legend/);
  const straight = objects(xml, 'cxnSp').find(shape => shape.name === 'connector-flow');
  assert.ok(straight);
  assert.match(straight.xml, /<a:stCxn/);
  assert.match(straight.xml, /<a:endCxn/);
  assert.match(straight.xml, /prst="straightConnector1"/);
  assert.ok(objects(xml).some(shape => shape.name === 'connector-elbow' && shape.xml.includes('<a:custGeom>')));
  for (const connector of objects(xml, 'cxnSp')) assert.doesNotMatch(connector.xml, /<a:custGeom>/);
  assert.equal(objects(xml, 'pic').length, 0, 'native input must not introduce a screenshot');
  const notes = await Promise.all(Object.keys(result.zip.files).filter(path => /^ppt\/notesSlides\/notesSlide\d+\.xml$/.test(path))
    .map(path => result.zip.file(path)!.async('string')));
  assert.match(notes.join(''), /Services on this slide/);
  assert.deepEqual({ nodes, edges }, before);
});

test('single and customer Blob exports retain SVG originals, PNG fallback chains and exact script fonts', async () => {
  const path = '/test-icon.svg';
  const input = [
    service('a&1', 0, 0, { label: '日本語 API', iconPath: path }),
    service('b', 500, 0, { label: 'Database', iconPath: path }),
  ];
  const icons: ExportIcons = new Map([[path, icon]]);
  for (const blob of [
    await buildDiagramPptxBlob({ nodes: input, edges: [] }, options, icons),
    await buildArchitectureDeckBlob({ nodes: input, edges: [] }, { ...options, services: [] }, icons),
  ]) {
    const result = await packageFrom(blob);
    let pictures = 0;
    for (const [index, xml] of result.slides.entries()) {
      assert.doesNotMatch(xml, /<a:ea typeface="Arial"/);
      assert.match(xml, /<a:latin typeface="Arial"/);
      assert.match(xml, /<a:ea typeface="Yu Gothic UI"/);
      const relXml = await result.zip.file(`ppt/slides/_rels/slide${index + 1}.xml.rels`)!.async('string');
      const rels = new Map([...relXml.matchAll(/<Relationship\b[^>]*>/g)].map(match => {
        const attrs = attributes(match[0]);
        return [attrs.Id, attrs];
      }));
      const ids = [...relXml.matchAll(/\bId="([^"]*)"/g)].map(match => match[1]);
      assert.equal(ids.length, new Set(ids).size);
      for (const pic of objects(xml, 'pic')) {
        pictures++;
        const pngId = pic.xml.match(/<a:blip r:embed="([^"]+)"/)![1];
        const svgId = pic.xml.match(/<asvg:svgBlip[^>]*r:embed="([^"]+)"/)![1];
        for (const [id, extension] of [[pngId, 'png'], [svgId, 'svg']]) {
          const rel = rels.get(id);
          assert.ok(rel && rel.Type.endsWith('/image'));
          const part = posix.normalize(posix.join('ppt/slides', rel.Target));
          assert.ok(part.endsWith(`.${extension}`));
          const entry = result.zip.file(part);
          assert.ok(entry, `missing internal ${extension} payload`);
          if (extension === 'png') assert.deepEqual(await entry.async('uint8array'), new Uint8Array(bytes));
          else assert.equal(await entry.async('string'), icon.svg);
        }
      }
      if (objects(xml, 'pic').length) {
        assert.equal([...relXml.matchAll(/Target="[^"]*\.svg"/g)].length, 1, 'identical SVGs share one media part per slide');
      }
    }
    assert.equal(pictures, 2);
  }
});

test('connector chips emit the same vertical padding and absolute leading their fitter reserves', async () => {
  const input = [service('a', 0, 0), service('b', 500, 0)];
  for (const label of ['HTTPS / TLS 1.2', '日本語の接続ラベル']) {
    const result = await packageFrom(await buildDiagramPptxBlob({
      nodes: input, edges: [{ id: 'label', source: 'a', target: 'b', label }],
    }, options, new Map()));
    const chip = objects(result.slides[0]).find(shape => shape.name === 'connector-label-label')!;
    assert.ok(chip);
    const frame = attributes(chip.xml.match(/<a:bodyPr\b[^>]*>/)![0]);
    const ext = attributes(chip.xml.match(/<a:ext\b[^>]*>/)![0]);
    const pt = Number(chip.xml.match(/<a:rPr\b[^>]*sz="(\d+)"/)![1]) / 100;
    assert.equal(Number(frame.lIns) / EMU, 0.06);
    assert.equal(Number(frame.rIns) / EMU, 0.06);
    assert.equal(Number(frame.tIns) / EMU, 0.03);
    assert.equal(Number(frame.bIns) / EMU, 0.03);
    const leading = Number(chip.xml.match(/<a:lnSpc><a:spcPts val="(\d+)"/)![1]) / 100;
    assert.ok(Math.abs(leading - pt * 1.3) <= 0.01);
    const width = (Number(ext.cx) - Number(frame.lIns) - Number(frame.rIns)) / EMU;
    const height = (Number(ext.cy) - Number(frame.tIns) - Number(frame.bIns)) / EMU;
    assert.ok(wrappedLineCount(label, width, pt) * leading / 72 <= height + 0.0001);
  }
});

test('public graph and legacy builder deck paths share the mature pagination and inventory', async () => {
  const services = Array.from({ length: 53 }, (_, index) => ({
    name: `Service-${String(index).padStart(3, '0')} 日本語サービス`,
    category: 'Application services', group: 'Production',
  }));
  const deck: ArchitectureDeckOptions = { ...options, services };
  const result = await packageFrom(await buildArchitectureDeckBlob({ nodes, edges }, deck, new Map()));
  const vectors = new Map<string, string>();
  const direct = await buildArchitectureDeckPptx('', { ...deck, diagram: { nodes, edges }, presetIcons: new Map() }, vectors);
  const tableIds = (xml: string) => xml.replace(/<p:graphicFrame>[\s\S]*?<\/p:graphicFrame>/g,
    frame => frame.replace(/(<p:cNvPr\b[^>]*\bid=")\d+(")/, '$1table-id$2'));
  assert.deepEqual(result.slides.map(tableIds), (await expectedPackage(direct, vectors)).map(tableIds),
    'table ID repair changes no drawing geometry, typography, content or attachment');
  for (const xml of result.slides) {
    const ids = [...xml.matchAll(/<p:cNvPr\b[^>]*\bid="(\d+)"/g)].map(match => match[1]);
    assert.equal(ids.length, new Set(ids).size, 'inventory tables must not reuse header shape IDs');
  }
  const inventory = result.slides.filter(xml => allText(xml).includes('Services  ·  53 components'));
  assert.ok(inventory.length > 1);
  for (const entry of services) assert.ok(inventory.some(xml => allText(xml).includes(entry.name)));
});

test('unavailable, genuinely free and capacity-covered prices remain distinct in both PPTX paths', async () => {
  const input = [
    service('unknown', 0, 0, { pricing: { estimatedCost: null } }),
    service('free', 500, 0, { pricing: { estimatedCost: 0 } }),
    service('lake', 1000, 0, { serviceName: 'Fabric Lakehouse', pricing: { estimatedCost: 0 } }),
  ];
  const exportOptions = { ...options, priceUnavailableLabel: '価格未確認', capacityLabel: 'Shared capacity' };
  for (const isDarkMode of [false, true]) {
    for (const blob of [
      await buildDiagramPptxBlob({ nodes: input, edges: [] }, { ...exportOptions, isDarkMode }, new Map()),
      await buildArchitectureDeckBlob({ nodes: input, edges: [] }, { ...exportOptions, isDarkMode, services: [] }, new Map()),
    ]) {
      const { slides } = await packageFrom(blob);
      const shapes = slides.flatMap(xml => objects(xml));
      assert.equal(allText(shapes.find(shape => shape.name === 'service-meta-unknown')!.xml), '価格未確認');
      assert.equal(allText(shapes.find(shape => shape.name === 'service-meta-free')!.xml), 'Free');
      assert.equal(allText(shapes.find(shape => shape.name === 'service-meta-lake')!.xml), 'Shared capacity');
    }
  }
  const hidden = await packageFrom(await buildDiagramPptxBlob({ nodes: nodesForExport(input, false), edges: [] }, exportOptions, new Map()));
  assert.doesNotMatch(hidden.slides.join(''), /service-meta-(?:unknown|free|lake)|価格未確認|Shared capacity/);
});

test('partial estimates retain every excluded service on legible pages and suppress regional rankings', async () => {
  const unpricedServices = [
    ...Array.from({ length: 70 }, (_, i) => `Unpriced service ${i + 1} with an unconfigured consumption assumption`),
    `${'価格が未確認の長いサービス名'.repeat(230)} END-OF-LAST-SERVICE`,
    'Cafe\u0301 & Partners\u000bProduction',
  ];
  const deck: ArchitectureDeckOptions = {
    ...options, services: [],
    cost: {
      totalMonthly: 0, annual: 0, currency: 'USD', byCategory: [], topServices: [], unpricedServices,
      regionComparisonIncomplete: true, unavailableRegions: ['Japan West: selected SKU unavailable'],
      regions: [
        { name: 'Japan East', monthly: 0, annual: 0, isCheapest: true },
        { name: 'Japan West', monthly: 10, annual: 120 },
      ],
    },
  };
  const before = structuredClone(deck);
  const result = await packageFrom(await buildArchitectureDeckBlob({ nodes: [], edges: [] }, deck, new Map()));
  const text = result.slides.map(allText).join('\n');
  assert.match(text, /Partial cost estimate/);
  assert.match(text, /known-cost subtotal per month \(incomplete\)/);
  assert.match(text, /excluded, not free/);
  assert.match(text, /selected SKUs are unavailable/);
  assert.doesNotMatch(text, /Regional cost comparison|potential saving|Already on the cheapest/);
  const exclusionPages = result.slides.filter(xml => allText(xml).includes('Services excluded from the estimate'));
  assert.ok(exclusionPages.length > 3);
  const lines = exclusionPages.flatMap(xml => objects(xml)).filter(shape => shape.name.startsWith('unpriced-service-'));
  for (const [index, name] of unpricedServices.entries()) {
    const drawn = lines.filter(line => line.name.startsWith(`unpriced-service-${index}-`)).map(line => allText(line.xml)).join('');
    assert.equal(drawn.replace(/\s/g, ''), `${index + 1}.${singleLineName(name)}`.replace(/\s/g, ''));
  }
  for (const line of lines) {
    const off = attributes(line.xml.match(/<a:off\b[^>]*>/)![0]);
    const ext = attributes(line.xml.match(/<a:ext\b[^>]*>/)![0]);
    const pt = Number(line.xml.match(/<a:rPr\b[^>]*sz="(\d+)"/)![1]) / 100;
    assert.equal(pt, 12);
    assert.ok(Number(off.y) / EMU >= 1.12);
    assert.ok((Number(off.y) + Number(ext.cy)) / EMU <= 7.04 + 1e-5, `${line.name} overlaps the footer`);
    assert.ok(estimateTextWidthIn(allText(line.xml), pt) <= Number(ext.cx) / EMU);
    assert.match(line.xml, /<a:lnSpc><a:spcPts /, 'disclosure rows use absolute leading');
    assert.doesNotMatch(line.xml, /<a:lnSpc><a:spcPct /);
  }
  assert.deepEqual(deck, before);
});

test('complete estimates keep the existing regional comparison and price freshness disclosures', async () => {
  const result = await packageFrom(await buildArchitectureDeckBlob({ nodes: [], edges: [] }, {
    ...options, services: [],
    cost: {
      totalMonthly: 100, annual: 1200, currency: 'USD', byCategory: [], topServices: [],
      unpricedServices: [], pricesAsOf: '2026-09-01', oldestMeterAsOf: '2024-09-01',
      regions: [
        { name: 'Japan East', monthly: 100, annual: 1200, isCheapest: true },
        { name: 'Japan West', monthly: 110, annual: 1320, isCurrent: true },
      ],
    },
  }, new Map()));
  const text = result.slides.map(allText).join('\n');
  assert.match(text, /Regional cost comparison/);
  assert.match(text, /potential saving/);
  assert.match(text, /prices as of 2026-09-01/);
  assert.match(text, /unchanged since 2024-09-01/);
  assert.doesNotMatch(text, /Partial cost estimate|Services excluded from the estimate/);
});

test('empty graph and legacy image inputs remain valid with no browser download side effects', async () => {
  for (const isDarkMode of [false, true]) {
    const empty = await packageFrom(await buildDiagramPptxBlob({ nodes: [], edges: [] }, { ...options, isDarkMode }));
    assert.equal(empty.slides.length, 1);
    assert.equal(objects(empty.slides[0], 'pic').length, 0);
    const legacy = await packageFrom(await buildDiagramPptxBlob(icon.dataUrl, { ...options, isDarkMode }));
    assert.equal(legacy.slides.length, 1);
    assert.equal(objects(legacy.slides[0], 'pic').length, 1);
  }
});
