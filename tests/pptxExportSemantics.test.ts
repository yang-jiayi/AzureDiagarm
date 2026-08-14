import test from 'node:test';
import assert from 'node:assert/strict';
import JSZip from 'jszip';
import type { Edge, Node } from 'reactflow';
import { buildDiagramSlidePptx } from '../src/services/pptxExporter.ts';

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
