import test from 'node:test';
import assert from 'node:assert/strict';
import type { Edge, Node } from 'reactflow';
import { buildVsdxPackage } from '../src/services/visioVsdxExporter.ts';

function service(id: string, label: string, x: number, y: number, data: Record<string, unknown> = {}): Node {
  return {
    id,
    type: 'azureNode',
    position: { x, y },
    width: 150,
    height: 75,
    data: {
      label,
      serviceName: label,
      iconPath: '/Azure_Public_Service_Icons/Icons/compute/10021-icon-service-Function-Apps.svg',
      ...data,
    },
  } as Node;
}

async function pageXml(nodes: Node[], edges: Edge[]): Promise<string> {
  const pkg = await buildVsdxPackage(nodes, edges, 'Semantics');
  const page = pkg.parts.find((part) => part.path === 'visio/pages/page1.xml');
  assert.ok(page, 'page1.xml must be present');
  return page.data as string;
}

test('VSDX colour-codes a security connector and appends a connection legend (fix 4)', async () => {
  const nodes = [
    service('a', 'Firewall', 0, 0),
    service('b', 'Gateway', 600, 0),
  ];
  const edges = [
    { id: 'sec', source: 'a', target: 'b', data: { connectionType: 'security' } },
    { id: 'tel', source: 'a', target: 'b', data: { connectionType: 'telemetry' } },
  ] as unknown as Edge[];

  const xml = await pageXml(nodes, edges);
  // Canonical connection colours reach the Visio connectors.
  assert.ok(xml.includes('V="#dc2626"'), 'the security connector is red');
  assert.ok(xml.includes('V="#7c3aed"'), 'the telemetry connector is purple');
  // A colour key is emitted so the drawing agrees with the PNG legend.
  assert.ok(/NameU="Legend/.test(xml), 'a connection legend group is present');
  assert.ok(xml.includes('Security'), 'the legend names the security connection type');
  assert.ok(xml.includes('Telemetry'), 'the legend names the telemetry connection type');
});

test('VSDX surfaces service metadata both visibly and as shape data (fix 10)', async () => {
  const nodes = [
    service('svc', 'API', 0, 0, {
      sku: 'P1v3',
      region: 'eastus',
      pricing: { estimatedCost: 50, quantity: 1 },
    }),
  ];
  const xml = await pageXml(nodes, []);
  // Visible sub-line (second text run) so the info is not hidden in Shape Data.
  assert.ok(xml.includes('P1v3'), 'the SKU appears on the tile');
  assert.ok(xml.includes('eastus'), 'the region appears on the tile');
  // Still queryable as Visio Shape Data too.
  assert.ok(xml.includes('N="Sku"'), 'SKU is exposed as shape data');
  assert.ok(xml.includes('N="Region"'), 'region is exposed as shape data');
  assert.ok(xml.includes('N="MonthlyCost"'), 'monthly cost is exposed as shape data');
});

test('VSDX honours a zone custom colour (fix 6)', async () => {
  const nodes = [
    { id: 'zone', type: 'groupNode', position: { x: 0, y: 0 }, style: { width: 500, height: 320 }, data: { label: 'Prod', customColor: { border: '#dc2626' } } },
    service('svc', 'API', 60, 80, {}),
  ] as unknown as Node[];
  (nodes[1] as { parentNode?: string }).parentNode = 'zone';

  const xml = await pageXml(nodes, []);
  // The zone's own colour (not a palette-by-index guess) tints the band.
  assert.ok(xml.toLowerCase().includes('dc2626'), 'the custom zone colour is applied');
});
