import test from 'node:test';
import assert from 'node:assert/strict';
import JSZip from 'jszip';
import type { Edge, Node } from 'reactflow';
import { buildDiagramSlidePptx } from '../src/services/pptxExporter.ts';

function service(id: string, label: string, x: number, y: number, parentNode?: string): Node {
  return {
    id,
    type: 'azureNode',
    position: { x, y },
    width: 150,
    height: 75,
    ...(parentNode ? { parentNode } : {}),
    data: {
      label,
      serviceName: label,
      iconPath: '/Azure_Public_Service_Icons/Icons/compute/10021-icon-service-Function-Apps.svg',
    },
  } as Node;
}

function group(id: string, label: string, x: number, y: number): Node {
  return {
    id,
    type: 'groupNode',
    position: { x, y },
    style: { width: 520, height: 320 },
    data: { label },
  } as Node;
}

const PIXEL_PNG =
  'data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg==';

async function slideXml(nodes: Node[], edges: Edge[]): Promise<string> {
  const pptx = await buildDiagramSlidePptx(PIXEL_PNG, {
    diagramName: 'Contoso Platform',
    author: 'Tester',
    date: '2026-08-10',
    isDarkMode: false,
    diagram: { nodes, edges },
  });
  const buffer = (await pptx.write({ outputType: 'nodebuffer' })) as Buffer;
  const zip = await JSZip.loadAsync(buffer);
  const file = zip.file('ppt/slides/slide1.xml');
  assert.ok(file, 'slide1.xml must exist in the package');
  return file.async('string');
}

test('PPTX export renders the diagram as native editable shapes, not a picture', async () => {
  const nodes = [
    group('zone-1', 'Application zone', 0, 0),
    service('svc-a', 'API Management', 60, 80, 'zone-1'),
    service('svc-b', 'Azure Functions', 340, 80, 'zone-1'),
  ];
  const edges: Edge[] = [
    { id: 'edge-1', source: 'svc-a', target: 'svc-b', label: 'Invoke worker' },
  ];

  const xml = await slideXml(nodes, edges);

  // The whole point of the change: no rasterised canvas on the diagram slide.
  assert.equal(xml.includes('<p:pic>'), false, 'diagram slide must not embed the canvas screenshot');
  // Each service, the zone, the labels and the connector are separate shapes.
  assert.ok(xml.includes('name="service-svc-a"'));
  assert.ok(xml.includes('name="service-svc-b"'));
  assert.ok(xml.includes('name="zone-zone-1"'));
  assert.ok(xml.includes('name="connector-edge-1"'));
  assert.ok(xml.includes('name="connector-label-edge-1"'));
  // Text lives in real text runs so it can be edited in PowerPoint.
  assert.ok(xml.includes('API Management'));
  assert.ok(xml.includes('Azure Functions'));
  assert.ok(xml.includes('Invoke worker'));
  assert.ok(xml.includes('Application zone'));
  // Rounded rectangles for tiles/zones.
  assert.ok(xml.includes('prst="roundRect"'));
});

test('straight connectors use a line shape and elbows use freeform geometry', async () => {
  const nodes = [
    service('a', 'Front Door', 0, 0),
    service('b', 'App Service', 600, 0),
    service('c', 'Azure SQL', 600, 400),
  ];
  const edges: Edge[] = [
    { id: 'straight', source: 'a', target: 'b' },
    { id: 'elbow', source: 'a', target: 'c' },
  ];

  const xml = await slideXml(nodes, edges);

  assert.ok(xml.includes('prst="line"'), 'aligned nodes should connect with a straight line shape');
  assert.ok(xml.includes('<a:custGeom>'), 'offset nodes should connect with a right-angle freeform path');
  assert.ok(xml.includes('<a:tailEnd type="triangle"'), 'connectors must carry an arrow head');
  assert.ok(xml.includes('<a:lnTo>'), 'the elbow path must contain segments');
});

test('an empty canvas falls back to the captured image', async () => {
  const pptx = await buildDiagramSlidePptx(PIXEL_PNG, {
    diagramName: 'Empty',
    author: 'Tester',
    date: '2026-08-10',
    isDarkMode: true,
    diagram: { nodes: [], edges: [] },
  });
  const buffer = (await pptx.write({ outputType: 'nodebuffer' })) as Buffer;
  const zip = await JSZip.loadAsync(buffer);
  const xml = await zip.file('ppt/slides/slide1.xml')!.async('string');
  assert.ok(xml.includes('<p:pic>'), 'without shapes the slide keeps the screenshot');
});
