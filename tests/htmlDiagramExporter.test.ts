import test from 'node:test';
import assert from 'node:assert/strict';
import type { Edge, Node } from 'reactflow';
import { buildInteractiveDiagramHtml } from '../src/services/htmlDiagramExporter.ts';

function service(id: string, label: string, extra: Record<string, unknown> = {}): Node {
  return {
    id,
    type: 'azureNode',
    position: { x: 0, y: 0 },
    data: {
      label,
      serviceName: 'App Service',
      ...extra,
    },
  } as Node;
}

interface HtmlLayout {
  nodes: Array<{ id: string; name: string; category: string; color: string; icon: string; meta: string }>;
  edges: Array<{ id: string; label: string; color: string; dashed: boolean; points: Array<{ x: number; y: number }> }>;
  groups: Array<{ id: string; label: string; color: string }>;
  connectionLegend: Array<{ type: string; label: string; color: string; dashed: boolean }>;
  width: number;
  height: number;
}

function extractLayout(html: string): HtmlLayout {
  const match = html.match(/const layout = (.+?);\n\nlet scale/s);
  assert.ok(match, 'layout JSON should be embedded in the HTML');
  const json = match[1]
    .replace(/\\u003c/g, '<')
    .replace(/\\u2028/g, '\u2028')
    .replace(/\\u2029/g, '\u2029');
  return JSON.parse(json) as HtmlLayout;
}

test('interactive HTML keeps services with duplicate labels distinct', async () => {
  const edge: Edge = {
    id: 'edge-1',
    source: 'service-a',
    target: 'service-b',
  };
  const html = await buildInteractiveDiagramHtml([
    service('service-a', 'Web App'),
    service('service-b', 'Web App'),
  ], [edge], 'Duplicate labels');

  assert.ok(html);
  const layout = extractLayout(html);

  assert.equal(layout.nodes.length, 2);
  assert.deepEqual(layout.nodes.map(node => node.id).sort(), ['service-a', 'service-b']);
  assert.deepEqual(layout.nodes.map(node => node.name), ['Web App', 'Web App']);
  assert.equal(layout.edges.length, 1);
  assert.equal(layout.edges[0].id, 'edge-1');
  assert.ok(layout.edges[0].points.length >= 2);
});

test('interactive HTML embeds no emoji category icons and uses the shared palette', async () => {
  const html = await buildInteractiveDiagramHtml([
    service('svc-ai', 'GPT-4o', {
      category: 'ai + machine learning',
      iconPath: '/Azure_Public_Service_Icons/Icons/ai + machine learning/x.svg',
    }),
  ], [], 'No emoji');

  assert.ok(html);
  // The old exporter injected emoji codepoints and a CATEGORY_ICONS table.
  assert.ok(!/CATEGORY_ICONS/.test(html), 'no emoji icon table should be embedded');
  const layout = extractLayout(html);
  assert.equal(layout.nodes[0].category, 'ai + machine learning');
  assert.match(layout.nodes[0].color, /^#[0-9a-fA-F]{6}$/);
});

test('interactive HTML carries per-connection colour and a connection legend', async () => {
  const securityEdge = {
    id: 'sec',
    source: 'a',
    target: 'b',
    data: { connectionType: 'security' },
  } as Edge;
  const html = await buildInteractiveDiagramHtml([
    service('a', 'Firewall'),
    service('b', 'Gateway'),
  ], [securityEdge], 'Security link');

  assert.ok(html);
  const layout = extractLayout(html);
  const edge = layout.edges.find(e => e.id === 'sec');
  assert.ok(edge);
  // Security connections are red across every export format.
  assert.equal(edge!.color.toLowerCase(), '#dc2626');
  const legendTypes = layout.connectionLegend.map(l => l.type);
  assert.ok(legendTypes.includes('security'), 'security should appear in the connection legend');
});

test('interactive HTML honours a zone custom colour', async () => {
  const group = {
    id: 'zone-1',
    type: 'groupNode',
    position: { x: 0, y: 0 },
    style: { width: 400, height: 300 },
    data: { label: 'Secure Zone', customColor: { border: '#dc2626' } },
  } as Node;
  const child = service('c1', 'API');
  (child as { parentNode?: string }).parentNode = 'zone-1';

  const html = await buildInteractiveDiagramHtml([group, child], [], 'Zone colour');
  assert.ok(html);
  const layout = extractLayout(html);
  const zone = layout.groups.find(g => g.id === 'zone-1');
  assert.ok(zone);
  assert.equal(zone!.color.toLowerCase(), '#dc2626');
});

test('interactive HTML returns null when there are no service nodes', async () => {
  const group = {
    id: 'zone-only',
    type: 'groupNode',
    position: { x: 0, y: 0 },
    data: { label: 'Empty Zone' },
  } as Node;
  const html = await buildInteractiveDiagramHtml([group], [], 'No services');
  assert.equal(html, null);
});

