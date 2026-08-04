import test from 'node:test';
import assert from 'node:assert/strict';
import type { Edge, Node } from 'reactflow';
import { buildInteractiveDiagramHtml } from '../src/services/htmlDiagramExporter.ts';

function service(id: string, label: string): Node {
  return {
    id,
    type: 'azureNode',
    position: { x: 0, y: 0 },
    data: {
      label,
      serviceName: 'App Service',
    },
  } as Node;
}

test('interactive HTML keeps services with duplicate labels distinct', () => {
  const edge: Edge = {
    id: 'edge-1',
    source: 'service-a',
    target: 'service-b',
  };
  const html = buildInteractiveDiagramHtml([
    service('service-a', 'Web App'),
    service('service-b', 'Web App'),
  ], [edge], 'Duplicate labels');

  assert.ok(html);
  const match = html.match(/const layout = (.+);\n\nconst CATEGORY_COLORS/);
  assert.ok(match);
  const layout = JSON.parse(match[1]) as {
    nodes: Array<{ id: string; name: string }>;
    edges: Array<{
      id: string;
      fromId: string;
      toId: string;
      label: string;
      type: string;
      points: Array<{ x: number; y: number }>;
    }>;
  };

  assert.equal(layout.nodes.length, 2);
  assert.deepEqual(layout.nodes.map(node => node.id).sort(), ['service-a', 'service-b']);
  assert.deepEqual(layout.nodes.map(node => node.name), ['Web App', 'Web App']);
  assert.equal(layout.edges.length, 1);
  assert.equal(layout.edges[0].id, 'edge-1');
  assert.equal(layout.edges[0].fromId, 'service-a');
  assert.equal(layout.edges[0].toId, 'service-b');
  assert.ok(layout.edges[0].points.length >= 2);
});
