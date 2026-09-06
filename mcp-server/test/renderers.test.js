import assert from 'node:assert/strict';
import test from 'node:test';
import { runInNewContext } from 'node:vm';

import { computeLayout, reflowLayoutForPresentation } from '../dist/layoutEngine.js';
import { renderHtml } from '../dist/htmlRenderer.js';
import { renderSvg, resolveRenderEdgeSemantics } from '../dist/svgRenderer.js';
import {
  connections,
  groups,
  services,
} from '../scripts/test-render-healthcare.mjs';

function edgeLabelRects(svg) {
  return [...svg.matchAll(
    /<g class="edge-label">[\s\S]*?<rect x="([^"]+)" y="([^"]+)" width="([^"]+)" height="([^"]+)"/g,
  )].map((match) => ({
    x: Number(match[1]),
    y: Number(match[2]),
    w: Number(match[3]),
    h: Number(match[4]),
  }));
}

function overlaps(a, b) {
  return !(
    a.x + a.w <= b.x ||
    b.x + b.w <= a.x ||
    a.y + a.h <= b.y ||
    b.y + b.h <= a.y
  );
}

test('SVG renderer keeps dense edge labels separate in both directions', () => {
  for (const direction of ['TB', 'LR']) {
    const layout = computeLayout(services, connections, groups, direction);
    const svg = renderSvg(layout, 'Healthcare Imaging Eventing Architecture');
    const labels = edgeLabelRects(svg);

    assert.equal(labels.length, connections.length);
    for (let left = 0; left < labels.length; left += 1) {
      for (let right = left + 1; right < labels.length; right += 1) {
        assert.equal(
          overlaps(labels[left], labels[right]),
          false,
          `${direction} labels ${left} and ${right} overlap`,
        );
      }
    }
  }
});

test('SVG renderer wraps long titles away from metadata', () => {
  const layout = computeLayout(services, connections, groups, 'TB');
  const svg = renderSvg(
    layout,
    'Healthcare Imaging Eventing Architecture - High Throughput Ordered Events',
    { author: 'Microsoft Scout', generatedBy: 'GPT-5.5', date: '2026-07-07' },
  );
  const title = svg.match(/<text class="diagram-title"[\s\S]*?<\/text>/)?.[0] ?? '';

  assert.ok(title);
  assert.ok((title.match(/<tspan /g) ?? []).length >= 2);
});

test('HTML renderer script-escapes diagram data', () => {
  const injectedName = '</script><img src=x onerror=globalThis.pwned=true>\u2028\u2029';
  const layout = computeLayout(
    [
      { name: injectedName, type: 'App Service' },
      { name: 'Target service', type: 'SQL Database' },
    ],
    [{ from: injectedName, to: 'Target service', label: injectedName }],
    [],
    'TB',
  );
  const html = renderHtml(layout, 'Safe diagram');

  assert.equal((html.match(/<\/script>/g) ?? []).length, 1);
  assert.doesNotMatch(html, /<\/script><img/);
  assert.match(html, /\\u003c\/script>/);
  assert.match(html, /\\u2028\\u2029/);
});

test('both renderers preserve loop and parallel excursions, semantic styles and labels', () => {
  const connections = [
    { from: 'Worker', to: 'Queue', label: 'Publish', type: 'async' },
    { from: 'Worker', to: 'Queue', label: 'Health check', type: 'sync' },
    { from: 'Worker', to: 'Worker', label: 'Retry', type: 'async' },
    { from: 'Worker', to: 'Worker', label: 'Recover', type: 'optional' },
  ];
  for (const direction of ['TB', 'LR']) {
    for (const grouped of [false, true]) {
      const services = [
        { name: 'Worker', type: 'Function App', ...(grouped ? { groupId: 'app' } : {}) },
        { name: 'Queue', type: 'Service Bus', ...(grouped ? { groupId: 'data' } : {}) },
      ];
      const groups = grouped ? [{ id: 'app', label: 'Application' }, { id: 'data', label: 'Data' }] : [];
      const raw = computeLayout(services, connections, groups, direction);
      for (const profile of ['technical', 'presentation']) {
        const layout = profile === 'presentation' ? reflowLayoutForPresentation(raw) : raw;
        const svg = renderSvg(layout, 'Connections', { profile });
        const html = renderHtml(layout, 'Connections', { profile });
        const drawn = [...svg.matchAll(/<g class="edge(?: [^"]*)?"[^>]*>[\s\S]*?<\/g>/g)].map(match => match[0]);
        assert.equal(drawn.length, connections.length);
        const paths = drawn.map(edge => edge.match(/<path d="([^"]*)"/)[1]);
        assert.equal(new Set(paths).size, connections.length);
        assert.match(drawn[0], /stroke="#8764B8"/);
        assert.match(drawn[0], /stroke-dasharray="6,4"/);
        assert.match(drawn[1], /stroke="#0078D4"/);
        assert.doesNotMatch(drawn[1], /stroke-dasharray=/);
        const worker = layout.nodes.find(node => node.name === 'Worker');
        for (const path of paths.slice(2)) {
          const numbers = path.match(/-?\d+(?:\.\d+)?/g).map(Number);
          assert.ok(numbers.some((number, index) => index % 2 === 0
            && (number < worker.x || number > worker.x + worker.width
              || numbers[index + 1] < worker.y || numbers[index + 1] > worker.y + worker.height)),
            'SVG preserves the outside excursion, not just the endpoints');
        }
        for (const connection of connections) assert.ok(svg.includes(connection.label));

        const embedded = JSON.parse(html.match(/const layout = (.+);\nconst showCosts/)[1]);
        assert.equal(embedded.edges.length, connections.length);
        const routeFunction = html.slice(html.indexOf('function orthogonalRoute('), html.indexOf('function roundedOrthoPathD('));
        for (const edge of embedded.edges) {
          const actual = runInNewContext(`(${routeFunction})(edge, direction, [], { w: 10000, h: 10000 })`, { edge, direction });
          assert.deepEqual(JSON.parse(JSON.stringify(actual)), edge.points,
            'the emitted HTML routing function preserves the same exceptional route');
        }
        assert.match(html, /const route = orthogonalRoute\(e,/);
      }
    }
  }
});

test('per-connection identity keeps the presentation label budget for repeated endpoints', () => {
  const layout = computeLayout(
    [{ name: 'Worker', type: 'Function App' }, { name: 'Queue', type: 'Service Bus' }],
    Array.from({ length: 14 }, (_, index) => ({
      from: 'Worker', to: 'Queue', label: `Operation ${Math.floor(index / 2)}`, type: 'async',
    })),
    [],
    'TB',
  );
  assert.equal(new Set(layout.edges.map(edge => edge.key)).size, 14);
  for (const profile of ['technical', 'presentation', 'cost']) {
    const semantics = resolveRenderEdgeSemantics(layout, profile);
    const expectedLabels = profile === 'technical' ? 14 : 12;
    assert.equal(semantics.labeled.size, expectedLabels);
    assert.equal(edgeLabelRects(renderSvg(layout, 'Repeated operations', { profile })).length, expectedLabels);
    const html = renderHtml(layout, 'Repeated operations', { profile });
    const embeddedSemantics = JSON.parse(html.match(/const edgeSemantics = (.+);\n/)[1]);
    assert.equal(embeddedSemantics.labeled.length, expectedLabels);
  }
});

test('renderers include negative loop excursions when given an external positioned layout', () => {
  const layout = {
    nodes: [{ name: 'Worker', type: 'Function App', category: 'compute', x: 40, y: 40, width: 70, height: 70 }],
    groups: [],
    edges: [{
      from: 'Worker', to: 'Worker', label: 'Retry', type: 'async',
      points: [{ x: 40, y: 60 }, { x: -80, y: 60 }, { x: -80, y: 90 }, { x: 40, y: 90 }],
    }],
    width: 150, height: 150, direction: 'TB',
  };
  const original = structuredClone(layout);
  const html = renderHtml(layout, 'Retry');
  const embedded = JSON.parse(html.match(/const layout = (.+);\nconst showCosts/)[1]);
  assert.ok(embedded.width > layout.width);
  assert.ok(embedded.edges[0].points.every(point => point.x >= 0 && point.x <= embedded.width));
  const svg = renderSvg(layout, 'Retry');
  const path = svg.match(/<g class="edge(?: [^"]*)?"[^>]*>[\s\S]*?<path d="([^"]*)"/)[1];
  assert.doesNotMatch(path, /-\d/);
  assert.deepEqual(layout, original);
});
