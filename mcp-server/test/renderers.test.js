import assert from 'node:assert/strict';
import test from 'node:test';

import { computeLayout } from '../dist/layoutEngine.js';
import { renderHtml } from '../dist/htmlRenderer.js';
import { renderSvg } from '../dist/svgRenderer.js';
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
