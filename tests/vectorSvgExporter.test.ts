// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

/**
 * The SVG export has to be openable by something that is not a browser.
 *
 * The previous implementation photographed the DOM into a single
 * `<foreignObject>`, which only a browser engine can render. It passed every
 * check anyone ran, because every check was run in a browser. These tests
 * assert the shape of the file instead of its appearance: no foreignObject, no
 * HTML namespace, real geometry elements, and icon artwork present as vectors
 * rather than as a reference to a file the recipient does not have.
 */

import test from 'node:test';
import assert from 'node:assert/strict';
import type { Node, Edge } from 'reactflow';

import { exportToSvg, __testing } from '../src/services/vectorSvgExporter.ts';
import { animateEdgeFlow } from '../src/utils/animateEdges.ts';
import { sequenceWorkflowSvg } from '../src/utils/sequenceWorkflow.ts';

const ICON = `<svg id="root-uuid" xmlns="http://www.w3.org/2000/svg" width="18" height="18" viewBox="0 0 18 18"><defs><linearGradient id="grad-uuid" x1="0" y1="0" x2="1" y2="1"><stop offset="0" stop-color="#fff"/></linearGradient></defs><path class="cls-1" d="M0 0 L18 18" fill="url(#grad-uuid)"/></svg>`;

function service(id: string, x: number, y: number, extra: Record<string, unknown> = {}): Node {
  return {
    id,
    type: 'azureNode',
    position: { x, y },
    width: 150,
    height: 75,
    data: { label: id, category: 'compute', ...extra },
  } as Node;
}

function diagram(): { nodes: Node[]; edges: Edge[] } {
  return {
    nodes: [
      {
        id: 'zone',
        type: 'groupNode',
        position: { x: 0, y: 0 },
        style: { width: 600, height: 400 },
        data: { label: 'Production VNet' },
      } as Node,
      service('web', 60, 80),
      service('db', 360, 80, { category: 'databases' }),
    ],
    edges: [
      { id: 'e1', source: 'web', target: 'db', label: 'SQL', data: { connectionType: 'data' } } as Edge,
    ],
  };
}

test('the output contains no foreignObject and no embedded HTML', async () => {
  const { nodes, edges } = diagram();
  const svg = await exportToSvg(nodes, edges);
  assert.ok(!/foreignObject/i.test(svg), 'a foreignObject only a browser can render');
  assert.ok(!/xmlns=["']http:\/\/www\.w3\.org\/1999\/xhtml/.test(svg), 'XHTML namespace present');
  assert.ok(!/<div\b/i.test(svg), 'HTML div in an SVG');
});

test('services and zones are drawn as real geometry', async () => {
  const { nodes, edges } = diagram();
  const svg = await exportToSvg(nodes, edges);
  assert.match(svg, /<rect\b/, 'no rect elements');
  assert.match(svg, /<text\b/, 'no text elements');
  assert.match(svg, /data-zone="zone"/);
  assert.match(svg, /data-service="web"/);
  assert.match(svg, /data-service="db"/);
});

test('edges are real paths with an arrowhead marker', async () => {
  const { nodes, edges } = diagram();
  const svg = await exportToSvg(nodes, edges);
  assert.match(svg, /<path class="react-flow__edge-path"[^>]*\bd="M /);
  assert.match(svg, /<marker id="arrow-0"/);
  assert.match(svg, /marker-end="url\(#arrow-0\)"/);
});

test('the tokens the animation post-processors search for are present', async () => {
  const { nodes, edges } = diagram();
  const svg = await exportToSvg(nodes, edges);
  // sequenceWorkflowSvg maps an edge to its path through this attribute.
  assert.match(svg, /data-testid="rf__edge-e1"/);
  // Both post-processors find paths by this class name.
  assert.match(svg, /class="react-flow__edge-path"/);
});

test('the root size is integral so the sequencer can grow the canvas', async () => {
  const { nodes, edges } = diagram();
  const svg = await exportToSvg(nodes, edges);
  const m = /<svg\b[^>]*\bwidth="(\d+)"[^>]*\bheight="(\d+)"[^>]*\bviewBox="0 0 (\d+) (\d+)"/.exec(svg);
  assert.ok(m, 'root <svg> does not match the sequencer\u2019s width/height/viewBox regex');
  assert.equal(m[1], m[3]);
  assert.equal(m[2], m[4]);
});

test('animateEdgeFlow injects a dot that takes the edge colour', async () => {
  const { nodes, edges } = diagram();
  const svg = await exportToSvg(nodes, edges);
  const animated = animateEdgeFlow(svg);
  assert.match(animated, /<animateMotion\b/, 'no motion injected');
  assert.match(animated, /<mpath[^>]*href="#rfflow-0"/);
  const colour = /data-flow-color="(#[0-9a-fA-F]{6})"/.exec(svg)?.[1];
  assert.ok(colour, 'edge carries no flow colour');
  assert.ok(
    animated.includes(`<circle r="7" fill="${colour}"`),
    'the flow dot fell back to the default palette instead of the edge colour',
  );
});

test('an icon is inlined as vector artwork, not linked or rasterised', async () => {
  const svg = __testing.inlineIcon(ICON, 's0', 10, 20, 32);
  assert.match(svg, /<svg x="10" y="20" width="32" height="32" viewBox="0 0 18 18"/);
  assert.match(svg, /<path[^>]*d="M0 0 L18 18"/);
  assert.ok(!/<image\b/.test(svg), 'icon embedded as a raster image');
  assert.ok(!/base64/.test(svg), 'icon embedded as base64 instead of vectors');
});

test('inlined icon ids and class names are scoped to their tile', async () => {
  const a = __testing.inlineIcon(ICON, 's0', 0, 0, 32);
  const b = __testing.inlineIcon(ICON, 's1', 0, 0, 32);
  assert.match(a, /id="s0-grad-uuid"/);
  assert.match(a, /url\(#s0-grad-uuid\)/);
  assert.match(b, /id="s1-grad-uuid"/);
  assert.match(a, /class="s0-cls-1"/);
  assert.ok(!a.includes('"grad-uuid"'), 'an unscoped id escaped into the document');
  assert.ok(!/id="s0-root-uuid"/.test(a), 'the icon root id should be dropped with the root tag');
});

test('a diagram with no icons still renders every tile', async () => {
  const { nodes, edges } = diagram();
  const svg = await exportToSvg(nodes, edges);
  assert.equal((svg.match(/data-service=/g) ?? []).length, 2);
});

test('dark mode paints the dark surfaces, not the light ones', async () => {
  const { nodes, edges } = diagram();
  const light = await exportToSvg(nodes, edges, { isDarkMode: false });
  const dark = await exportToSvg(nodes, edges, { isDarkMode: true });
  assert.match(light, /fill="#f8fafc"/, 'light page background missing');
  assert.match(dark, /fill="#1e293b"/, 'dark page background missing');
  assert.match(dark, /fill="#27333d"/, 'dark tile surface missing');
  assert.ok(!dark.includes('fill="#f8fafc"'), 'a light background leaked into the dark export');
});

test('background:false leaves the page transparent', async () => {
  const { nodes, edges } = diagram();
  const svg = await exportToSvg(nodes, edges, { background: false });
  assert.ok(!/x="0" y="0" width="\d+" height="\d+" fill="#f8fafc"/.test(svg));
});

test('user text is XML-escaped', async () => {
  // Both spellings, because escaping is case-blind and a test that only knows
  // the lowercase one cannot tell a working escape from one that happens to
  // have been written against the same single spelling. The uppercase tag also
  // covers the whitespace HTML parsers tolerate inside a tag ("< SCRIPT").
  const nodes = [
    service('x', 0, 0, { label: 'A & B <script>' }),
    service('y', 0, 200, { label: '< SCRIPT >' }),
  ];
  const svg = await exportToSvg(nodes, [], { title: 'T&<>"' });
  assert.ok(!/<\s*\/?\s*script/i.test(svg), 'raw markup from a label reached the file');
  assert.match(svg, /A &amp; B &lt;script&gt;/);
  assert.match(svg, /&lt; SCRIPT &gt;/);
  assert.match(svg, /<title>T&amp;&lt;&gt;&quot;<\/title>/);
});

test('the viewport transform brings a far-off drawing inside the viewBox', async () => {
  const nodes = [service('far', -900, -400)];
  const svg = await exportToSvg(nodes, []);

  // Geometry stays in canvas coordinates and a single viewport transform moves
  // it into view, exactly as ReactFlow does. sequenceWorkflowSvg depends on
  // this: it reconstructs node boxes from raw path endpoints and injects the
  // highlight rects *inside* an edge path, so they inherit the same transform.
  // Baking the offset into every coordinate instead would put every highlight
  // one full offset away from the tile it is supposed to be ringing.
  const t = /<g class="react-flow__viewport" transform="translate\((-?[\d.]+),(-?[\d.]+)\)"/.exec(svg);
  assert.ok(t, 'no viewport transform');

  const rect = /data-service="far"><rect x="(-?[\d.]+)" y="(-?[\d.]+)"/.exec(svg);
  assert.ok(rect, 'service rect not found');

  const dims = /<svg\b[^>]*\bwidth="(\d+)"[^>]*\bheight="(\d+)"/.exec(svg);
  assert.ok(dims, 'no root dimensions');

  const screenX = Number(rect[1]) + Number(t[1]);
  const screenY = Number(rect[2]) + Number(t[2]);
  assert.ok(screenX >= 0 && screenX <= Number(dims[1]), `x ${screenX} outside 0..${dims[1]}`);
  assert.ok(screenY >= 0 && screenY <= Number(dims[2]), `y ${screenY} outside 0..${dims[2]}`);
});

test('the accent stripe never becomes a fill on a narrow tile', async () => {
  const narrow = {
    id: 'thin',
    type: 'azureNode',
    position: { x: 0, y: 0 },
    width: 10,
    height: 40,
    data: { label: 'thin', category: 'compute' },
  } as Node;
  const svg = await exportToSvg([narrow], []);
  const stripe = /data-service="thin"><rect[^>]*\/><rect[^>]*width="([\d.]+)"/.exec(svg);
  assert.ok(stripe, 'stripe rect not found');
  assert.ok(Number(stripe[1]) <= 10 / 4 + 0.01, `stripe is ${stripe[1]}px on a 10px tile`);
});

test('long labels wrap rather than overflow the tile', () => {
  const lines = __testing.wrapToWidth('Azure Database for PostgreSQL Flexible Server', 130, 12, 2);
  assert.ok(lines.length <= 2);
  for (const line of lines) {
    assert.ok(__testing.widthPx(line, 12) <= 130, `"${line}" overflows`);
  }
  assert.ok(lines[lines.length - 1].endsWith('…'), 'truncation is not signalled to the reader');
});

test('sequenceWorkflowSvg can grow the canvas for its caption band', async () => {
  const { nodes, edges } = diagram();
  const svg = await exportToSvg(nodes, edges);
  const before = /<svg\b[^>]*\bheight="(\d+)"/.exec(svg);
  assert.ok(before);

  const out = sequenceWorkflowSvg(svg, {
    nodes,
    edges,
    workflow: [{ step: 1, description: 'The web tier queries the database.', services: ['web', 'db'] }],
    stepDurSec: 3,
  });

  const after = /<svg\b[^>]*\bheight="(\d+)"/.exec(out);
  assert.ok(after);
  assert.ok(
    Number(after[1]) > Number(before[1]),
    'the caption band was added without making room for it, so it covers the diagram',
  );

  // The old html-to-image capture emitted `viewBox="-130 -250 3530 1600"`, and
  // the sequencer only rewrites a viewBox anchored at `0 0`. Height and viewBox
  // therefore disagreed and the diagram was cropped. Anchoring at 0 0 fixes it.
  assert.match(out, new RegExp(`viewBox="0 0 \\d+ ${after[1]}"`));
  assert.match(out, /<text[^>]*>Step 1 of 1<\/text>/);
});


/** Every rect drawn inside an edge group, in document order. */
function rectsOfEdges(svg: string): Array<{ x: number; y: number; w: number; h: number }> {
  const out: Array<{ x: number; y: number; w: number; h: number }> = [];
  for (const g of svg.matchAll(/<g class="react-flow__edge"[^>]*>([\s\S]*?)<\/g>/g)) {
    for (const r of g[1].matchAll(/<rect x="(-?[\d.]+)" y="(-?[\d.]+)" width="([\d.]+)" height="([\d.]+)"/g)) {
      out.push({ x: +r[1], y: +r[2], w: +r[3], h: +r[4] });
    }
  }
  return out;
}

/** Every step-badge disc, as a bounding box. */
function badgeBoxes(svg: string): Array<{ x: number; y: number; w: number; h: number }> {
  const out: Array<{ x: number; y: number; w: number; h: number }> = [];
  for (const c of svg.matchAll(/<circle cx="(-?[\d.]+)" cy="(-?[\d.]+)" r="([\d.]+)"/g)) {
    const r = +c[3];
    out.push({ x: +c[1] - r, y: +c[2] - r, w: r * 2, h: r * 2 });
  }
  return out;
}

function hit(a: { x: number; y: number; w: number; h: number }, b: { x: number; y: number; w: number; h: number }) {
  return a.x < b.x + b.w && a.x + a.w > b.x && a.y < b.y + b.h && a.y + a.h > b.y;
}

/**
 * A grid layout puts most arrows well under 200px, and at that length a badge
 * placed at 30% and a chip centred at 50% are within a disc-width of each
 * other. The first render of this exporter printed "1|TTPS" with the numbered
 * disc sitting on the label's leading H.
 */
test('a step badge never lands on a connector label', async () => {
  const nodes: Node[] = [
    service('a', 0, 0),
    service('b', 240, 0),
    service('c', 480, 0),
    service('d', 240, 200),
  ];
  const edges = [
    { id: 'e1', source: 'a', target: 'b', label: 'HTTPS', data: { stepNumber: 1 } },
    { id: 'e2', source: 'b', target: 'c', label: 'TDS', data: { stepNumber: 2 } },
    { id: 'e3', source: 'b', target: 'd', label: 'queue messages', data: { stepNumber: 3 } },
  ] as Edge[];

  const svg = await exportToSvg(nodes, edges);
  const chips = rectsOfEdges(svg);
  const badges = badgeBoxes(svg);

  assert.equal(chips.length, 3, 'expected one chip per labelled edge');
  assert.equal(badges.length, 3, 'expected one badge per numbered edge');

  for (const badge of badges) {
    for (const chip of chips) {
      assert.ok(!hit(badge, chip), `badge at ${badge.x},${badge.y} covers a label chip`);
    }
  }
});

test('two step badges never stack on each other', async () => {
  const nodes: Node[] = [service('hub', 0, 0), service('x', 300, 0), service('y', 300, 140)];
  const edges = [
    { id: 'e1', source: 'hub', target: 'x', data: { stepNumber: 1 } },
    { id: 'e2', source: 'hub', target: 'y', data: { stepNumber: 2 } },
  ] as Edge[];
  const badges = badgeBoxes(await exportToSvg(nodes, edges));
  assert.equal(badges.length, 2);
  assert.ok(!hit(badges[0], badges[1]), 'the two numbered discs overlap');
});

/**
 * `Standard · japaneast · $128/mo` trimmed to a 150px tile used to render as
 * `Standard · japaneast ·…`, which spends four characters saying something was
 * dropped without saying what, and leaves a separator pointing at nothing.
 */
test('a meta subline drops whole segments instead of trailing off', async () => {
  const nodes = [
    service('svc', 0, 0, {
      pricing: { estimatedCost: 128, tier: 'Standard', region: 'japaneast' },
    }),
  ];
  const svg = await exportToSvg(nodes, []);
  const meta = /font-size="9"[^>]*>([^<]*)</.exec(svg);
  assert.ok(meta, 'no meta subline rendered');
  assert.ok(!meta[1].includes('…'), `subline was ellipsised: "${meta[1]}"`);
  assert.ok(!/·\s*$/.test(meta[1]), `subline ends on a dangling separator: "${meta[1]}"`);
  assert.match(meta[1], /Standard/);
});

/**
 * The dotted and ruled canvas backgrounds are CSS gradients on screen — exactly
 * the kind of thing a screenshot preserves for free and a hand-built file drops
 * unless somebody carries it across on purpose.
 */
test('the dotted canvas background is carried into the file', async () => {
  const { nodes, edges } = diagram();
  const svg = await exportToSvg(nodes, edges, { background: 'dots' });
  assert.match(svg, /<pattern id="bg-dots"[^>]*width="20" height="20"/);
  assert.match(svg, /<circle cx="10" cy="10" r="1.2" fill="rgba\(96, 165, 250, 0.32\)"/);
  assert.match(svg, /fill="url\(#bg-dots\)"/);
});

test('the ruled canvas background is carried into the file', async () => {
  const { nodes, edges } = diagram();
  const svg = await exportToSvg(nodes, edges, { background: 'grid' });
  assert.match(svg, /<pattern id="bg-grid"/);
  assert.match(svg, /stroke="rgba\(96, 165, 250, 0.24\)"/);
  assert.match(svg, /fill="url\(#bg-grid\)"/);
});

test('a plain background adds no pattern', async () => {
  const { nodes, edges } = diagram();
  const svg = await exportToSvg(nodes, edges, { background: 'plain' });
  assert.ok(!/<pattern\b/.test(svg), 'a pattern was drawn for a plain background');
  assert.match(svg, /fill="#f8fafc"/);
});