import test from 'node:test';
import assert from 'node:assert/strict';
import type { Edge, Node } from 'reactflow';
import { buildInteractiveDiagramHtml } from '../src/services/htmlDiagramExporter.ts';
import { zoneStyleFor, contrastRatio, type ExportBox } from '../src/services/diagramExportGeometry.ts';

/** The same zone uildInteractiveDiagramHtml will derive, as an ExportBox. */
function exportZone(label: string): ExportBox {
  return { id: 'z', kind: 'group', label, category: '', x: 0, y: 0, w: 400, h: 300 };
}

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
  nodes: Array<{
    id: string; name: string; category: string; color: string; icon: string; meta: string;
    x: number; y: number; width: number; height: number;
  }>;
  edges: Array<{ id: string; label: string; color: string; dashed: boolean; points: Array<{ x: number; y: number }> }>;
  groups: Array<{
    id: string; label: string; color: string; bg: string; textColor: string;
    x: number; y: number; width: number; height: number;
  }>;
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

function zone(id: string, x: number, y: number, parentNode?: string): Node {
  return {
    id, type: 'groupNode', position: { x, y }, parentNode,
    style: { width: 400, height: 300 }, data: { label: id },
  };
}

test('a positioned single-service nested graph preserves authored group dimensions and relative coordinates', async () => {
  const nodes = [
    { ...zone('outer', 100, 100), style: { width: 800, height: 600 } },
    zone('inner', 100, 100, 'outer'),
    { ...service('api', 'API'), parentNode: 'inner', position: { x: 75, y: 90 } },
  ];
  const before = structuredClone(nodes);
  const layout = extractLayout((await buildInteractiveDiagramHtml(nodes, [], 'Nested'))!);
  const outer = layout.groups.find(group => group.id === 'outer')!;
  const inner = layout.groups.find(group => group.id === 'inner')!;
  const api = layout.nodes[0];
  assert.deepEqual([inner.x - outer.x, inner.y - outer.y, inner.width, inner.height], [100, 100, 400, 300]);
  assert.deepEqual([api.x - inner.x, api.y - inner.y], [75, 90]);
  assert.deepEqual([outer.width, outer.height], [800, 600]);
  assert.deepEqual(nodes, before);
});

test('equal local service positions in distinct positioned groups are an authored layout', async () => {
  const nodes = [
    zone('east', 100, 100), zone('west', 800, 100),
    { ...service('a', 'API'), parentNode: 'east', position: { x: 50, y: 80 } },
    { ...service('b', 'API'), parentNode: 'west', position: { x: 50, y: 80 } },
  ];
  const layout = extractLayout((await buildInteractiveDiagramHtml(nodes, [], 'Regions'))!);
  for (const [index, groupId] of ['east', 'west'].entries()) {
    const group = layout.groups.find(group => group.id === groupId)!;
    assert.deepEqual([group.width, group.height], [400, 300]);
    assert.deepEqual([layout.nodes[index].x - group.x, layout.nodes[index].y - group.y], [50, 80]);
  }
  assert.equal(layout.nodes[1].x - layout.nodes[0].x, 700);
});

test('an unpositioned nested fallback contains every descendant and preserves empty-group dimensions', async () => {
  const nodes = [
    zone('inner', 0, 0, 'outer'), zone('empty', 0, 0, 'outer'), zone('outer', 0, 0),
    { ...service('a', 'API'), parentNode: 'inner' },
    { ...service('b', 'Database'), parentNode: 'inner' },
  ];
  const layout = extractLayout((await buildInteractiveDiagramHtml(nodes, [
    { id: 'flow', source: 'a', target: 'b' },
  ], 'Fallback'))!);
  const boxes = new Map([...layout.nodes, ...layout.groups].map(box => [box.id, box]));
  for (const node of nodes) {
    const box = boxes.get(node.id)!;
    assert.ok([box.x, box.y, box.width, box.height].every(Number.isFinite));
    if (!node.parentNode) continue;
    const parent = boxes.get(node.parentNode)!;
    assert.ok(box.x >= parent.x && box.y >= parent.y
      && box.x + box.width <= parent.x + parent.width
      && box.y + box.height <= parent.y + parent.height, `${node.id} is contained in ${node.parentNode}`);
  }
  assert.deepEqual([boxes.get('empty')!.width, boxes.get('empty')!.height], [400, 300]);
  assert.equal(layout.groups[0].id, 'outer', 'the resized fallback parent paints behind its descendants');
  assert.notDeepEqual([boxes.get('a')!.x, boxes.get('a')!.y], [boxes.get('b')!.x, boxes.get('b')!.y]);
});

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
  // Carrying the colour and painting with it are different claims: the title
  // ink sat unused in this same object for as long as the renderer read a
  // different field. Assert the emitted renderer actually reaches for it.
  assert.match(
    html!,
    /borderColor\s*=\s*g\.color/,
    'the zone border is not painted from the colour the layout carries',
  );
});

test('a zone drawn around another is still visible in the delivered HTML', async () => {
  // The fill is opaque here -- it is the composited panel colour -- so paint
  // order is what decides whether a nested zone exists at all. Authoring order
  // puts the enclosing zone last, which would bury the tier it was drawn
  // around, label and all.
  const inner = {
    id: 'inner', type: 'groupNode', position: { x: 120, y: 120 },
    style: { width: 160, height: 120 }, data: { label: 'Inner tier' },
  } as unknown as Node;
  const outer = {
    id: 'outer', type: 'groupNode', position: { x: 40, y: 40 },
    style: { width: 420, height: 340 }, data: { label: 'Outer boundary' },
  } as unknown as Node;
  const child = service('c1', 'API');
  (child as { position?: { x: number; y: number } }).position = { x: 150, y: 150 };

  // Authoring order: the boundary is drawn last, as `addGroupBoxAtPosition` appends.
  const html = await buildInteractiveDiagramHtml([inner, outer, child], [], 'Nested zones');
  assert.ok(html);
  const layout = extractLayout(html);
  const order = layout.groups.map(g => g.id);
  assert.deepEqual(order, ['outer', 'inner'], 'the container must be emitted first so it paints behind');

  const encloses = (a: typeof layout.groups[number], b: typeof layout.groups[number]) =>
    a.x <= b.x && a.y <= b.y && a.x + a.width >= b.x + b.width && a.y + a.height >= b.y + b.height;
  for (let i = 0; i < layout.groups.length; i++) {
    for (let j = i + 1; j < layout.groups.length; j++) {
      assert.ok(
        !encloses(layout.groups[j], layout.groups[i]),
        `"${layout.groups[j].label}" is drawn after "${layout.groups[i].label}" and covers it`,
      );
    }
  }
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


test('the interactive HTML zone panel is the tint the canvas shows, not a second dilution', async () => {
  // `g.bg` arrives already composited onto the page by `zoneStyleFor`. The
  // renderer used to append an alpha byte to it, applying the 8-10% tint a
  // second time, so a green zone rendered at under 1% of its accent — the
  // zone colour the author picked was effectively absent from the file.
  const group = {
    id: 'zone-1',
    type: 'groupNode',
    position: { x: 0, y: 0 },
    style: { width: 400, height: 300 },
    data: { label: 'Data Layer' },
  } as Node;
  const child = service('c1', 'API');
  (child as { parentNode?: string }).parentNode = 'zone-1';

  const html = await buildInteractiveDiagramHtml([group, child], [], 'Zone tint');
  assert.ok(html);
  const zone = extractLayout(html!).groups.find((g) => g.id === 'zone-1');
  assert.ok(zone);
  assert.equal(zone!.bg.toLowerCase(), zoneStyleFor(exportZone('Data Layer')).bg.toLowerCase());
  // The renderer must use the value as given. An alpha suffix is the specific
  // defect, and it is invisible in the layout JSON — it is applied at paint.
  assert.ok(
    !/el\.style\.background = g\.bg \+/.test(html!),
    'the zone fill is painted as given rather than re-diluted',
  );
});

test('the interactive HTML zone title uses the readable ink, not the raw accent', async () => {
  // The export drops the canvas header bar and floats the title above the
  // panel on the bare page, where an amber accent is 2.04:1 and a green one
  // 2.41:1. Every other exporter draws this title in `style.text`.
  const zones: Array<[string, string]> = [
    ['Data Layer', '#f8f9fa'],
    ['AI Services', '#f8f9fa'],
    ['Security Perimeter', '#f8f9fa'],
  ];
  for (const [label, page] of zones) {
    const group = {
      id: 'z',
      type: 'groupNode',
      position: { x: 0, y: 0 },
      style: { width: 400, height: 300 },
      data: { label },
    } as Node;
    const child = service('c1', 'API');
    (child as { parentNode?: string }).parentNode = 'z';

    const html = await buildInteractiveDiagramHtml([group, child], [], 'Zone ink');
    assert.ok(html);
    const zone = extractLayout(html!).groups.find((g) => g.id === 'z');
    assert.ok(zone, `${label} is present`);
    assert.equal(zone!.textColor.toLowerCase(), zoneStyleFor(exportZone(label)).text.toLowerCase());
    assert.ok(
      contrastRatio(zone!.textColor, page) >= 4.5,
      `${label}: title ${zone!.textColor} on the page is `
      + `${contrastRatio(zone!.textColor, page).toFixed(2)}:1`,
    );
    // Carrying the ink in the layout is not the same as painting with it: the
    // renderer read `g.color` for years while `textColor` sat unused beside it.
    assert.ok(
      /class="group-label" style="color:' \+ g\.textColor \+ '/.test(html!),
      'the renderer paints the title with the readable ink',
    );
  }
});