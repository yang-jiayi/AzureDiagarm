import test from 'node:test';
import assert from 'node:assert/strict';
import type { Edge, Node } from 'reactflow';
import { buildInteractiveDiagramHtml } from '../src/services/htmlDiagramExporter.ts';
import {
  advanceWidthIn, buildExportRoutes, collectExportBoxes, compactEmptyGutters,
  zoneStyleFor, contrastRatio, type ExportBox,
} from '../src/services/diagramExportGeometry.ts';

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
  edges: Array<{
    id: string; label: string; color: string; dashed: boolean; dashPattern: string; opacity: number;
    bidirectional: boolean; stepNumber?: number;
    points: Array<{ x: number; y: number }>; labelAnchor: { x: number; y: number };
    labelPosition: { x: number; y: number }; stepAnchor: { x: number; y: number };
  }>;
  groups: Array<{
    id: string; label: string; color: string; bg: string; textColor: string;
    x: number; y: number; width: number; height: number;
  }>;
  connectionLegend: Array<{
    type: string; label: string; color: string; dashed: boolean; dashPattern: string;
    opacity: number; hasMixedStyles: boolean;
  }>;
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

test('interactive HTML carries authored connector paint and truthful legend variants', async () => {
  const nodes = [service('a', 'Source'), { ...service('b', 'Target'), position: { x: 400, y: 0 } }];
  const authored: Edge = {
    id: 'authored', source: 'a', target: 'b', data: { connectionType: 'security' },
    style: { stroke: '#006D77', strokeDasharray: '10 2 3 2', opacity: 0.45 },
  };
  const html = await buildInteractiveDiagramHtml(nodes, [authored], 'Authored');
  assert.ok(html);
  const layout = extractLayout(html);
  assert.equal(layout.edges[0].color, '#006d77');
  assert.equal(layout.edges[0].dashPattern, '10, 2, 3, 2');
  assert.equal(layout.edges[0].opacity, 0.45);
  assert.equal(layout.connectionLegend[0].color, '#006d77');
  assert.equal(layout.connectionLegend[0].dashPattern, '10, 2, 3, 2');
  assert.equal(layout.connectionLegend[0].opacity, 0.45);
  const mixed = await buildInteractiveDiagramHtml(nodes, [
    authored, { id: 'default', source: 'a', target: 'b', data: { connectionType: 'security' } },
  ], 'Varied');
  assert.ok(mixed);
  assert.equal(extractLayout(mixed).connectionLegend[0].hasMixedStyles, true);
});

function assertAnnotationsInside(layout: HtmlLayout): void {
  const inside = (x: number, y: number) => {
    assert.ok(Number.isFinite(x) && Number.isFinite(y));
    assert.ok(x >= 0 && x <= layout.width && y >= 0 && y <= layout.height,
      `(${x}, ${y}) must fit ${layout.width} x ${layout.height}`);
  };
  for (const edge of layout.edges) {
    for (const point of edge.points) inside(point.x, point.y);
    assert.ok(edge.labelAnchor, 'the authored label anchor is serialized');
    const at = edge.stepAnchor;
    assert.ok(at, 'the step anchor is serialized independently');
    if (edge.stepNumber !== undefined) {
      inside(at.x - 11, at.y - 11);
      inside(at.x + 11, at.y + 11);
    }
    if (edge.label) {
      const halfWidth = advanceWidthIn(edge.label, 7.5) * 96 / 2 + 2;
      const position = edge.labelPosition;
      inside(position.x - halfWidth, position.y - 12);
      inside(position.x + halfWidth, position.y + 5);
    }
  }
}

test('HTML annotations use translated route anchors for horizontal, vertical, reverse and bidirectional edges', async () => {
  for (const position of [{ x: 300, y: 0 }, { x: 0, y: 300 }]) {
    for (const direction of ['forward', 'reverse', 'bidirectional']) {
      const nodes = [
        { ...service('a', 'API'), width: 150, height: 75 },
        { ...service('b', 'Database'), width: 150, height: 75, position },
      ];
      const edges: Edge[] = [{
        id: 'request', source: 'a', target: 'b',
        data: { label: 'READ', stepNumber: 1, direction },
      }];
      const before = structuredClone({ nodes, edges });
      const boxes = compactEmptyGutters(collectExportBoxes(nodes));
      const route = buildExportRoutes(edges, boxes)[0];
      const html = (await buildInteractiveDiagramHtml(nodes, edges))!;
      const layout = extractLayout(html);
      const dx = layout.nodes[0].x - boxes.get('a')!.x;
      const dy = layout.nodes[0].y - boxes.get('a')!.y;
      assert.deepEqual(layout.edges[0].labelAnchor, { x: route.labelAnchor.x + dx, y: route.labelAnchor.y + dy });
      assert.deepEqual(layout.edges[0].labelPosition, { x: route.labelAnchor.x + dx, y: route.labelAnchor.y + dy - 18 });
      assert.deepEqual(layout.edges[0].stepAnchor, { x: route.labelAnchor.x + dx, y: route.labelAnchor.y + dy });
      assert.deepEqual(layout.edges[0].points, route.points.map(p => ({ x: p.x + dx, y: p.y + dy })));
      assert.equal(layout.edges[0].bidirectional, direction === 'bidirectional');
      assert.notDeepEqual(layout.edges[0].labelAnchor, layout.edges[0].points.at(-1));
      assert.match(html, /const position = e\.labelPosition;/, 'text consumes the serialized label seat');
      assert.match(html, /const mid = e\.stepAnchor;/, 'badges consume their own route anchor');
      assertAnnotationsInside(layout);
      assert.deepEqual({ nodes, edges }, before);
    }
  }
});

test('HTML bounds contain singleton and multiple self-loop routes, labels and step halos', async () => {
  for (const count of [1, 4]) {
    const nodes = [{ ...service('worker', 'Worker'), width: 150, height: 75 }];
    const edges: Edge[] = Array.from({ length: count }, (_, index) => ({
      id: `retry-${index}`, source: 'worker', target: 'worker',
      data: { label: `Retry request after a transient processing failure ${index}`, stepNumber: index + 1 },
    }));
    const html = (await buildInteractiveDiagramHtml(nodes, edges))!;
    const layout = extractLayout(html);
    assert.equal(layout.edges.length, count);
    assertAnnotationsInside(layout);
    assert.equal(new Set(layout.edges.map(edge => edge.labelAnchor.x)).size, count);
    assert.doesNotMatch(html, /overflow:\s*visible/);
    assert.deepEqual([layout.nodes[0].width, layout.nodes[0].height], [150, 75]);
  }
});

test('HTML long loop labels clear nodes and each other while badges stay on their own routes', async () => {
  const overlaps = (a: { x: number; y: number; width: number; height: number },
    b: { x: number; y: number; width: number; height: number }) =>
    a.x < b.x + b.width && a.x + a.width > b.x && a.y < b.y + b.height && a.y + a.height > b.y;
  for (const count of [1, 4]) {
    const labels = [
      'Retry request after a transient processing failure',
      '一時的な処理エラーが発生した場合は要求を再試行してください',
    ];
    for (const label of labels) {
      const nodes = [{ ...service('worker', 'Worker'), width: 150, height: 75 }];
      const edges: Edge[] = Array.from({ length: count }, (_, index) => ({
        id: `retry-${index}`, source: 'worker', target: 'worker',
        data: { label: `${label} ${index + 1}`, stepNumber: index + 1 },
      }));
      const layout = extractLayout((await buildInteractiveDiagramHtml(nodes, edges))!);
      const placed: Array<{ x: number; y: number; width: number; height: number }> = [];
      for (const edge of layout.edges) {
        assert.ok(edge.labelPosition, 'text has a seat independent of the on-route badge');
        const width = advanceWidthIn(edge.label, 7.5) * 96 + 6;
        const rect = { x: edge.labelPosition.x - width / 2, y: edge.labelPosition.y - 13, width, height: 21 };
        assert.ok(layout.nodes.every(node => !overlaps(rect, node)), 'no label is buried under a node');
        assert.ok(placed.every(other => !overlaps(rect, other)), 'loop sentences remain separate');
        assert.ok(rect.x >= 0 && rect.y >= 0 && rect.x + rect.width <= layout.width && rect.y + rect.height <= layout.height);
        placed.push(rect);
        const point = edge.stepAnchor;
        assert.ok(point, 'step has its own route anchor');
        assert.ok(edge.points.slice(1).some((end, index) => {
          const start = edge.points[index];
          return Math.abs(Math.hypot(point.x - start.x, point.y - start.y)
            + Math.hypot(point.x - end.x, point.y - end.y)
            - Math.hypot(end.x - start.x, end.y - start.y)) < 1e-6;
        }), 'the badge is on a segment of its own route');
      }
      assert.equal(layout.nodes[0].width, 150);
      assert.equal(layout.nodes[0].height, 75);
    }
  }
});

test('HTML preserves finite manual label offsets without reversing them or moving authored boxes', async () => {
  const nodes = [
    { ...zone('outer', 100, 100), style: { width: 800, height: 600 } },
    zone('inner', 100, 100, 'outer'),
    { ...service('a', 'API'), parentNode: 'inner', position: { x: 50, y: 60 } },
    { ...service('b', 'Database'), parentNode: 'inner', position: { x: 220, y: 170 } },
  ];
  for (const offset of [{ x: -900, y: -700 }, { x: 900, y: 700 }, { x: NaN, y: Infinity }]) {
    const edges: Edge[] = [{
      id: 'request', source: 'a', target: 'b',
      data: { label: 'Manual placement', stepNumber: 1, direction: 'reverse',
        labelOffsetAuto: false, labelOffsetX: offset.x, labelOffsetY: offset.y },
    }];
    const boxes = compactEmptyGutters(collectExportBoxes(nodes));
    const route = buildExportRoutes(edges, boxes)[0];
    const layout = extractLayout((await buildInteractiveDiagramHtml(nodes, edges))!);
    const first = layout.nodes.find(node => node.id === 'a')!;
    const dx = first.x - boxes.get('a')!.x;
    const dy = first.y - boxes.get('a')!.y;
    assert.deepEqual(layout.edges[0].labelAnchor, {
      x: route.labelAnchor.x + dx + (Number.isFinite(offset.x) ? offset.x : 0),
      y: route.labelAnchor.y + dy + (Number.isFinite(offset.y) ? offset.y : 0),
    });
    assert.deepEqual(layout.edges[0].labelPosition, {
      x: layout.edges[0].labelAnchor.x, y: layout.edges[0].labelAnchor.y - 18,
    }, 'manual label offsets are not replaced by automatic placement');
    assert.deepEqual(layout.edges[0].stepAnchor, {
      x: route.labelAnchor.x + dx, y: route.labelAnchor.y + dy,
    }, 'moving the text does not detach the numbered badge from its route');
    for (const box of [...layout.nodes, ...layout.groups]) {
      const original = boxes.get(box.id)!;
      assert.deepEqual([box.x - dx, box.y - dy, box.width, box.height],
        [original.x, original.y, original.w, original.h]);
    }
    assertAnnotationsInside(layout);
  }
});

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