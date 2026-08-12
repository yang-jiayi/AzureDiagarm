import test from 'node:test';
import assert from 'node:assert/strict';
import type { Edge, Node } from 'reactflow';
import { buildVsdxPackage } from '../src/services/visioVsdxExporter.ts';

/**
 * Minimal well-formedness check: tags must nest and close correctly and text
 * must not contain unescaped markup. Enough to catch the mistakes an XML
 * string builder actually makes, without pulling in a parser dependency.
 */
function assertWellFormed(xml: string, label: string): void {
  const stack: string[] = [];
  const tagPattern = /<\/?([A-Za-z_][\w.:-]*)([^>]*?)(\/?)>/g;
  let lastIndex = 0;
  let match: RegExpExecArray | null;
  while ((match = tagPattern.exec(xml)) !== null) {
    const text = xml.slice(lastIndex, match.index);
    assert.equal(/&(?!(amp|lt|gt|quot|apos|#\d+|#x[0-9a-fA-F]+);)/.test(text), false,
      `${label}: unescaped ampersand in text near ${JSON.stringify(text.slice(0, 60))}`);
    lastIndex = tagPattern.lastIndex;

    const raw = match[0];
    const name = match[1];
    if (raw.startsWith('<?') || raw.startsWith('<!')) continue;
    if (raw.startsWith('</')) {
      assert.equal(stack.pop(), name, `${label}: mismatched closing tag </${name}>`);
    } else if (!match[3]) {
      stack.push(name);
    }
  }
  assert.deepEqual(stack, [], `${label}: unclosed tags ${stack.join(', ')}`);
}

function parseConnects(xml: string): Array<Record<string, string>> {
  return Array.from(xml.matchAll(/<Connect\s([^>]*?)\/>/g)).map((match) => {
    const attributes: Record<string, string> = {};
    for (const attribute of match[1].matchAll(/([A-Za-z]+)="([^"]*)"/g)) {
      attributes[attribute[1]] = attribute[2];
    }
    return attributes;
  });
}

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
      iconPath: '/Azure_Public_Service_Icons/Icons/networking/10061-icon-service-Front-Doors.svg',
    },
  } as Node;
}

function zone(id: string, label: string, x: number, y: number): Node {
  return {
    id,
    type: 'groupNode',
    position: { x, y },
    style: { width: 640, height: 360 },
    data: { label },
  } as Node;
}

const nodes: Node[] = [
  zone('zone-1', 'Landing zone', 0, 0),
  service('web', 'App Service', 60, 80, 'zone-1'),
  service('db', 'Azure SQL', 380, 80, 'zone-1'),
  service('monitor', 'Azure Monitor', 900, 500),
];

const edges: Edge[] = [
  { id: 'e1', source: 'web', target: 'db', label: 'SQL 1433' },
  { id: 'e2', source: 'web', target: 'monitor', animated: true },
];

async function pageXml(): Promise<string> {
  const pkg = await buildVsdxPackage(nodes, edges, 'Contoso');
  const page = pkg.parts.find((part) => part.path === 'visio/pages/page1.xml');
  assert.ok(page, 'page1.xml must be present');
  return page.data as string;
}

test('every VSDX XML part is well formed', async () => {
  const pkg = await buildVsdxPackage(nodes, edges, 'Contoso & Partners <Ltd>');
  const xmlParts = pkg.parts.filter((part) => part.path.endsWith('.xml') || part.path.endsWith('.rels'));
  assert.ok(xmlParts.length >= 10);
  for (const part of xmlParts) {
    assertWellFormed(part.data as string, part.path);
  }
});

test('services are groups that carry their icon and shape data', async () => {
  const xml = await pageXml();
  assert.ok(xml.includes('Type="Group"'), 'service tiles must be Visio groups');
  assert.ok(xml.includes('<Shapes>'), 'the group must contain member shapes');
  assert.ok(xml.includes('N="AzureService"'), 'shape data should expose the Azure service name');
  assert.ok(xml.includes('N="Category"'));
  assert.ok(xml.includes('App Service'));
  assert.ok(xml.includes('Landing zone'));
});

test('edges become glued 1-D connectors with a populated Connects table', async () => {
  const xml = await pageXml();

  assert.ok(xml.includes('<Connects>'), 'the page must declare connector glue');
  assert.ok(xml.includes('N="ObjType" V="2"'), 'connectors must be 1-D shapes');
  assert.ok(xml.includes('N="BeginX"') && xml.includes('N="EndX"'));

  const connects = parseConnects(xml);
  assert.equal(connects.length, 4, 'two edges glue at both ends');

  const idsOnPage = new Set(Array.from(xml.matchAll(/<Shape ID="(\d+)"/g)).map((match) => match[1]));
  for (const connect of connects) {
    assert.ok(idsOnPage.has(connect.FromSheet), 'glue references a connector on the page');
    assert.ok(idsOnPage.has(connect.ToSheet), 'glue references a shape on the page');
    assert.ok(['9', '12'].includes(connect.FromPart));
    assert.equal(connect.ToPart, '3', 'glue must target the whole shape so Visio can reroute');
  }

  // Labels ride on the connector, not on a detached chip.
  assert.ok(xml.includes('<Text>SQL 1433</Text>'));
  // Animated edges export as dashed lines.
  assert.ok(xml.includes('N="LinePattern" V="2"'));
});

test('connector geometry maps back onto the routed page points', async () => {
  const xml = await pageXml();

  // Re-apply Visio's own transform (rotate the local rows by Angle around
  // LocPin, then translate to Pin) and compare with the routed page points.
  const shapes = xml.split('<Shape ').filter((chunk) => chunk.includes('N="ObjType" V="2"'));
  assert.ok(shapes.length >= 1, 'at least one connector must be present');

  const cell = (chunk: string, name: string): number => {
    const match = chunk.match(new RegExp(`<Cell N="${name}" V="(-?[\\d.]+)"/>`));
    assert.ok(match, `connector must declare ${name}`);
    return Number(match[1]);
  };

  for (const chunk of shapes) {
    const pinX = cell(chunk, 'PinX');
    const pinY = cell(chunk, 'PinY');
    const locPinX = cell(chunk, 'LocPinX');
    const angle = cell(chunk, 'Angle');
    const beginX = cell(chunk, 'BeginX');
    const beginY = cell(chunk, 'BeginY');
    const endX = cell(chunk, 'EndX');
    const endY = cell(chunk, 'EndY');

    const rows = Array.from(
      chunk.matchAll(/<Row T="(MoveTo|LineTo)" IX="\d+"><Cell N="X" V="(-?[\d.]+)"\/><Cell N="Y" V="(-?[\d.]+)"\/><\/Row>/g),
    ).map((match) => ({ x: Number(match[2]), y: Number(match[3]) }));
    assert.ok(rows.length >= 2, 'a connector needs at least two geometry rows');

    const cos = Math.cos(angle);
    const sin = Math.sin(angle);
    const toPage = (point: { x: number; y: number }) => ({
      x: pinX + (point.x - locPinX) * cos - point.y * sin,
      y: pinY + (point.x - locPinX) * sin + point.y * cos,
    });

    const first = toPage(rows[0]);
    const last = toPage(rows[rows.length - 1]);
    assert.ok(Math.hypot(first.x - beginX, first.y - beginY) < 0.01,
      `first geometry row must land on BeginX/BeginY (got ${first.x}, ${first.y})`);
    assert.ok(Math.hypot(last.x - endX, last.y - endY) < 0.01,
      `last geometry row must land on EndX/EndY (got ${last.x}, ${last.y})`);

    // The elbow must stay axis aligned on the page, not just in local space.
    if (rows.length > 2) {
      const pagePoints = rows.map(toPage);
      for (let index = 1; index < pagePoints.length; index += 1) {
        const previous = pagePoints[index - 1];
        const current = pagePoints[index];
        const axisAligned = Math.abs(previous.x - current.x) < 0.01
          || Math.abs(previous.y - current.y) < 0.01;
        assert.equal(axisAligned, true, 'every connector segment must be horizontal or vertical');
      }
    }
  }
});

test('shape IDs are unique across groups, members and connectors', async () => {
  const xml = await pageXml();
  const ids = Array.from(xml.matchAll(/<Shape ID="(\d+)"/g)).map((match) => match[1]);
  assert.ok(ids.length >= 8);
  assert.equal(new Set(ids).size, ids.length, 'duplicate shape IDs corrupt the drawing');
});

test('the page is centred with layers and never smaller than a Letter sheet', async () => {
  const pkg = await buildVsdxPackage(nodes, edges, 'Contoso');
  assert.ok(pkg.pageWidthIn >= 11);
  assert.ok(pkg.pageHeightIn >= 8.5);
  const pages = pkg.parts.find((part) => part.path === 'visio/pages/pages.xml')!.data as string;
  assert.ok(pages.includes('N="Name" V="Zones"'));
  assert.ok(pages.includes('N="Name" V="Azure services"'));
  assert.ok(pages.includes('N="Name" V="Connections"'));
});

test('an empty diagram still produces a loadable package', async () => {
  const pkg = await buildVsdxPackage([], [], 'Empty');
  const page = pkg.parts.find((part) => part.path === 'visio/pages/page1.xml')!.data as string;
  assertWellFormed(page, 'page1.xml');
  assert.equal(page.includes('<Connects>'), false);
});

/**
 * Reconstruct where Visio will actually draw each connector's text. Visio pins
 * connector text to the midpoint of the begin-to-end chord and measures TxtPin
 * in the connector's own rotated frame, so a label position can only be judged
 * after re-applying that transform.
 */
function connectorLabelBoxes(xml: string): Array<{ text: string; x: number; y: number; w: number; h: number }> {
  const boxes: Array<{ text: string; x: number; y: number; w: number; h: number }> = [];
  for (const block of xml.matchAll(/<Shape [^>]*NameU="Connector\.\d+"[\s\S]*?<\/Shape>/g)) {
    const shape = block[0];
    const text = /<Text>([^<]*)<\/Text>/.exec(shape)?.[1] ?? '';
    if (!text.trim()) continue;
    const pin = /<Cell N="PinX" V="([\d.-]+)"\/>\s*<Cell N="PinY" V="([\d.-]+)"\/>/.exec(shape);
    const txt = /<Cell N="TxtPinX" V="([\d.-]+)"\/>\s*<Cell N="TxtPinY" V="([\d.-]+)"\/>\s*<Cell N="TxtWidth" V="([\d.-]+)"\/>\s*<Cell N="TxtHeight" V="([\d.-]+)"\/>/.exec(shape);
    assert.ok(pin, 'a connector must declare its pin');
    assert.ok(txt, `connector text "${text}" must declare an explicit position, or Visio piles the fan on one point`);
    const theta = Number(/<Cell N="Angle" V="([\d.-]+)"\/>/.exec(shape)?.[1] ?? 0);
    const length = Number(/<Cell N="Width" V="([\d.-]+)"\/>/.exec(shape)?.[1] ?? 0);
    const lx = Number(txt[1]) - length / 2;
    const ly = Number(txt[2]);
    const cx = Number(pin[1]) + lx * Math.cos(theta) - ly * Math.sin(theta);
    const cy = Number(pin[2]) + lx * Math.sin(theta) + ly * Math.cos(theta);
    boxes.push({ text, x: cx - Number(txt[3]) / 2, y: cy - Number(txt[4]) / 2, w: Number(txt[3]), h: Number(txt[4]) });
  }
  return boxes;
}

function overlapArea(
  a: { x: number; y: number; w: number; h: number },
  b: { x: number; y: number; w: number; h: number },
): number {
  const ow = Math.min(a.x + a.w, b.x + b.w) - Math.max(a.x, b.x);
  const oh = Math.min(a.y + a.h, b.y + b.h) - Math.max(a.y, b.y);
  return ow > 0 && oh > 0 ? ow * oh : 0;
}

/** A pair of services joined by `count` parallel numbered hops. */
function parallelHops(count: number): { nodes: Node[]; edges: Edge[] } {
  const fanNodes = [service('src', 'Event Hubs', 0, 0), service('dst', 'Stream Analytics', 460, 0)];
  const fanEdges = Array.from({ length: count }, (unused, index) => ({
    id: `p${index}`,
    source: 'src',
    target: 'dst',
    label: `イベントを Service Bus に発行します ${index + 1}`,
    data: { stepNumber: index + 1 },
  } as Edge));
  return { nodes: fanNodes, edges: fanEdges };
}

async function pageOf(diagramNodes: Node[], diagramEdges: Edge[]): Promise<string> {
  const pkg = await buildVsdxPackage(diagramNodes, diagramEdges, 'Contoso');
  return pkg.parts.find((part) => part.path === 'visio/pages/page1.xml')!.data as string;
}

test('parallel hops write their sentences on separate lines, not one pile', async () => {
  const fan = parallelHops(4);
  const boxes = connectorLabelBoxes(await pageOf(fan.nodes, fan.edges));
  assert.ok(boxes.length >= 4, `all four hops must keep their wording, got ${boxes.length}`);
  for (let i = 0; i < boxes.length; i += 1) {
    for (let j = i + 1; j < boxes.length; j += 1) {
      const smaller = Math.min(boxes[i].w * boxes[i].h, boxes[j].w * boxes[j].h);
      assert.ok(
        overlapArea(boxes[i], boxes[j]) <= 0.25 * smaller,
        `"${boxes[i].text}" and "${boxes[j].text}" are written on top of each other`,
      );
    }
  }
});

/** A grid of hops with two parallel fans hanging off it, as dense pages have. */
function twinLadders(): { nodes: Node[]; edges: Edge[] } {
  const gridNodes: Node[] = [];
  for (let row = 0; row < 2; row += 1) {
    for (let col = 0; col < 3; col += 1) {
      gridNodes.push(service(`n${row}-${col}`, `Service ${row}${col}`, col * 290, row * 180));
    }
  }
  const gridEdges: Edge[] = gridNodes.slice(1).map((node, index) => ({
    id: `w${index}`,
    source: gridNodes[index].id,
    target: node.id,
    label: `ホップ ${index + 1}`,
    data: { stepNumber: index + 1 },
  } as Edge));
  for (let index = 0; index < 4; index += 1) {
    gridEdges.push({
      id: `u${index}`,
      source: 'n1-0',
      target: 'n1-1',
      label: `マネージド ID で参照系を照会します ${index + 1}`,
      data: { stepNumber: 20 + index },
    } as Edge);
  }
  for (let index = 0; index < 6; index += 1) {
    gridEdges.push({
      id: `d${index}`,
      source: 'n0-0',
      target: 'n0-1',
      label: `イベントを Service Bus に発行します ${index + 1}`,
      data: { stepNumber: 40 + index },
    } as Edge);
  }
  return { nodes: gridNodes, edges: gridEdges };
}

test('a connector label is never written off the edge of the sheet', async () => {
  // Visio does not grow a page to fit stray text: whatever falls outside the
  // sheet is simply not there when the file opens. Two fans on one page is the
  // case that finds it, because the second ladder is pushed clear of the first.
  const dense = twinLadders();
  const pkg = await buildVsdxPackage(dense.nodes, dense.edges, 'Contoso');
  const xml = pkg.parts.find((part) => part.path === 'visio/pages/page1.xml')!.data as string;
  const labels = connectorLabelBoxes(xml);
  assert.ok(labels.length >= 5, `the fixture must keep most of its wording, got ${labels.length}`);
  for (const label of labels) {
    assert.ok(
      label.x >= -0.01 && label.y >= -0.01
      && label.x + label.w <= pkg.pageWidthIn + 0.01
      && label.y + label.h <= pkg.pageHeightIn + 0.01,
      `"${label.text}" is drawn outside the ${pkg.pageWidthIn} x ${pkg.pageHeightIn} sheet at ${label.x.toFixed(2)},${label.y.toFixed(2)}`,
    );
  }
});

test('a fan too deep to write out keeps its numbers and the panel explains them', async () => {
  // Twelve hops between one pair of services need a ladder taller than the
  // sheet. Half-hidden sentences are worse than callouts, so the wording is
  // dropped - but only if every one of those callouts is still spelled out.
  const fan = parallelHops(12);
  const xml = await pageOf(fan.nodes, fan.edges);
  const badges = Array.from(xml.matchAll(/<Shape [^>]*NameU="StepBadge\.\d+"[\s\S]*?<Text>(\d+)<\/Text>/g))
    .map((match) => match[1]);
  assert.equal(badges.length, 12, 'every hop must still carry its callout');
  const narrated = new Set(
    Array.from(xml.matchAll(/<Shape [^>]*NameU="LegendText\.\d+"[\s\S]*?<Text>(\d+)\.<\/Text>/g))
      .map((match) => match[1]),
  );
  for (const badge of badges) {
    assert.ok(narrated.has(badge), `callout ${badge} is drawn but no workflow row explains it`);
  }
});


/**
 * A tight grid where the lane between two columns is narrower than a sentence.
 * There is no clear air anywhere on the sheet, so every vertical hop's label
 * lands on a service tile or on the horizontal hop's label beside it.
 *
 * A fan has always had a way out of this — drop the wording, keep the numbers,
 * let the workflow band say it — but a lone label did not, and simply shipped
 * on top of whatever it landed on. Being unable to read the words is the
 * problem; how many hops share the arrow never had anything to do with it.
 */
test('a label with nowhere legible to go gives its sentence to the workflow band', async () => {
  const nodes: Node[] = [];
  for (let r = 0; r < 5; r += 1) {
    for (let c = 0; c < 5; c += 1) {
      nodes.push({
        id: `t${r}${c}`,
        type: 'azureNode',
        position: { x: c * 215, y: r * 150 },
        width: 150,
        height: 75,
        data: { label: `Azure Service ${r}${c}`, serviceName: `Azure Service ${r}${c}` },
      } as Node);
    }
  }
  const edges: Edge[] = [];
  let step = 0;
  for (let r = 0; r < 5; r += 1) {
    for (let c = 0; c + 1 < 5; c += 1) {
      step += 1;
      edges.push({ id: `h${r}${c}`, source: `t${r}${c}`, target: `t${r}${c + 1}`, label: 'writes the order document to Cosmos DB', data: { stepNumber: step, stepDescription: `Step ${step}` } } as Edge);
    }
  }
  for (let r = 0; r + 1 < 5; r += 1) {
    for (let c = 0; c < 5; c += 1) {
      step += 1;
      edges.push({ id: `v${r}${c}`, source: `t${r}${c}`, target: `t${r + 1}${c}`, label: 'queries the read model with a managed identity', data: { stepNumber: step, stepDescription: `Step ${step}` } } as Edge);
    }
  }

  const pkg = await buildVsdxPackage(nodes, edges, 'Contoso');
  const xml = pkg.parts.find((part) => part.path === 'visio/pages/page1.xml')!.data as string;

  // No two sentences may be written on the same spot: whichever is drawn second
  // simply covers the first, and Visio has no autolayout to pull them apart.
  const labels = connectorLabelBoxes(xml);
  for (let i = 0; i < labels.length; i += 1) {
    for (let j = i + 1; j < labels.length; j += 1) {
      const hit = overlapArea(labels[i], labels[j]);
      assert.ok(
        hit <= 0.01,
        `"${labels[i].text}" and "${labels[j].text}" are written on top of each other (${hit.toFixed(3)}sq in)`,
      );
    }
  }

  // And nothing may be quietly deleted to achieve that. Counted, not merely
  // looked for: a drawing repeats its wording, and a rule that stops at the
  // first surviving copy is passed by dropping all the others.
  const fold = (s: string): string => s.toLowerCase().replace(/[\s\u3000]+/g, '').replace(/[.,;:!?、。（）()-]/g, '');
  const shapeText = (prefix: string): string => [
    ...xml.matchAll(new RegExp(`<Shape [^>]*NameU="${prefix}\\.\\d+"[\\s\\S]*?<\\/Shape>`, 'g')),
  ]
    .map((m) => /<Text>([\s\S]*?)<\/Text>/.exec(m[0])?.[1] ?? '')
    .join('\u0000');
  const spoken = fold(`${shapeText('Connector')}\u0000${shapeText('LegendText')}`);
  const need = new Map<string, number>();
  for (const edge of edges) {
    const stem = fold(String(edge.label)).slice(0, 12);
    need.set(stem, (need.get(stem) ?? 0) + 1);
  }
  for (const [stem, count] of need) {
    let found = 0;
    for (let at = spoken.indexOf(stem); at >= 0; at = spoken.indexOf(stem, at + 1)) found += 1;
    assert.ok(found >= count, `the sheet says "${stem}" ${found} times but the author wrote it ${count} times`);
  }
});
