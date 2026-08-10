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
