import JSZip from 'jszip';
import type { Edge, Node } from 'reactflow';
import { buildExportRoutes, collectExportBoxes } from '../../src/services/diagramExportGeometry';
import { loadDiagramIcons } from '../../src/services/diagramExportIcons';
import { buildDiagramPptxBlob, buildArchitectureDeckBlob, type ArchitectureDeckOptions } from '../../src/services/pptxExporter';
import { buildVsdxBlob } from '../../src/services/visioVsdxExporter';
import { iconCategories, loadIconsFromCategory } from '../../src/utils/iconLoader';

function check(condition: unknown, message: string): asserts condition {
  if (!condition) throw new Error(message);
}

async function checkPackage(blob: Blob): Promise<void> {
  const zip = await JSZip.loadAsync(await blob.arrayBuffer());
  const relationships = new Map<string, Set<string>>();
  for (const [name, part] of Object.entries(zip.files)) {
    if (!name.endsWith('.rels')) continue;
    const document = new DOMParser().parseFromString(await part.async('string'), 'application/xml');
    check(!document.querySelector('parsererror'), `Malformed Office relationships: ${name}`);
    const ids = new Set<string>();
    for (const relationship of Array.from(document.getElementsByTagNameNS('*', 'Relationship'))) {
      const id = relationship.getAttribute('Id')!;
      check(id && !ids.has(id), `Duplicate or missing relationship ID: ${name}/${id}`);
      ids.add(id);
      if (relationship.getAttribute('TargetMode') === 'External') continue;
      const target = relationship.getAttribute('Target');
      check(target, `Missing relationship target: ${name}/${id}`);
      const directory = name === '_rels/.rels' ? '' : name.replace(/(^|\/)_rels\/[^/]+\.rels$/, '$1');
      const resolved = new URL(target, `https://office.invalid/${directory}`).pathname.slice(1);
      check(zip.file(decodeURIComponent(resolved)), `Missing Office part: ${name} -> ${target}`);
    }
    relationships.set(name, ids);
  }
  for (const [name, part] of Object.entries(zip.files)) {
    if (!/\.(?:xml|rels)$/.test(name)) continue;
    const document = new DOMParser().parseFromString(await part.async('string'), 'application/xml');
    check(!document.querySelector('parsererror'), `Malformed Office XML: ${name}`);
    if (name.endsWith('.xml')) {
      const relPath = name.replace(/([^/]+)$/, '_rels/$1.rels');
      for (const element of Array.from(document.getElementsByTagName('*'))) {
        for (const attribute of Array.from(element.attributes)) {
          if (attribute.namespaceURI !== 'http://schemas.openxmlformats.org/officeDocument/2006/relationships') continue;
          check(relationships.get(relPath)?.has(attribute.value), `Dangling Office relationship: ${name}/${attribute.value}`);
        }
      }
    }
    if (name.startsWith('visio/')) {
      for (const shape of Array.from(document.getElementsByTagNameNS('*', 'Shape'))) {
        const children = Array.from(shape.children).map(child => child.localName);
        const shapes = children.indexOf('Shapes'), text = children.indexOf('Text');
        check(shapes < 0 || text < 0 || shapes < text, 'Visio ShapeSheet sequence requires Shapes before Text');
      }
    } else if (name.startsWith('ppt/slides/')) {
      for (const connector of Array.from(document.getElementsByTagNameNS('*', 'cxnSp'))) {
        check(connector.getElementsByTagNameNS('*', 'custGeom').length === 0,
          'Desktop PowerPoint requires preset geometry for connector shapes');
      }
    }
    const shapeTags = name.startsWith('visio/') ? ['Shape'] : ['cNvPr'];
    for (const tag of shapeTags) {
      const elements = Array.from(document.getElementsByTagNameNS('*', tag));
      const ids = elements.map(element => element.getAttribute(tag === 'Shape' ? 'ID' : 'id'));
      check(new Set(ids).size === ids.length, `Duplicate shape IDs: ${name}`);
    }
  }
}

async function blobData(blob: Blob): Promise<string> {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = () => resolve(String(reader.result).split(',')[1]);
    reader.onerror = () => reject(reader.error);
    reader.readAsDataURL(blob);
  });
}

export async function runOfficeExportFixtures() {
  const catalog = (await Promise.all(iconCategories.map(loadIconsFromCategory))).flat();
  const iconFor = (name: string) => {
    const icon = catalog.find(icon => icon.name.toLowerCase().includes(name));
    check(icon, `Fixture icon unavailable: ${name}`);
    return icon.path;
  };
  const paths = [iconFor('app service'), iconFor('sql database'), iconFor('front door'),
    iconFor('functions'), iconFor('key vault'), iconFor('lakehouse')];
  const service = (id: string, label: string, x: number, y: number, iconIndex: number, parentNode?: string): Node => ({
    id, type: 'azureNode', parentNode, position: { x, y }, width: 210, height: 155,
    data: { label, iconPath: paths[iconIndex % paths.length], category: iconIndex % 2 ? 'databases' : 'compute' },
  });
  const flow = (id: string, source: string, target: string, label: string, data = {}): Edge => ({
    id, source, target, sourceHandle: 'right', targetHandle: 'left', label,
    data: { pathStyle: 'orthogonal', direction: 'forward', ...data },
  });
  const ja = '\u65e5\u672c\u8a9e\u30a2\u30fc\u30ad\u30c6\u30af\u30c1\u30e3';
  const cases: Array<{ name: string; nodes: Node[]; edges: Edge[]; dark: boolean }> = [
    {
      name: 'simple-light', dark: false,
      nodes: [
        { id: 'zone', type: 'groupNode', position: { x: -30, y: -65 }, style: { width: 1120, height: 315 }, data: { label: 'Application platform' } },
        // Role-like IDs must retain independent labels, native groups and glue.
        service('api', 'Azure Front Door', 35, 105, 2, 'zone'),
        service('meta-api', 'App Service', 450, 105, 0, 'zone'),
        service('label-api', 'SQL Database', 855, 105, 1, 'zone'),
      ],
      edges: [flow('e1', 'api', 'meta-api', 'HTTPS / TLS 1.2'), flow('e2', 'meta-api', 'label-api', 'Private connection')],
    },
    {
      name: 'nested-japanese', dark: false,
      nodes: [
        { id: 'region', type: 'groupNode', position: { x: -550, y: -260 }, style: { width: 1200, height: 740 }, data: { label: `${ja} / Japan East` } },
        { id: 'app', type: 'groupNode', parentNode: 'region', position: { x: 45, y: 90 }, style: { width: 500, height: 590 }, data: { label: '\u696d\u52d9\u30a2\u30d7\u30ea\u30b1\u30fc\u30b7\u30e7\u30f3', customColor: { border: '#0078d4', header: '#0078d4' } } },
        { id: 'data', type: 'groupNode', parentNode: 'region', position: { x: 670, y: 90 }, style: { width: 485, height: 590 }, data: { label: '\u30c7\u30fc\u30bf\u57fa\u76e4', customColor: { border: '#10b981', header: '#10b981' } } },
        { id: 'private', type: 'groupNode', parentNode: 'app', position: { x: 24, y: 60 }, style: { width: 452, height: 500 }, data: { label: 'Private subnet' } },
        service('web', '\u304a\u5ba2\u69d8\u5411\u3051\u30dd\u30fc\u30bf\u30eb / App Service', 120, 60, 0, 'private'),
        service('function', '\u975e\u540c\u671f\u30c7\u30fc\u30bf\u51e6\u7406 / Azure Functions', 120, 300, 3, 'private'),
        service('sql', '\u696d\u52d9\u30c7\u30fc\u30bf\u30d9\u30fc\u30b9 / SQL Database', 138, 120, 1, 'data'),
        service('lake', '\u5206\u6790\u7528\u30ec\u30a4\u30af\u30cf\u30a6\u30b9 / Fabric Lakehouse', 138, 360, 5, 'data'),
      ],
      edges: [
        flow('private-query', 'web', 'sql', '\u30c7\u30fc\u30bf\u53c2\u7167 / Private Link', { direction: 'bidirectional' }),
        { ...flow('queue', 'web', 'function', '\u975e\u540c\u671f\u51e6\u7406', { connectionType: 'async' }), sourceHandle: 'bottom', targetHandle: 'top', style: { stroke: '#0078d4', strokeDasharray: '5, 5' } },
        flow('result', 'function', 'lake', '\u96c6\u8a08\u7d50\u679c\u306e\u66f4\u65b0', { direction: 'reverse', labelOffsetY: 24 }),
      ],
    },
  ];
  const denseNodes = Array.from({ length: 30 }, (_, i) => service(`service-${i}`,
    i % 4 === 0 ? `${ja} ${i + 1}` : ['App Service', 'SQL Database', 'Front Door', 'Azure Functions', 'Key Vault', 'Fabric Lakehouse'][i % 6],
    (i % 6) * 300 - 600, Math.floor(i / 6) * 230 - 450, i));
  denseNodes[0].data.pricing = { estimatedCost: 123.45, quantity: 2, isUsageBased: true };
  denseNodes[1].data.pricing = { estimatedCost: null };
  denseNodes[5].data.serviceName = 'Fabric Lakehouse';
  denseNodes[5].data.pricing = { estimatedCost: 0 };
  const denseEdges = Array.from({ length: 25 }, (_, i) => flow(`dense-${i}`, `service-${i}`, `service-${i + 1}`,
    i % 3 ? '' : `Flow ${i + 1}`, { direction: i % 5 === 0 ? 'bidirectional' : 'forward', pathStyle: i % 2 ? 'smooth' : 'orthogonal' }));
  cases.push({ name: 'dense-dark', dark: true, nodes: denseNodes, edges: denseEdges });

  const files: Array<{ name: string; data: string }> = [];
  const report: Array<{ name: string; nodes: number; edges: number; icons: number }> = [];
  for (const fixture of cases) {
    const before = JSON.stringify(fixture);
    const boxes = collectExportBoxes(fixture.nodes);
    const icons = await loadDiagramIcons(fixture.nodes);
    for (const [path, icon] of icons) {
      check(new DataView(icon.bytes.buffer, icon.bytes.byteOffset).getUint32(16) === 256, `Low-resolution icon: ${path}`);
      check(icon.svg?.includes('<svg'), `Vector original missing: ${path}`);
      const image = new Image();
      image.src = icon.dataUrl;
      await image.decode();
      const canvas = document.createElement('canvas');
      canvas.width = 256;
      canvas.height = 256;
      const context = canvas.getContext('2d')!;
      context.drawImage(image, 0, 0);
      const pixels = context.getImageData(0, 0, 256, 256).data;
      let visible = 0;
      for (let i = 3; i < pixels.length; i += 4) if (pixels[i]) visible++;
      check(visible > 50, `Blank export icon: ${path}`);
      check(visible < 256 * 256, `Icon transparency lost: ${path}`);
    }
    const options = {
      diagramName: fixture.name === 'nested-japanese' ? `${ja} - \u696d\u52d9\u30c7\u30fc\u30bf\u57fa\u76e4\u306e\u8a2d\u8a08` : fixture.name,
      author: 'Architecture team', date: '2026-09-05', isDarkMode: fixture.dark,
    };
    const pptx = await buildDiagramPptxBlob({ nodes: fixture.nodes, edges: fixture.edges }, options, icons);
    const vsdx = await buildVsdxBlob(fixture.nodes, fixture.edges, options.diagramName, { isDarkMode: fixture.dark, icons });
    await checkPackage(pptx);
    await checkPackage(vsdx);
    files.push({ name: `${fixture.name}.pptx`, data: await blobData(pptx) }, { name: `${fixture.name}.vsdx`, data: await blobData(vsdx) });
    if (fixture.name === 'dense-dark') {
      const deckOptions: ArchitectureDeckOptions = {
        ...options, diagramName: `${ja} - \u8a2d\u8a08\u30ec\u30d3\u30e5\u30fc\u3068\u904b\u7528\u8a08\u753b`,
        prompt: 'A secure, private application platform with asynchronous processing and analytics.',
        services: fixture.nodes.map((node, i) => ({ name: `${node.data.label} ${i + 1}`,
          category: i % 3 ? 'Application services' : ja, group: i % 2 ? 'Production' : `${ja} / Data platform` })),
        validation: {
          overallScore: 73, summary: `${ja} / Review summary. `.repeat(24),
          pillars: ['Security', 'Reliability', 'Performance', 'Cost optimization', 'Operational excellence'].map(pillar => ({ pillar, score: 73 })),
          findings: Array.from({ length: 8 }, (_, i) => ({ severity: i % 2 ? 'medium' : 'high', category: `Finding ${i + 1}`,
            issue: `${ja}: \u63a5\u7d9a\u8a2d\u5b9a\u3068\u76e3\u8996\u69cb\u6210\u306e\u6539\u5584\u304c\u5fc5\u8981\u3067\u3059\u3002`.repeat(4),
            recommendation: 'Use private connectivity and configure diagnostic logs, alerts and recovery procedures. '.repeat(3) })),
        },
        cost: {
          currency: 'USD', totalMonthly: 12345, annual: 148140, fixedCost: 10000, usageCost: 2345,
          region: 'Japan East', term: 'PAYG', byCategory: [],
          topServices: Array.from({ length: 10 }, (_, i) => ({ serviceName: `${ja} / Resource ${i + 1}`, cost: 500, tier: 'Standard production tier' })),
          regions: Array.from({ length: 8 }, (_, i) => ({ name: `Region ${i + 1}`, monthly: 10000 + i * 500, annual: (10000 + i * 500) * 12, isCurrent: i === 2, isCheapest: i === 0 })),
          unpricedServices: [String(denseNodes[1].data.label)],
        },
      };
      const deck = await buildArchitectureDeckBlob({ nodes: fixture.nodes, edges: fixture.edges }, deckOptions, icons);
      await checkPackage(deck);
      files.push({ name: 'customer-deck.pptx', data: await blobData(deck) });
    }
    check(JSON.stringify(fixture) === before, 'Export mutated the live diagram.');
    report.push({ name: fixture.name, nodes: boxes.size, edges: buildExportRoutes(fixture.edges, boxes).length, icons: icons.size });
  }
  return { files, report, uiDiagram: { nodes: cases[0].nodes, edges: cases[0].edges,
    metadata: { architectureName: 'Office export smoke test', author: 'Architecture team', date: '2026-09-05' } } };
}
