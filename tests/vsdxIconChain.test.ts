import test from 'node:test';
import assert from 'node:assert/strict';
import type { Node, Edge } from 'reactflow';
import { buildVsdxPackage } from '../src/services/visioVsdxExporter.ts';

/**
 * Visio shows an icon only if a whole chain holds: the shape carries a
 * ForeignData whose child Rel has an r:id, that r:id resolves through the
 * *page's own* rels part to a target, the target resolves to a part that
 * exists, the part's bytes are a real image, and the package declares a
 * content type for that extension. Break any single link and Visio opens the
 * drawing perfectly happily with every icon silently missing -- which is what
 * a user reported, and what counting cannot see: the existing checks assert
 * that media parts exist and that ForeignData shapes exist, never that the one
 * resolves to the other.
 */

const PIXEL_PNG = Uint8Array.from(atob(
  'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg==',
), (c) => c.charCodeAt(0));

const ICONS: Array<[string, string]> = [
  ['front-door', '/Azure_Public_Service_Icons/Icons/networking/10061-icon-service-Front-Doors.svg'],
  ['sql', '/Azure_Public_Service_Icons/Icons/databases/10130-icon-service-SQL-Database.svg'],
  ['storage', '/Azure_Public_Service_Icons/Icons/storage/10086-icon-service-Storage-Accounts.svg'],
];

type Part = { path: string; data: string | Uint8Array };

function partNamed(parts: Part[], path: string): Part | undefined {
  const want = path.replace(/^\//, '');
  return parts.find((p) => p.path.replace(/^\//, '') === want);
}

function textOf(parts: Part[], path: string): string {
  const hit = partNamed(parts, path);
  assert.ok(hit, `the package must contain ${path}`);
  return typeof hit!.data === 'string' ? hit!.data : new TextDecoder().decode(hit!.data);
}

test('every VSDX icon reference resolves to real image bytes the package declares', async () => {
  const nodes: Node[] = ICONS.map(([id, iconPath], i) => ({
    id,
    type: 'azureNode',
    position: { x: i * 260, y: 0 },
    width: 150,
    height: 75,
    data: { label: `Service ${id}`, serviceName: `Service ${id}`, iconPath },
  } as unknown as Node));
  const edges: Edge[] = [
    { id: 'e1', source: 'front-door', target: 'sql' } as Edge,
    { id: 'e2', source: 'sql', target: 'storage' } as Edge,
  ];
  const presetIcons = new Map(ICONS.map(([, iconPath]) => [
    iconPath,
    { bytes: PIXEL_PNG, dataUrl: 'data:image/png;base64,iVBORw0KGgo=', sizePx: 1 },
  ]));

  const { parts } = await buildVsdxPackage(nodes, edges, 'Icon chain', presetIcons as never);
  const pagePath = parts
    .map((p) => p.path.replace(/^\//, ''))
    .find((p) => /^visio\/pages\/page\d+\.xml$/.test(p));
  assert.ok(pagePath, 'the drawing must have a page part');

  const pageXml = textOf(parts as Part[], pagePath!);
  // The id sits on a child <Rel>, not on ForeignData itself.
  const refs = [...pageXml.matchAll(/<ForeignData\b[^>]*>\s*<Rel\b[^>]*r:id="([^"]+)"/g)]
    .map((m) => m[1]);
  assert.equal(refs.length, ICONS.length,
    'each service tile must reference its icon from the page');

  // The rels part governing THIS page, resolved the way a consumer does it:
  // <dir>/_rels/<file>.rels, not a guessed constant.
  const relsPath = pagePath!.replace(/([^/]+)$/, '_rels/$1.rels');
  const relsXml = textOf(parts as Part[], relsPath);
  const rels = new Map(
    [...relsXml.matchAll(/<Relationship\b[^>]*>/g)].map((m) => [
      /Id="([^"]+)"/.exec(m[0])?.[1] ?? '',
      /Target="([^"]+)"/.exec(m[0])?.[1] ?? '',
    ] as const),
  );

  const declared = new Set(
    [...textOf(parts as Part[], '[Content_Types].xml')
      .matchAll(/<Default\s+Extension="([^"]+)"/gi)].map((m) => m[1].toLowerCase()),
  );

  const seen = new Set<string>();
  for (const rid of refs) {
    const target = rels.get(rid);
    assert.ok(target, `ForeignData ${rid} must resolve through ${relsPath}`);

    // Targets are relative to the page part's own directory.
    const resolved = new URL(target!, `file:///${pagePath!.replace(/[^/]+$/, '')}`)
      .pathname.replace(/^\//, '');
    const media = partNamed(parts as Part[], resolved);
    assert.ok(media, `${rid} -> ${target} must resolve to a part that exists`);
    seen.add(resolved);

    const buf = typeof media!.data === 'string'
      ? new TextEncoder().encode(media!.data)
      : media!.data;
    assert.ok(buf.byteLength > 0, `${resolved} must not be empty`);

    const isPng = buf[0] === 0x89 && buf[1] === 0x50 && buf[2] === 0x4e && buf[3] === 0x47;
    const isEmf = buf[0] === 0x01 && buf[1] === 0x00 && buf[2] === 0x00 && buf[3] === 0x00;
    const isJpeg = buf[0] === 0xff && buf[1] === 0xd8;
    assert.ok(isPng || isEmf || isJpeg,
      `${resolved} must carry real image bytes, got ${[...buf.slice(0, 4)].join(',')}`);

    const ext = resolved.split('.').pop()!.toLowerCase();
    assert.ok(declared.has(ext),
      `[Content_Types].xml must declare "${ext}" or Visio drops the image`);
  }

  assert.equal(seen.size, ICONS.length,
    'each tile must get its own icon part, not one shared by all');
});
