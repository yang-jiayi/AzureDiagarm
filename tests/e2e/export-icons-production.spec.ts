import { test, expect } from '@playwright/test';
import { createServer, type Server } from 'node:http';
import { readFile, readdir, stat } from 'node:fs/promises';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import JSZip from 'jszip';

/**
 * The exports, under the conditions the user actually downloads them in.
 *
 * Every other icon test runs against the dev server, and the dev server is the
 * one place this class of bug cannot appear. Vite inlines any asset under
 * `build.assetsInlineLimit` as a `data:` URL -- in a production build that is
 * 1,017 of this app's 1,114 icons -- and the shipped CSP allows `data:` for
 * `img-src` but not for `connect-src`. So the canvas drew every icon while the
 * exporter, which read them with `fetch`, was refused every one of them, and
 * the PowerPoint and Visio files came out with no icons at all. Dev inlines
 * nothing, so dev passed throughout.
 *
 * This serves the real build under the real policy and takes the real route
 * through the UI, which is the only configuration where that could have been
 * caught.
 */
const HERE = path.dirname(fileURLToPath(import.meta.url));
const DIST = path.resolve(HERE, '../../dist');

/** Kept in step with nginx.conf; the missing `data:` in connect-src is the point. */
const CSP = "default-src 'self'; base-uri 'self'; object-src 'none'; frame-ancestors 'none'; "
  + "form-action 'self'; script-src 'self' 'wasm-unsafe-eval'; style-src 'self' 'unsafe-inline'; "
  + "img-src 'self' data: blob:; font-src 'self' data:; connect-src 'self' blob:; "
  + "worker-src 'self' blob:; media-src 'self' blob: data:; manifest-src 'self'";

const TYPES: Record<string, string> = {
  '.html': 'text/html; charset=utf-8',
  '.js': 'text/javascript; charset=utf-8',
  '.css': 'text/css; charset=utf-8',
  '.svg': 'image/svg+xml',
  '.png': 'image/png',
  '.json': 'application/json; charset=utf-8',
  '.woff2': 'font/woff2',
};

const SERVICES = ['Applens', 'Azure Chaos Studio', 'Azure Blockchain Service'];

let server: Server;
let origin = '';

/**
 * Every file under dist/, keyed by the URL that should serve it. Reading the
 * tree once and answering only from this map means no request can name a path:
 * the URL selects an entry or it selects nothing.
 */
async function readBuild(dir: string, prefix = ''): Promise<Map<string, string>> {
  const served = new Map<string, string>();
  for (const entry of await readdir(dir, { withFileTypes: true })) {
    const url = `${prefix}/${entry.name}`;
    if (entry.isDirectory()) {
      for (const [key, value] of await readBuild(path.join(dir, entry.name), url)) {
        served.set(key, value);
      }
    } else if (entry.isFile()) {
      served.set(url, path.join(dir, entry.name));
    }
  }
  return served;
}

test.beforeAll(async () => {
  const built = await stat(path.join(DIST, 'index.html')).catch(() => null);
  // Skipping locally is a convenience; skipping in CI would report green for
  // the one check that covers the shipped bundle, so there it is a failure.
  if (!built && process.env.CI) {
    throw new Error('dist/index.html is missing: CI must build before running the export gate');
  }
  test.skip(!built, 'run `npm run build` first: this gate is about the production bundle');

  const served = await readBuild(DIST);
  const entry = served.get('/index.html')!;

  server = createServer((request, response) => {
    void (async () => {
      const requested = decodeURIComponent((request.url ?? '/').split('?')[0]);
      const target = served.get(requested);
      // Anything that is not a built file is either the SPA entry or an API the
      // build does not serve; the app already tolerates the latter being absent.
      if (!target && requested.startsWith('/api/')) {
        response.writeHead(404, { 'Content-Type': 'application/json' });
        response.end('{}');
        return;
      }
      const body = await readFile(target ?? entry);
      response.writeHead(200, {
        'Content-Type': TYPES[path.extname(target ?? entry)] ?? 'application/octet-stream',
        'Content-Security-Policy': CSP,
      });
      response.end(body);
    })();
  });
  await new Promise<void>((resolve) => server.listen(0, '127.0.0.1', resolve));
  const address = server.address();
  origin = `http://127.0.0.1:${typeof address === 'object' && address ? address.port : 0}`;
});

test.afterAll(async () => {
  await new Promise<void>((resolve) => (server ? server.close(() => resolve()) : resolve()));
});

test('the production build exports icons under the shipped CSP', async ({ page }) => {
  const refusals: string[] = [];
  page.on('console', (message) => {
    if (/content security policy|refused to (connect|load)/i.test(message.text())) {
      refusals.push(message.text());
    }
  });

  await page.goto(origin);
  const canvas = page.getByRole('region', { name: 'Architecture canvas' });
  await expect(canvas).toBeVisible({ timeout: 30_000 });

  const palette = page.getByTestId('command-palette');
  const search = palette.getByRole('combobox', { name: 'Search commands and services' });

  for (const service of SERVICES) {
    await canvas.focus();
    await page.keyboard.press('Control+K');
    await search.fill(service);
    await palette.getByRole('option', { name: new RegExp(`^${service}`) }).first().click();
  }
  await expect(page.locator('.react-flow__node-azureNode')).toHaveCount(SERVICES.length);
  // The canvas draws icons through `img-src`, which permits `data:`. If this is
  // green and the export is empty, the difference is the CSP directive, not the
  // icon library.
  await expect(page.locator('.react-flow__node-azureNode img.node-icon'))
    .toHaveCount(SERVICES.length);

  const download = async (command: string | RegExp): Promise<Buffer> => {
    await page.getByRole('button', { name: 'Export', exact: true }).click();
    const menu = page.getByRole('menu', { name: 'Export options' });
    const pending = page.waitForEvent('download');
    await menu.getByRole('menuitem', { name: command }).click();
    const stream = await (await pending).createReadStream();
    const chunks: Buffer[] = [];
    for await (const chunk of stream) chunks.push(chunk as Buffer);
    return Buffer.concat(chunks);
  };

  const vsdx = await JSZip.loadAsync(await download('Export Visio (VSDX)'));
  const vsdxMedia = Object.values(vsdx.files).filter((entry) => !entry.dir && entry.name.startsWith('visio/media/'));
  const drawnPage = await vsdx.file('visio/pages/page1.xml')?.async('string') ?? '';
  expect(vsdxMedia.length, 'every service tile must embed its icon in the .vsdx')
    .toBe(SERVICES.length);
  expect((drawnPage.match(/<ForeignData /g) ?? []).length, 'each .vsdx icon needs a ForeignData shape')
    .toBe(SERVICES.length);

  const pptx = await JSZip.loadAsync(await download('Export PPTX Slide'));
  const slide = await pptx.file('ppt/slides/slide1.xml')?.async('string') ?? '';
  expect(Object.values(pptx.files).filter((entry) => !entry.dir && entry.name.startsWith('ppt/media/')).length,
    'the delivered deck must carry the tile icons').toBeGreaterThanOrEqual(SERVICES.length);
  expect((slide.match(/<p:pic>/g) ?? []).length, 'each tile must draw its icon on the slide')
    .toBeGreaterThanOrEqual(SERVICES.length);
  // A raster-only deck is the one the user called unusable: the icons turn to
  // mush the moment the slide is projected or printed. The vector original is
  // attached after pptxgenjs has written the package, keyed on a name that the
  // writer escapes, so this can fail in production without failing anywhere
  // else -- and it fails quietly, because the raster is still there and every
  // count above still passes. Follow it to the part, the way PowerPoint does.
  const svgRefs = [...slide.matchAll(/<asvg:svgBlip[^>]*r:embed="([^"]+)"/g)].map((m) => m[1]);
  expect(svgRefs.length, 'every drawn icon must ship its vector original, not just a raster')
    .toBeGreaterThanOrEqual(SERVICES.length);
  const slideRels = await pptx.file('ppt/slides/_rels/slide1.xml.rels')?.async('string') ?? '';
  const relTargets = new Map([...slideRels.matchAll(/Id="([^"]+)"[^>]*Target="([^"]+)"/g)]
    .map((m) => [m[1], m[2]] as const));
  for (const id of new Set(svgRefs)) {
    const target = relTargets.get(id);
    expect(target, `${id} must resolve through the slide's relationships`).toBeTruthy();
    const resolved = `ppt/media/${(target ?? '').split('/').pop()}`;
    expect(pptx.file(resolved), `${id} must point at a part that is in the package`).toBeTruthy();
    expect(await pptx.file(resolved)?.async('string'), 'the vector original must be real SVG')
      .toMatch(/<svg[\s>]/);
  }
  expect((slide.match(/<p:cxnSp>[\s\S]*?<\/p:cxnSp>/g) ?? [])
    .filter((xml) => xml.includes('<a:custGeom>')).length,
  'a <p:cxnSp> with custom geometry makes PowerPoint refuse the file').toBe(0);

  // The primary image export produced no file at all in production: it read the
  // captured canvas back with `fetch(dataUrl)`, which `connect-src` refuses, so
  // the user got an error dialog and nothing else.
  const png = await download('Export PNG');
  expect(png.subarray(0, 8).toString('hex'), 'Export PNG must deliver a real PNG')
    .toBe('89504e470d0a1a0a');
  expect(png.byteLength, 'a three-tile diagram is not a few hundred bytes')
    .toBeGreaterThan(10_000);

  // Draw.io embedded its icons as `data:image/svg+xml,<base64>` -- without the
  // `;base64` marker the payload is percent-decoded, so every icon resolved to
  // the literal base64 text and none of them rendered.
  const drawio = (await download(/^Export Draw/)).toString('utf8');
  const embedded = drawio.match(/image=data:image\/svg\+xml;base64,([A-Za-z0-9+/=]+)/g) ?? [];
  expect(embedded.length, 'every tile must carry an icon Draw.io can decode')
    .toBeGreaterThanOrEqual(SERVICES.length);
  for (const token of embedded) {
    const payload = token.slice(token.indexOf(',') + 1);
    expect(Buffer.from(payload, 'base64').toString('utf8').trimStart(),
      'the decoded icon must be SVG source').toMatch(/^<(\?xml|svg)/);
  }

  expect(refusals, 'nothing the exports depend on may be refused by the shipped CSP').toEqual([]);
});

/**
 * The SVG export, in the two ways it has been broken.
 *
 * It shipped as an `html-to-image` capture, which wraps the whole drawing in a
 * single `<foreignObject>` of XHTML. Browsers render that, so the file looked
 * right to everyone who checked it in a browser, and it opened blank in
 * Inkscape, Illustrator, librsvg, Office and Preview — the tools people pick
 * SVG in order to use. It is now built from the shared export geometry as real
 * elements, and this asserts that it stays that way.
 *
 * It also has to survive the production CSP, because it reads icon source
 * through the same loader that emptied the deck and the Visio sheet.
 */
test('the SVG export is native vector art carrying real icons', async ({ page }) => {
  const refusals: string[] = [];
  page.on('console', (message) => {
    if (/content security policy|refused to (connect|load)/i.test(message.text())) {
      refusals.push(message.text());
    }
  });

  await page.goto(origin);
  const canvas = page.getByRole('region', { name: 'Architecture canvas' });
  await expect(canvas).toBeVisible({ timeout: 30_000 });

  const palette = page.getByTestId('command-palette');
  const search = palette.getByRole('combobox', { name: 'Search commands and services' });
  for (const service of SERVICES) {
    await canvas.focus();
    await page.keyboard.press('Control+K');
    await search.fill(service);
    await palette.getByRole('option', { name: new RegExp(`^${service}`) }).first().click();
  }
  await expect(page.locator('.react-flow__node-azureNode')).toHaveCount(SERVICES.length);

  await page.getByRole('button', { name: 'Export', exact: true }).click();
  const menu = page.getByRole('menu', { name: 'Export options' });
  const pending = page.waitForEvent('download');
  await menu.getByRole('menuitem', { name: /^Export SVG/ }).click();
  const stream = await (await pending).createReadStream();
  const chunks: Buffer[] = [];
  for await (const chunk of stream) chunks.push(chunk as Buffer);
  const svg = Buffer.concat(chunks).toString('utf8');

  expect(svg, 'an SVG only a browser can open is a PNG with extra steps')
    .not.toMatch(/foreignObject/i);
  expect(svg, 'the drawing must be real SVG geometry').toMatch(/<path\b/);
  expect(svg, 'every service needs a tile').toMatch(/data-service=/);
  expect((svg.match(/data-service=/g) ?? []).length).toBe(SERVICES.length);

  // Icon artwork is inlined as nested <svg>, so this is the check that the
  // loader got real source through the CSP rather than falling back to nothing.
  const icons = svg.match(/<svg x="[-\d.]+" y="[-\d.]+"/g) ?? [];
  expect(icons.length, 'every tile must carry its icon as vector art')
    .toBe(SERVICES.length);

  expect(refusals, 'nothing the SVG export depends on may be refused by the shipped CSP')
    .toEqual([]);
});
