import { test, expect } from '@playwright/test';
import { createServer, type Server } from 'node:http';
import { readFile, stat } from 'node:fs/promises';
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

test.beforeAll(async () => {
  const built = await stat(path.join(DIST, 'index.html')).catch(() => null);
  test.skip(!built, 'run `npm run build` first: this gate is about the production bundle');

  server = createServer((request, response) => {
    void (async () => {
      const requested = decodeURIComponent((request.url ?? '/').split('?')[0]);
      const candidate = path.join(DIST, requested);
      const isFile = candidate.startsWith(DIST)
        && Boolean((await stat(candidate).catch(() => null))?.isFile());
      // Anything that is not a built file is either the SPA entry or an API the
      // build does not serve; the app already tolerates the latter being absent.
      if (!isFile && requested.startsWith('/api/')) {
        response.writeHead(404, { 'Content-Type': 'application/json' });
        response.end('{}');
        return;
      }
      const target = isFile ? candidate : path.join(DIST, 'index.html');
      const body = await readFile(target);
      response.writeHead(200, {
        'Content-Type': TYPES[path.extname(target)] ?? 'application/octet-stream',
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

  const download = async (command: string): Promise<Buffer> => {
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
  expect((slide.match(/<p:cxnSp>[\s\S]*?<\/p:cxnSp>/g) ?? [])
    .filter((xml) => xml.includes('<a:custGeom>')).length,
  'a <p:cxnSp> with custom geometry makes PowerPoint refuse the file').toBe(0);

  expect(refusals, 'nothing the exports depend on may be refused by the shipped CSP').toEqual([]);
});
