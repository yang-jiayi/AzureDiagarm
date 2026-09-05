import assert from 'node:assert/strict';
import { createServer } from 'node:http';
import { readFile, writeFile, mkdir, rm, stat } from 'node:fs/promises';
import { randomUUID } from 'node:crypto';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import { build } from 'vite';
import { chromium } from 'playwright';
import JSZip from 'jszip';

const root = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');
const work = path.join(root, `.office-export-check-${randomUUID()}`);
await mkdir(work);
const fixturesOnly = process.env.OFFICE_EXPORT_FIXTURES_ONLY === '1';
const output = path.join(work, 'build');
const artifactDirectory = process.env.OFFICE_EXPORT_ARTIFACT_DIR
  ? path.resolve(process.env.OFFICE_EXPORT_ARTIFACT_DIR) : path.join(work, 'artifacts');
let browser;
let server;

try {
  await build({
    configFile: path.join(root, 'vite.config.ts'), root, publicDir: false, logLevel: 'error',
    build: {
      outDir: output, emptyOutDir: true, assetsInlineLimit: 4096,
      rollupOptions: {
        preserveEntrySignatures: 'strict',
        input: {
          ...(!fixturesOnly ? { app: path.join(root, 'index.html') } : {}),
          'office-export-harness': path.join(root, 'tests', 'fixtures', 'office-export-harness.ts'),
        },
        output: { entryFileNames: chunk => chunk.name === 'office-export-harness' ? 'office-export-harness.js' : 'assets/[name]-[hash].js' },
      },
    },
  });
  const types = { '.js': 'text/javascript', '.svg': 'image/svg+xml', '.png': 'image/png', '.css': 'text/css' };
  server = createServer(async (request, response) => {
    try {
      const pathname = decodeURIComponent(new URL(request.url, 'http://localhost').pathname);
      if (pathname === '/export-harness') {
        response.writeHead(200, { 'Content-Type': 'text/html' });
        response.end('<!doctype html><html><head><meta charset="utf-8"><title>Office export fixtures</title></head><body></body></html>');
        return;
      }
      const file = path.resolve(output, '.' + (pathname === '/' ? '/index.html' : pathname));
      if (!file.startsWith(output + path.sep) || !(await stat(file)).isFile()) {
        response.writeHead(404).end();
        return;
      }
      response.writeHead(200, { 'Content-Type': path.extname(file) === '.html' ? 'text/html' : types[path.extname(file)] ?? 'application/octet-stream' });
      response.end(await readFile(file));
    } catch (error) {
      if (error.code === 'ENOENT') response.writeHead(404).end();
      else {
        console.error(error);
        response.writeHead(500).end();
      }
    }
  });
  await new Promise((resolve, reject) => {
    server.once('error', reject);
    server.listen(0, '127.0.0.1', resolve);
  });
  const address = server.address();
  assert.ok(address && typeof address !== 'string');
  browser = await chromium.launch({
    headless: true,
    ...(process.env.OFFICE_EXPORT_BROWSER_CHANNEL ? { channel: process.env.OFFICE_EXPORT_BROWSER_CHANNEL } : {}),
  });
  const page = await browser.newPage({ viewport: { width: 1920, height: 1080 }, acceptDownloads: true });
  page.setDefaultTimeout(30000);
  await page.route('**/*', route => {
    const url = new URL(route.request().url());
    if (url.pathname === '/api/access/me') return route.fulfill({
      json: { enabled: false, authenticated: false, allowed: true, email: null, isAdmin: false },
    });
    if (url.protocol.startsWith('http') && !['localhost', '127.0.0.1'].includes(url.hostname)) return route.abort();
    return route.continue();
  });
  const errors = [];
  page.on('pageerror', error => errors.push(error.message));
  const baseUrl = `http://127.0.0.1:${address.port}`;
  await page.goto(`${baseUrl}/export-harness`, { waitUntil: 'load' });
  const result = await page.evaluate(async () => {
    const { runOfficeExportFixtures } = await import('/office-export-harness.js');
    return runOfficeExportFixtures();
  });
  assert.deepEqual(errors, []);
  assert.equal(result.files.length, 7);
  await mkdir(artifactDirectory, { recursive: true });
  for (const file of result.files) {
    const bytes = Buffer.from(file.data, 'base64');
    assert.ok(bytes.length > 1000, `Unexpectedly small export: ${file.name}`);
    await writeFile(path.join(artifactDirectory, file.name), bytes);
  }
  if (!fixturesOnly) {
    const dialogs = [];
    page.on('dialog', async dialog => {
      dialogs.push(dialog.message());
      await dialog.dismiss();
    });
    await page.addInitScript(() => localStorage.setItem('azure-diagram-builder.language.v1', 'en'));
    await page.goto(baseUrl, { waitUntil: 'load' });
    await page.getByRole('tab', { name: 'Home', exact: true }).click();
    await page.locator('input[type="file"][aria-label="Load diagram"]').setInputFiles({
      name: 'office-fixture.json', mimeType: 'application/json',
      buffer: Buffer.from(JSON.stringify(result.uiDiagram)),
    });
    await page.waitForFunction(() => document.querySelectorAll('.azure-node img').length === 3
      && Array.from(document.querySelectorAll('.azure-node img')).every(image => image.complete && image.naturalWidth > 0));
    await page.waitForTimeout(1000); // The existing saved-diagram loader schedules fitView.
    await page.mouse.move(1500, 800);
    await page.mouse.wheel(0, 250);
    await page.waitForTimeout(250);
    const viewportBefore = await page.locator('.react-flow__viewport').getAttribute('style');
    for (const [name, filename] of [
      [/Export PPTX Slide/, 'ui-slide.pptx'],
      [/Visio/, 'ui-diagram.vsdx'],
      [/Export Customer Deck/, 'ui-deck.pptx'],
    ]) {
      await page.getByRole('button', { name: 'Export', exact: true }).click();
      const downloadPromise = page.waitForEvent('download', { timeout: 90000 });
      await page.getByRole('menuitem', { name }).click();
      const download = await downloadPromise;
      const file = path.join(artifactDirectory, filename);
      await download.saveAs(file);
      const zip = await JSZip.loadAsync(await readFile(file));
      if (filename.endsWith('.pptx')) {
        const slideNames = Object.keys(zip.files).filter(name => /^ppt\/slides\/slide\d+\.xml$/.test(name));
        const slides = await Promise.all(slideNames.map(name => zip.file(name).async('string')));
        assert.ok(slides.some(slide => slide.includes('App Service') && slide.includes('SQL Database')), 'UI must export native diagram text, not a viewport screenshot');
      } else {
        const xml = await zip.file('visio/pages/page1.xml').async('string');
        assert.match(xml, /<Connect FromSheet=/);
        assert.match(xml, /<ForeignData/);
      }
      assert.equal(await page.locator('.react-flow__viewport').getAttribute('style'), viewportBefore, 'Export must not change canvas pan/zoom');
    }
    assert.deepEqual(dialogs, [], 'Office export must not show an error alert');
    assert.deepEqual(errors, [], 'Office export must not trigger a browser exception');
  }
  console.log(JSON.stringify({ fixtures: result.report, exports: result.files.length,
    uiExports: fixturesOnly ? 0 : 3,
    ...(process.env.OFFICE_EXPORT_ARTIFACT_DIR ? { artifacts: artifactDirectory } : {}) }, null, 2));
} finally {
  if (browser) await browser.close();
  if (server) await new Promise(resolve => server.close(resolve));
  await rm(work, { recursive: true, force: true });
}
