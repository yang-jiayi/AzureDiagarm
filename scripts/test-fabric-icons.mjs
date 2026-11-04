#!/usr/bin/env node
// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

import assert from 'node:assert/strict';
import { mkdir, mkdtemp, readFile, rm } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { dirname, resolve } from 'node:path';
import { fileURLToPath } from 'node:url';
import { build, createServer, preview } from 'vite';
import { chromium, expect } from '@playwright/test';

const root = resolve(dirname(fileURLToPath(import.meta.url)), '..');
const artifacts = process.env.FABRIC_ARTIFACT_DIR;
const snapshotOnly = process.argv.includes('--snapshot-only');
let server;
let browser;
const output = snapshotOnly ? null : await mkdtemp(resolve(tmpdir(), 'fabric-icons-build-'));
try {
  if (snapshotOnly) {
    server = await createServer({ root, server: { port: 0, host: '127.0.0.1' }, logLevel: 'error' });
    await server.listen();
  } else {
    await build({
      root, publicDir: false, logLevel: 'error',
      build: {
        outDir: output, emptyOutDir: true,
        rollupOptions: {
          preserveEntrySignatures: 'strict',
          input: { app: resolve(root, 'index.html'), 'fabric-icons-harness': resolve(root, 'tests', 'fixtures', 'fabric-icons-harness.ts') },
          output: { entryFileNames: chunk => chunk.name === 'fabric-icons-harness' ? 'fabric-icons-harness.js' : 'assets/[name]-[hash].js' },
        },
      },
    });
    server = await preview({ root, build: { outDir: output }, preview: { port: 0, host: '127.0.0.1' }, logLevel: 'error' });
  }
  browser = await chromium.launch({ headless: true });
  const page = await browser.newPage({ viewport: { width: 1440, height: 1100 }, serviceWorkers: 'block' });
  const errors = [];
  page.on('pageerror', error => errors.push(error.message));
  await page.addInitScript(() => {
    localStorage.setItem('azure-diagram-builder.language.v1', 'en');
    localStorage.setItem('azure-diagram-builder.canvasHintDismissed.v1', '1');
  });
  await page.route('**/api/**', route => {
    const path = new URL(route.request().url()).pathname;
    if (path === '/api/access/me') return route.fulfill({ json: { enabled: false, allowed: true } });
    if (path === '/api/feedback/policy') return route.fulfill({ json: { archiveEnabled: false } });
    if (path === '/api/ai/budget') return route.fulfill({ json: { available: false } });
    return route.fulfill({ status: 503, json: { error: 'Offline icon fixture' } });
  });
  await page.goto(server.resolvedUrls.local[0]);
  const palette = page.locator('.icon-palette');
  await palette.getByRole('tab', { name: /^All/ }).click();
  const category = palette.locator('.category-section').filter({
    has: page.locator('.category-title', { hasText: /^Microsoft Fabric$/ }),
  });
  await category.locator('.category-header').click();
  await expect(category.locator('.icon-image').first()).toBeVisible();
  await expect.poll(() => category.locator('.icon-image').evaluateAll(images =>
    images.every(image => image.complete && image.naturalWidth > 0))).toBe(true);
  assert.ok(await category.locator('.icon-image').count() > 0);
  await category.scrollIntoViewIfNeeded();
  if (artifacts) {
    await mkdir(artifacts, { recursive: true });
    await palette.screenshot({ path: resolve(artifacts, snapshotOnly ? 'fabric-palette-before.png' : 'fabric-palette-after.png') });
  }
  if (!snapshotOnly) {
    const manifest = JSON.parse(await readFile(resolve(root, 'Azure_Public_Service_Icons', 'fabric-manifest.json'), 'utf8'));
    const result = await page.evaluate(async () => (await import('/fabric-icons-harness.js')).inspectFabricIcons());
    assert.equal(result.sourceCount, manifest.paletteIconCount);
    assert.equal(result.sourceCount, 110);
    assert.equal(result.rasterizedUi, 13);
    assert.ok(result.inlined > 0, 'Production must exercise Vite data-URL SVGs');
    await expect(category.locator('.icon-count').first()).toHaveText(String(result.categoryCount));
    await expect(category.getByRole('heading', { name: 'OneLake Security UI (13)', exact: true })).toHaveCount(1);
    await expect(category.getByRole('heading', { name: 'Legacy / Deprecated (14)', exact: true })).toHaveCount(1);
    if (artifacts) {
      for (const [heading, file] of [
        ['OneLake Security UI (13)', 'fabric-palette-onelake-ui.png'],
        ['Legacy / Deprecated (14)', 'fabric-palette-legacy.png'],
      ]) {
        await category.getByRole('heading', { name: heading, exact: true }).evaluate(element => element.scrollIntoView({ block: 'start' }));
        await expect.poll(() => category.locator('.icon-image').evaluateAll(images =>
          images.every(image => image.complete && image.naturalWidth > 0))).toBe(true);
        await palette.screenshot({ path: resolve(artifacts, file) });
      }
    }

    // Load the same stable paths used by diagrams saved before this refresh.
    await page.locator('input[accept=".json"]').setInputFiles({
      name: 'saved-fabric-diagram.json', mimeType: 'application/json',
      buffer: Buffer.from(JSON.stringify({
        nodes: [
          { id: 'existing-lake', type: 'azureNode', position: { x: 100, y: 120 }, data: { label: 'Lakehouse', serviceName: 'Lakehouse', category: 'fabric', iconPath: '/Azure_Public_Service_Icons/Icons/fabric/fabric-lakehouse.svg' } },
          { id: 'existing-links', type: 'azureNode', position: { x: 420, y: 120 }, data: { label: 'Links', serviceName: 'Fabric Links', category: 'fabric', iconPath: '/Azure_Public_Service_Icons/Icons/fabric/fabric-item-links.svg' } },
        ], edges: [], titleBlockData: { architectureName: 'Saved Fabric compatibility' },
      })),
    });
    await expect(page.locator('.react-flow__node-azureNode .node-icon')).toHaveCount(2);
    const search = palette.locator('.search-box input');
    for (const name of result.uiNames) {
      await search.fill(name);
      const button = palette.getByRole('button', { name: `Add ${name} to the canvas`, exact: true });
      await expect(button).toBeVisible();
      await expect(button.locator('.fabric-ui-icon')).toHaveCSS('background-color', 'rgb(255, 255, 255)');
      await button.click();
    }
    await search.fill('Graph Model Instance (Legacy)');
    await palette.getByRole('button', { name: 'Add Graph Model Instance (Legacy) to the canvas', exact: true }).click();
    await expect(page.locator('.react-flow__node-azureNode')).toHaveCount(16);
    await expect(page.locator('.react-flow__node-azureNode .fabric-ui-icon')).toHaveCount(13);
    await expect.poll(() => page.locator('.react-flow__node-azureNode .node-icon').evaluateAll(images =>
      images.length === 16 && images.every(image => image.complete && image.naturalWidth > 0))).toBe(true);

    for (const dark of [false, true]) {
      await page.evaluate(value => document.body.classList.toggle('dark-mode', value), dark);
      for (const image of await page.locator('.react-flow__node-azureNode .fabric-ui-icon').all()) {
        await expect(image).toHaveCSS('background-color', 'rgb(255, 255, 255)');
      }
      await search.fill('Column Security (filled)');
      const uiImage = palette.locator('.fabric-ui-icon').first();
      await expect(uiImage).toHaveCSS('background-color', 'rgb(255, 255, 255)');
      // DOM exporters inline images: the metadata-derived class must survive.
      await uiImage.evaluate(async image => {
        const svg = await (await fetch(image.src)).text();
        image.src = `data:image/svg+xml;charset=utf-8,${encodeURIComponent(svg)}`;
      });
      await expect(uiImage).toHaveCSS('background-color', 'rgb(255, 255, 255)');
      if (artifacts) await page.screenshot({ path: resolve(artifacts, `fabric-canvas-${dark ? 'dark' : 'light'}.png`) });
    }
    await page.getByRole('region', { name: 'Architecture canvas' }).focus();
    await page.keyboard.press('Control+K');
    const commands = page.getByTestId('command-palette');
    await commands.getByRole('combobox', { name: 'Search commands and services' }).fill('export png');
    const downloading = page.waitForEvent('download');
    await commands.getByRole('option', { name: /^Export PNG/ }).click();
    await page.getByRole('dialog', { name: 'Review data handling' }).getByRole('button', { name: 'Proceed as shown' }).click();
    const download = await downloading;
    const png = await readFile(await download.path());
    assert.ok(png.length > 10_000, 'Rendered diagram PNG must not be blank');
    if (artifacts) await download.saveAs(resolve(artifacts, 'fabric-diagram-export.png'));
    assert.deepEqual(errors, [], 'Fabric interactions must not throw browser errors');
    console.log(`[test:fabric:browser] ${result.sourceCount} production SVGs rendered (${result.inlined} data URLs), 13 native-export rasters; saved paths, drawing groups, 16 canvas nodes, light/dark contrast and PNG export verified`);
  }
  console.log('[test:fabric:browser] Fabric palette rendered successfully');
} finally {
  await browser?.close();
  if (server?.close) await server.close();
  else if (server?.httpServer) await new Promise((resolve, reject) => server.httpServer.close(error => error ? reject(error) : resolve()));
  if (output) await rm(output, { recursive: true, force: true });
}
