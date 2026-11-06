// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

import assert from 'node:assert/strict';
import { createHash } from 'node:crypto';
import { readFileSync } from 'node:fs';
import path from 'node:path';
import test from 'node:test';

// The synchronizer is JavaScript; load its public validation helpers dynamically.
const syncModule = '../scripts/sync-fabric-icons.mjs';
const { parseCatalog, assertSafeSvg } = await import(syncModule);
const root = path.resolve(import.meta.dirname, '..');

test('sync parses every catalog source version, including the historical filename with a space', () => {
  const parsed = parseCatalog(readFileSync(path.join(root, 'src', 'data', 'fabricIconCatalog.ts'), 'utf8'));
  assert.equal(parsed.entries.length, 117);
  assert.equal(parsed.entries.find((entry: { fileName: string }) => entry.fileName === 'fabric-legacy-synapse').sourceAsset, 'synapse_48_color .svg');
  const versions = new Set(parsed.entries.map((entry: { sourceVersion: string }) => entry.sourceVersion));
  assert.deepEqual(versions, new Set(['12.0.0', '8.2.0', 'local', '6.1.0', '1.1.37', 'article-2026']));
});

test('all 101 pinned article SVGs remain present and hash-verified', () => {
  const lock = JSON.parse(readFileSync(path.join(root, 'scripts', 'fabric-icon-sources.json'), 'utf8'));
  const manifest = JSON.parse(readFileSync(path.join(root, 'Azure_Public_Service_Icons', 'fabric-manifest.json'), 'utf8'));
  assert.equal(lock.articleIcons.length, 101);
  assert.equal(manifest.articleAssets.length, 101);
  assert.equal(new Set(manifest.articleAssets.map((asset: { path: string }) => asset.path)).size, 101);

  for (const source of lock.articleIcons) {
    const asset = manifest.articleAssets.find((item: { title: string }) => item.title === source.title);
    assert.ok(asset, `manifest contains ${source.title}`);
    assert.equal(asset.sourceUrl, source.url);
    assert.equal(asset.sha256, source.sha256);
    const content = readFileSync(path.join(root, 'Azure_Public_Service_Icons', asset.path));
    assert.equal(createHash('sha256').update(content).digest('hex'), source.sha256, source.title);
  }
});

test('sync rejects unsafe paths, duplicate destinations, and active or external SVG content', () => {
  const version = "const FABRIC_ICON_PACKAGE_VERSION = '8.2.0';";
  const entry = "defineFabricIcon('A', 'A', 'safe', 'a_48_item.svg', 'Group', 'item', [])";
  assert.throws(() => parseCatalog(`${version}\n${entry}\n${entry}`), /duplicate/);
  assert.throws(() => parseCatalog(`${version}\n${entry.replace("'safe'", "'../escape'")}`), /Unsafe/);
  assert.throws(() => parseCatalog(`${version}\n${entry.replace('a_48_item.svg', '../escape.svg')}`), /Unsafe/);
  for (const body of [
    '<script>alert(1)</script>', '<foreignObject/>', '<image href="https://example.com/a.svg"/>',
    '<path onclick="alert(1)"/>', '<path fill="url(https://example.com/a.svg)"/>',
    '<style>@import "https://example.com/style"</style>', '<use href="https://example.com/a.svg"/>',
  ]) {
    assert.throws(() => assertSafeSvg(Buffer.from(`<svg>${body}</svg>`), 'unsafe'), /Unsafe/);
  }
  assert.doesNotThrow(() => assertSafeSvg(Buffer.from('<svg><use xlink:href="#shape"/><path fill="url(#gradient)"/><path fill="currentColor"/></svg>'), 'safe'));
});
