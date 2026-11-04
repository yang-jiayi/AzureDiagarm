// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

import assert from 'node:assert/strict';
import { readFileSync } from 'node:fs';
import path from 'node:path';
import test from 'node:test';

// The synchronizer is JavaScript; load its public validation helpers dynamically.
const syncModule = '../scripts/sync-fabric-icons.mjs';
const { parseCatalog, assertSafeSvg } = await import(syncModule);
const root = path.resolve(import.meta.dirname, '..');

test('sync parses every catalog source version, including the historical filename with a space', () => {
  const parsed = parseCatalog(readFileSync(path.join(root, 'src', 'data', 'fabricIconCatalog.ts'), 'utf8'));
  assert.equal(parsed.entries.length, 110);
  assert.equal(parsed.entries.find((entry: { fileName: string }) => entry.fileName === 'fabric-legacy-synapse').sourceAsset, 'synapse_48_color .svg');
  const versions = new Set(parsed.entries.map((entry: { sourceVersion: string }) => entry.sourceVersion));
  assert.deepEqual(versions, new Set(['8.2.0', 'local', '6.1.0', '1.1.37', 'article-2026']));
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
    '<style>@import "https://example.com/style"</style>', '<use href="#external"/>',
  ]) {
    assert.throws(() => assertSafeSvg(Buffer.from(`<svg>${body}</svg>`), 'unsafe'), /Unsafe/);
  }
  assert.doesNotThrow(() => assertSafeSvg(Buffer.from('<svg><path fill="url(#gradient)"/><path fill="currentColor"/></svg>'), 'safe'));
});
