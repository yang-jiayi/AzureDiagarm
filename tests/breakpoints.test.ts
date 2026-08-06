// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

import assert from 'node:assert/strict';
import { readdirSync, readFileSync } from 'node:fs';
import test from 'node:test';
import { BREAKPOINTS, MEDIA_QUERIES } from '../src/styles/breakpoints';

const sourceRoot = new URL('../src/', import.meta.url);

function sourceFiles(directory: URL): URL[] {
  return readdirSync(directory, { withFileTypes: true }).flatMap((entry) => {
    const child = new URL(entry.name + (entry.isDirectory() ? '/' : ''), directory);
    return entry.isDirectory() ? sourceFiles(child) : [child];
  });
}

test('semantic media queries use the canonical responsive boundaries', () => {
  assert.deepEqual(BREAKPOINTS, {
    micro: 480,
    compact: 640,
    narrow: 900,
    workspace: 1180,
    wide: 1440,
    lowHeight: 480,
    shortHeight: 600,
  });
  assert.equal(MEDIA_QUERIES.compact, '(max-width: 640px)');
  assert.equal(
    MEDIA_QUERIES.compactOrShortWorkspace,
    '(max-width: 640px), (max-width: 1180px) and (max-height: 600px)',
  );
});

test('React responsive behavior does not embed private breakpoint literals', () => {
  const responsiveConsumers = sourceFiles(sourceRoot).filter((path) => {
    if (!/\.(ts|tsx)$/.test(path.pathname) || path.pathname.endsWith('/breakpoints.ts')) return false;
    const source = readFileSync(path, 'utf8');
    return source.includes('matchMedia(') || source.includes('useMediaQuery(');
  });

  for (const path of responsiveConsumers) {
    const source = readFileSync(path, 'utf8');
    assert.doesNotMatch(source, /(?:matchMedia|useMediaQuery)\(\s*['"][^'"]+['"]/);
  }
});

test('CSS media queries use only canonical width and height boundaries', () => {
  const allowedWidths = new Set([480, 481, 640, 641, 900, 1180, 1440]);
  const allowedHeights = new Set([480, 600]);

  for (const path of sourceFiles(sourceRoot).filter((file) => file.pathname.endsWith('.css'))) {
    const css = readFileSync(path, 'utf8');
    for (const match of css.matchAll(/\((min|max)-(width|height):\s*(\d+)px\)/g)) {
      const value = Number(match[3]);
      const allowed = match[2] === 'width' ? allowedWidths : allowedHeights;
      assert.ok(allowed.has(value), `${path.pathname} uses noncanonical ${match[2]} ${value}px`);
    }
  }
});
