// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

import assert from 'node:assert/strict';
import { readFileSync } from 'node:fs';
import test from 'node:test';
import { contrastRatio } from '../src/services/diagramExportGeometry.ts';

const main = readFileSync(new URL('../src/main.tsx', import.meta.url), 'utf8');
const styles = readFileSync(
  new URL('../src/styles/high-contrast.css', import.meta.url),
  'utf8',
);

test('high-contrast overrides load after the global application styles', () => {
  assert.ok(
    main.indexOf("import './styles/high-contrast.css'") > main.indexOf("import './index.css'"),
  );
});

test('forced-colors mode uses system colors for controls and application surfaces', () => {
  assert.match(styles, /@media \(forced-colors: active\)/);
  assert.match(styles, /--azd-color-surface-app: Canvas/);
  assert.match(styles, /background: ButtonFace !important/);
  assert.match(styles, /color: FieldText !important/);
  assert.match(styles, /outline: 3px solid Highlight !important/);
});

test('forced-colors mode preserves canvas selection and connector visibility', () => {
  assert.match(styles, /\.azure-node\.selected, \.group-node\.selected/);
  assert.match(styles, /\.react-flow__edge\.selected \.react-flow__edge-path/);
  assert.match(styles, /stroke: Highlight !important/);
  assert.match(styles, /\.react-flow__handle/);
});

test('increased contrast also strengthens the regular light and dark themes', () => {
  assert.match(styles, /@media \(prefers-contrast: more\) and \(forced-colors: none\)/);
  assert.match(styles, /body\.dark-mode/);
  assert.match(styles, /border-width: 2px !important/);
});

test('danger callouts meet small-text contrast in light and dark themes', () => {
  const tokens = readFileSync(new URL('../src/styles/design-tokens.css', import.meta.url), 'utf8');
  const primitives = readFileSync(new URL('../src/styles/control-primitives.css', import.meta.url), 'utf8');
  const light = /:root\s*\{([^}]+)\}/.exec(tokens)?.[1];
  const dark = /body\.dark-mode\s*\{([^}]+)\}/.exec(tokens)?.[1];
  const danger = /\.azd-callout--danger\s*\{([^}]+)\}/.exec(primitives)?.[1];
  assert.ok(light && dark && danger);
  const darkDanger = /body\.dark-mode\s+\.azd-callout--danger\s*\{([^}]+)\}/.exec(primitives)?.[1] ?? danger;
  const backgroundToken = /^\s*background:\s*var\((--[\w-]+)\)/m.exec(danger)?.[1];
  assert.ok(backgroundToken);
  for (const [name, theme, rule] of [['light', light, danger], ['dark', dark, darkDanger]]) {
    const foregroundToken = /^\s*color:\s*var\((--[\w-]+)\)/m.exec(rule)?.[1];
    assert.ok(foregroundToken);
    const color = (token: string) => {
      const value = new RegExp(`${token}:\\s*(#[a-f\\d]{6})`, 'i').exec(theme)?.[1];
      assert.ok(value, `${name}: missing ${token}`);
      return value;
    };
    const ratio = contrastRatio(color(foregroundToken), color(backgroundToken));
    assert.ok(ratio >= 4.5, `${name} danger callout contrast ${ratio.toFixed(2)} is below 4.5:1`);
  }
});
