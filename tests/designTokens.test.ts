// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

import assert from 'node:assert/strict';
import { readFileSync } from 'node:fs';
import test from 'node:test';

const tokens = readFileSync(
  new URL('../src/styles/design-tokens.css', import.meta.url),
  'utf8',
);
const hardcodedColor = /#[0-9a-f]{3,8}\b|rgba?\([^)]*\)/gi;
const migratedStyles = [
  '../src/components/AlignmentToolbar.css',
  '../src/components/AzureNode.css',
  '../src/components/CanvasChrome.css',
  '../src/components/CommandPalette.css',
  '../src/components/EditableEdge.css',
  '../src/components/IconPalette.css',
  '../src/components/Legend.css',
  '../src/components/MobileCommandBar.css',
  '../src/components/TitleBlock.css',
  '../src/components/WorkflowPanel.css',
  '../src/components/WorkflowStepper.css',
];

test('shared design tokens cover surfaces, controls, status, spacing, and elevation', () => {
  for (const token of [
    '--azd-color-surface-elevated',
    '--azd-color-control-background',
    '--azd-color-action',
    '--azd-color-success',
    '--azd-color-warning',
    '--azd-color-danger',
    '--azd-color-overlay-backdrop',
    '--azd-shadow-overlay',
    '--azd-space-4',
  ]) {
    assert.match(tokens, new RegExp(`${token}:`));
  }
  assert.match(tokens, /body\.dark-mode\s*\{/);
});

test('core application chrome consumes tokens instead of local color literals', () => {
  for (const path of migratedStyles) {
    const css = readFileSync(new URL(path, import.meta.url), 'utf8');
    assert.equal(css.match(hardcodedColor)?.length ?? 0, 0, `${path} has local colors`);
    assert.match(css, /var\(--azd-/);
  }
});

test('the main shell has substantially migrated to shared design tokens', () => {
  const app = readFileSync(new URL('../src/App.css', import.meta.url), 'utf8');
  assert.ok((app.match(/var\(--azd-/g)?.length ?? 0) >= 120);
  assert.ok((app.match(hardcodedColor)?.length ?? 0) <= 250);
});
