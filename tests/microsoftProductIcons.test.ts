// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

import assert from 'node:assert/strict';
import { existsSync, readFileSync } from 'node:fs';
import path from 'node:path';
import test from 'node:test';
import { fileURLToPath } from 'node:url';
import {
  MICROSOFT_PRODUCT_ICON_CATALOG,
  MICROSOFT_PRODUCT_ICON_PACKAGE_VERSION,
  getMicrosoftProductIconByFileName,
} from '../src/data/microsoftProductIconCatalog';
import { SERVICE_ICON_MAP, resolveServiceIconMapping } from '../src/data/serviceIconMapping';
import { classifyIconPaletteCategory, iconPaletteCategories } from '../src/data/iconCatalog';

const here = path.dirname(fileURLToPath(import.meta.url));
const repoRoot = path.resolve(here, '..');
const iconRoot = path.join(repoRoot, 'Azure_Public_Service_Icons', 'Icons');

function readManifest() {
  return JSON.parse(
    readFileSync(
      path.join(repoRoot, 'Azure_Public_Service_Icons', 'microsoft-product-manifest.json'),
      'utf8',
    ),
  ) as {
    packageVersion: string;
    iconCount: number;
    files: Array<{ path: string; serviceName: string; family: string; sha256: string }>;
  };
}

test('every Power Platform and Dynamics 365 icon exists on disk', () => {
  assert.equal(MICROSOFT_PRODUCT_ICON_CATALOG.length, 24);
  for (const definition of MICROSOFT_PRODUCT_ICON_CATALOG) {
    const iconPath = path.join(iconRoot, definition.category, `${definition.fileName}.svg`);
    assert.ok(
      existsSync(iconPath),
      `${definition.serviceName} is missing its icon at ${definition.category}/${definition.fileName}.svg`,
    );
    const svg = readFileSync(iconPath, 'utf8');
    assert.ok(svg.includes('<svg'), `${definition.fileName}.svg is not an SVG document`);
  }
});

test('both official families are represented', () => {
  const byFamily = new Map<string, number>();
  for (const definition of MICROSOFT_PRODUCT_ICON_CATALOG) {
    byFamily.set(definition.family, (byFamily.get(definition.family) ?? 0) + 1);
  }
  assert.equal(byFamily.get('power-platform'), 8);
  assert.equal(byFamily.get('dynamics-365'), 16);

  // The products the palette must never be missing again.
  const required = [
    'Microsoft Copilot Studio',
    'Microsoft Power Apps',
    'Microsoft Power Automate',
    'Microsoft Power Pages',
    'Microsoft Dataverse',
    'Microsoft Power Platform',
    'Dynamics 365',
    'Dynamics 365 Sales',
    'Dynamics 365 Customer Service',
    'Dynamics 365 Business Central',
  ];
  const serviceNames = new Set(MICROSOFT_PRODUCT_ICON_CATALOG.map(entry => entry.serviceName));
  for (const name of required) {
    assert.ok(serviceNames.has(name), `${name} is missing from the catalog`);
  }
});

test('catalog service names and file names are unique and lookup-safe', () => {
  const serviceNames = new Set<string>();
  const fileNames = new Set<string>();
  for (const definition of MICROSOFT_PRODUCT_ICON_CATALOG) {
    const service = definition.serviceName.toLocaleLowerCase();
    const file = definition.fileName.toLocaleLowerCase();
    assert.ok(!serviceNames.has(service), `duplicate service name: ${definition.serviceName}`);
    assert.ok(!fileNames.has(file), `duplicate file name: ${definition.fileName}`);
    serviceNames.add(service);
    fileNames.add(file);

    assert.match(definition.fileName, /^[a-z0-9-]+$/);
    assert.equal(getMicrosoftProductIconByFileName(definition.fileName), definition);
  }
});

test('every catalog service and alias resolves to its own icon', () => {
  for (const definition of MICROSOFT_PRODUCT_ICON_CATALOG) {
    for (const name of [definition.serviceName, definition.displayName, ...definition.aliases]) {
      const resolved = resolveServiceIconMapping(name);
      assert.ok(resolved, `"${name}" does not resolve to any icon`);
      assert.equal(
        resolved.mapping.iconFile,
        definition.fileName,
        `"${name}" resolved to ${resolved.mapping.iconFile} instead of ${definition.fileName}`,
      );
      assert.equal(resolved.mapping.category, definition.category);
    }
  }
});

test('the new services join the shared service map without shadowing Azure or Fabric entries', () => {
  const seen = new Map<string, string>();
  for (const [serviceName, mapping] of Object.entries(SERVICE_ICON_MAP)) {
    const key = serviceName.toLocaleLowerCase();
    assert.ok(!seen.has(key), `duplicate service key in SERVICE_ICON_MAP: ${serviceName}`);
    seen.set(key, mapping.iconFile);
  }

  for (const definition of MICROSOFT_PRODUCT_ICON_CATALOG) {
    const mapping = SERVICE_ICON_MAP[definition.serviceName];
    assert.ok(mapping, `${definition.serviceName} is not in SERVICE_ICON_MAP`);
    assert.equal(mapping.iconFile, definition.fileName);
    // These products are licensed per user, never through an Azure meter.
    assert.equal(mapping.hasPricingData, false);
  }

  // Existing Azure and Fabric anchors must keep their icons.
  assert.equal(resolveServiceIconMapping('Microsoft Fabric')?.mapping.category, 'fabric');
  assert.equal(
    resolveServiceIconMapping('Microsoft Entra ID')?.mapping.category,
    'identity',
  );
});

test('the palette exposes the two new categories', () => {
  const ids = new Set(iconPaletteCategories.map(category => category.id));
  assert.ok(ids.has('power-platform'));
  assert.ok(ids.has('dynamics-365'));

  for (const id of ['power-platform', 'dynamics-365'] as const) {
    const category = iconPaletteCategories.find(entry => entry.id === id)!;
    assert.ok(category.label.en.length > 0 && category.label.ja.length > 0);
    assert.ok(category.description.en.length > 0 && category.description.ja.length > 0);
    assert.ok(category.keywords.length > 0);
  }

  assert.equal(
    classifyIconPaletteCategory('power platform', 'Copilot Studio', 'copilot-studio'),
    'power-platform',
  );
  assert.equal(
    classifyIconPaletteCategory('dynamics 365', 'Dynamics 365 Sales', 'dynamics-365-sales'),
    'dynamics-365',
  );
});

test('the AI prompts can emit the new icon categories', () => {
  const prompts = [
    'src/services/azureOpenAI.ts',
    'src/services/referenceArchitectureAI.ts',
  ];
  for (const relativePath of prompts) {
    const source = readFileSync(path.join(repoRoot, relativePath), 'utf8');
    assert.ok(
      source.includes('"power platform"'),
      `${relativePath} does not offer the "power platform" category`,
    );
    assert.ok(
      source.includes('"dynamics 365"'),
      `${relativePath} does not offer the "dynamics 365" category`,
    );
  }
});

test('the manifest matches the catalog', () => {
  const manifest = readManifest();
  assert.equal(manifest.packageVersion, MICROSOFT_PRODUCT_ICON_PACKAGE_VERSION);
  assert.equal(manifest.iconCount, MICROSOFT_PRODUCT_ICON_CATALOG.length);
  assert.equal(manifest.files.length, MICROSOFT_PRODUCT_ICON_CATALOG.length);

  const manifestByPath = new Map(manifest.files.map(entry => [entry.path, entry]));
  for (const definition of MICROSOFT_PRODUCT_ICON_CATALOG) {
    const entry = manifestByPath.get(`${definition.category}/${definition.fileName}.svg`);
    assert.ok(entry, `${definition.fileName}.svg is missing from the manifest`);
    assert.equal(entry.serviceName, definition.serviceName);
    assert.equal(entry.family, definition.family);
    assert.match(entry.sha256, /^[a-f0-9]{64}$/);
  }
});
