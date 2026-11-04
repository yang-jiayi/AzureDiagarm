// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

import assert from 'node:assert/strict';
import { createHash } from 'node:crypto';
import { readFileSync } from 'node:fs';
import path from 'node:path';
import test from 'node:test';
import { FABRIC_ICON_CATALOG, getFabricIconByFileName, isFabricUiIconPath } from '../src/data/fabricIconCatalog';
import { isCapacityConsumed, resolveServiceIconMapping, SERVICE_ICON_MAP } from '../src/data/serviceIconMapping';

const root = path.resolve(import.meta.dirname, '..');
const readJson = (file: string) => JSON.parse(readFileSync(path.join(root, file), 'utf8'));

test('all 101 article SVGs are present byte-for-byte, with 110 traceable palette entries', () => {
  const lock = readJson('scripts/fabric-icon-sources.json');
  const manifest = readJson('Azure_Public_Service_Icons/fabric-manifest.json');
  assert.equal(FABRIC_ICON_CATALOG.length, 110);
  assert.equal(manifest.officialFamilyCount, 82);
  assert.equal(manifest.legacyIconCount, 14);
  assert.equal(manifest.uiIconCount, 13);
  assert.equal(manifest.localIconCount, 1);
  assert.equal(lock.articleIcons.length, 101);
  assert.equal(new Set(lock.articleIcons.map((icon: { fileName: string }) => icon.fileName)).size, 101);
  for (const article of lock.articleIcons) {
    const definition = getFabricIconByFileName(article.fileName);
    assert.ok(definition, `Missing article icon: ${article.title}`);
    const bytes = readFileSync(path.join(root, 'Azure_Public_Service_Icons', 'Icons', 'fabric', `${article.fileName}.svg`));
    assert.equal(createHash('sha256').update(bytes).digest('hex'), article.sha256, article.title);
    const record = manifest.files.find((file: { path: string }) => file.path === `fabric/${article.fileName}.svg`);
    assert.equal(record.articleUrl, article.url);
    assert.equal(record.sourceVersion, definition.sourceVersion);
  }
});

test('legacy and UI symbols cannot become AI services or capacity charges', () => {
  const drawingOnly = FABRIC_ICON_CATALOG.filter(icon => icon.kind === 'ui' || icon.kind === 'legacy');
  assert.equal(drawingOnly.length, 27);
  for (const icon of drawingOnly) {
    assert.equal(icon.includeInServiceMap, false, icon.serviceName);
    assert.equal(icon.consumesCapacity, false, icon.serviceName);
    assert.equal(icon.hasPricingData, false, icon.serviceName);
    assert.equal(SERVICE_ICON_MAP[icon.serviceName], undefined, icon.serviceName);
    assert.equal(isCapacityConsumed(icon.serviceName), false, icon.serviceName);
  }
});

test('Shortcuts resolve to the links artwork, never OneLake storage, including old names', () => {
  for (const alias of ['Fabric Links', 'Links', 'Graph Links', 'Shortcuts', 'Fabric Shortcuts', 'OneLake Shortcuts', 'OneLake Shortcut']) {
    const result = resolveServiceIconMapping(alias);
    assert.equal(result?.serviceName, 'Fabric Links', alias);
    assert.equal(result?.mapping.iconFile, 'fabric-item-links', alias);
  }
  assert.equal(resolveServiceIconMapping('OneLake')?.mapping.iconFile, 'onelake');
  assert.equal(resolveServiceIconMapping('OneLake')?.mapping.pricingServiceName, 'OneLake Storage');
});

test('graph aliases use current designs while historical designs remain drawing-only', () => {
  assert.equal(resolveServiceIconMapping('Graph Model Instance')?.mapping.iconFile, 'fabric-item-graph-model');
  assert.equal(resolveServiceIconMapping('Graph Model Instance Queryset')?.mapping.iconFile, 'fabric-item-graph-queryset');
  assert.equal(getFabricIconByFileName('fabric-item-graph-model')?.sourceAsset, 'graph_model_48.svg');
  assert.equal(getFabricIconByFileName('fabric-legacy-graph-model-instance')?.sourceVersion, '6.1.0');
});

test('UI contrast styling derives from saved icon identity, not a build URL', () => {
  assert.equal(isFabricUiIconPath('/Azure_Public_Service_Icons/Icons/fabric/fabric-ui-column-security-filled.svg'), true);
  assert.equal(isFabricUiIconPath('/Azure_Public_Service_Icons/Icons/fabric/fabric-ui-column-security-filled.svg?url'), true);
  assert.equal(isFabricUiIconPath('/Azure_Public_Service_Icons/Icons/fabric/fabric-lakehouse.svg'), false);
  assert.equal(isFabricUiIconPath('data:image/svg+xml,...'), false);
  assert.equal(isFabricUiIconPath(undefined), false);
});

test('all existing stable Fabric paths remain mapped, with the single native-size exception', () => {
  const current = FABRIC_ICON_CATALOG.filter(icon => !['legacy', 'ui'].includes(icon.kind));
  assert.equal(current.length, 83);
  for (const definition of current) {
    if (definition.sourceAsset) {
      assert.match(definition.sourceAsset, definition.fileName === 'fabric-sample-workload' ? /_32_color\.svg$/ : /_48(?:_[a-z-]+)?\.svg$/);
    }
    if (definition.includeInServiceMap) {
      assert.equal(resolveServiceIconMapping(definition.serviceName)?.mapping.iconFile, definition.fileName);
    }
  }
  assert.equal(isCapacityConsumed('Fabric Links'), true);
  assert.equal(isCapacityConsumed('Microsoft Fabric Capacity'), false);
  assert.equal(SERVICE_ICON_MAP['Microsoft Fabric Capacity'].pricingServiceName, 'Microsoft Fabric');
});
