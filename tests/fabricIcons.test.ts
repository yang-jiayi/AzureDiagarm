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

test('all 101 article SVGs are preserved, with 117 traceable palette entries', () => {
  const lock = readJson('scripts/fabric-icon-sources.json');
  const manifest = readJson('Azure_Public_Service_Icons/fabric-manifest.json');
  assert.equal(FABRIC_ICON_CATALOG.length, 117);
  assert.equal(manifest.officialFamilyCount, 88);
  assert.equal(manifest.legacyIconCount, 15);
  assert.equal(manifest.uiIconCount, 13);
  assert.equal(manifest.localIconCount, 1);
  assert.equal(lock.articleIcons.length, 101);
  assert.equal(new Set(lock.articleIcons.map((icon: { fileName: string }) => icon.fileName)).size, 101);
  for (const article of lock.articleIcons) {
    const definition = getFabricIconByFileName(article.fileName);
    assert.ok(definition, `Missing article icon: ${article.title}`);
    const articleAsset = manifest.articleAssets.find((asset: { title: string }) => asset.title === article.title);
    assert.ok(articleAsset, `Missing article provenance: ${article.title}`);
    const bytes = readFileSync(path.join(root, 'Azure_Public_Service_Icons', articleAsset.path));
    assert.equal(createHash('sha256').update(bytes).digest('hex'), article.sha256, article.title);
    const record = manifest.files.find((file: { path: string }) => file.path === `fabric/${article.fileName}.svg`);
    assert.equal(record.sourceVersion, definition.sourceVersion);
    if (articleAsset.path === `Icons/fabric/${article.fileName}.svg`) {
      assert.equal(record.articleUrl, article.url);
    } else {
      assert.equal(articleAsset.activePath, record.path.replace('fabric/', 'Icons/fabric/'));
    }
  }
});

test('legacy and UI symbols cannot become AI services or capacity charges', () => {
  const drawingOnly = FABRIC_ICON_CATALOG.filter(
    icon => icon.kind === 'ui' || icon.kind === 'legacy' || icon.kind === 'concept',
  );
  assert.equal(drawingOnly.length, 34);
  for (const icon of drawingOnly) {
    assert.equal(icon.includeInServiceMap, false, icon.serviceName);
    assert.equal(SERVICE_ICON_MAP[icon.serviceName], undefined, icon.serviceName);
    if (icon.serviceName === 'Fabric Function Set') {
      assert.equal(icon.consumesCapacity, true);
      assert.equal(isCapacityConsumed(icon.serviceName), true);
      assert.equal(resolveServiceIconMapping(icon.serviceName)?.mapping.iconFile, icon.fileName);
    } else {
      assert.equal(icon.consumesCapacity, false, icon.serviceName);
      assert.equal(icon.hasPricingData, false, icon.serviceName);
      assert.equal(isCapacityConsumed(icon.serviceName), false, icon.serviceName);
    }
  }
  assert.equal(SERVICE_ICON_MAP['Fabric Function Set'], undefined);
  assert.equal(resolveServiceIconMapping('Function Set')?.mapping.iconFile, 'fabric-item-function-set');
  assert.equal(resolveServiceIconMapping('Fabric Apps')?.mapping.iconFile, 'fabric-item-fabric-apps');
  assert.equal(resolveServiceIconMapping('Power BI Apps')?.mapping.iconFile, 'fabric-item-apps');
  assert.equal(isCapacityConsumed('Fabric Apps'), true);
  for (const name of [
    'Fabric Database Hub',
    'Rayfin (unclassified)',
    'Fabric Action',
    'Fabric Feature Store',
    'Fabric Managed Identity',
    'Fabric PostgreSQL Database',
  ]) {
    const icon = FABRIC_ICON_CATALOG.find(entry => entry.serviceName === name);
    assert.equal(icon?.kind, 'concept', `${name} must remain a drawing-only concept`);
    assert.equal(icon?.includeInServiceMap, false, `${name} must stay out of service generation`);
    assert.equal(icon?.costRange, undefined, `${name} must not assert unverified billing semantics`);
    assert.equal(SERVICE_ICON_MAP[name], undefined, `${name} must remain drawing-only`);
    assert.equal(isCapacityConsumed(name), false, `${name} must not imply Fabric capacity billing`);
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
  const current = FABRIC_ICON_CATALOG.filter(
    icon => !['concept', 'legacy', 'ui'].includes(icon.kind),
  );
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

test('Fabric 12.0.0 catalog covers every current architecture family and keeps stable destinations', () => {
  const manifest = readJson('Azure_Public_Service_Icons/fabric-manifest.json');
  const lock = readJson('scripts/fabric-icon-sources.json');
  const current = FABRIC_ICON_CATALOG.filter(icon => icon.sourceVersion === '12.0.0');
  const sourceAssets = current.map(icon => icon.sourceAsset);
  const nonNullSourceAssets = sourceAssets.filter((asset): asset is string => asset !== null);
  const packageComparison = manifest.packageComparison;
  assert.equal(manifest.packageVersion, '12.0.0');
  assert.equal(manifest.officialFamilyCount, 88);
  assert.equal(current.length, 88);
  assert.equal(new Set(sourceAssets).size, 88);
  assert.deepEqual(manifest.officialAssets, nonNullSourceAssets.slice().sort((left, right) => left.localeCompare(right)));
  assert.equal(packageComparison.fromVersion, '8.2.0');
  assert.equal(packageComparison.previousFamilyCount, 82);
  assert.equal(packageComparison.currentFamilyCount, 88);
  assert.deepEqual(packageComparison.added.map((entry: { asset: string }) => entry.asset).sort(), [
    'action_48_item.svg',
    'database_hub_48_color.svg',
    'fabric_apps_48_item.svg',
    'feature_store_48_item.svg',
    'managed_identity_48_item.svg',
    'postgre_sql_database_48_item.svg',
    'rayfin_48_color.svg',
  ]);
  assert.equal(packageComparison.renamed.length, 1);
  assert.equal(packageComparison.renamed[0].from, 'user_data_function_48_item.svg');
  assert.equal(packageComparison.renamed[0].to, 'user_data_functions_48_item.svg');
  assert.match(packageComparison.renamed[0].previousSha256, /^[a-f0-9]{64}$/);
  assert.match(packageComparison.renamed[0].sha256, /^[a-f0-9]{64}$/);
  assert.deepEqual(packageComparison.removed.map((entry: { asset: string }) => entry.asset), [
    'function_set_48_item.svg',
  ]);
  assert.equal(packageComparison.added.length + packageComparison.renamed.length, 8);
  assert.equal(packageComparison.unchangedCount + packageComparison.updated.length, 80);
  assert.ok(lock.packages.some((pkg: { version: string }) => pkg.version === '12.0.0'));
  assert.ok(current.some(icon => icon.sourceAsset === 'user_data_functions_48_item.svg'
    && icon.fileName === 'fabric-item-user-data-function'));
  assert.ok(!sourceAssets.includes('function_set_48_item.svg'));
  assert.equal(getFabricIconByFileName('fabric-item-function-set')?.sourceVersion, '8.2.0');
});
