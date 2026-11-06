// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

import assert from 'node:assert/strict';
import { createHash } from 'node:crypto';
import { readFileSync } from 'node:fs';
import path from 'node:path';
import test from 'node:test';

const root = path.resolve(import.meta.dirname, '..');
const readJson = (relativePath: string) => JSON.parse(readFileSync(path.join(root, relativePath), 'utf8'));
const hash = (bytes: Buffer) => createHash('sha256').update(bytes).digest('hex');
const iconPath = (relativePath: string) => path.join(root, 'Azure_Public_Service_Icons', 'Icons', relativePath);

test('Azure V25 inventory and every distributed SVG match the pinned official package', () => {
  const manifest = readJson('Azure_Public_Service_Icons/manifest.json');
  const sourceLock = readJson('scripts/azure-icon-source.json');
  const baseline = readJson('scripts/azure-icon-v24-baseline.json');
  assert.equal(manifest.packageVersion, 'V25');
  assert.equal(manifest.packageSha256, sourceLock.packageSha256);
  assert.equal(manifest.officialIconCount, 721);
  assert.equal(manifest.files.length, 721);
  assert.equal(baseline.packageVersion, 'V24');
  assert.equal(baseline.files.length, 714);
  assert.equal(new Set(manifest.files.map((file: { path: string }) => file.path.toLocaleLowerCase())).size, 721);

  for (const file of manifest.files) {
    const bytes = readFileSync(iconPath(file.path));
    assert.match(file.sha256, /^[a-f0-9]{64}$/);
    assert.equal(hash(Buffer.from(bytes.toString('utf8').replace(/\r\n/g, '\n'))), file.normalizedSha256, file.path);
  }

  const comparison = manifest.releaseComparison;
  assert.equal(comparison.fromVersion, 'V24');
  assert.equal(comparison.toVersion, 'V25');
  assert.deepEqual(comparison.added.map((file: { path: string }) => file.path).sort(), [
    'new icons/036687350-icon-service-Azure-Confidential-Clean-Rooms.svg',
    'new icons/038647133-icon-service-ACA-Sandboxes.svg',
    'new icons/038749693-icon-service-Azure-HorizonDB.svg',
    'new icons/038806325-icon-service-Azure-Encrypted-Transport.svg',
    'new icons/039209793-icon-service-AI-Manager.svg',
    'new icons/039717636-icon-service-App-Service-Builder.svg',
    'new icons/039729786-icon-service-Azure-Device-Registry.svg',
  ]);
  assert.equal(comparison.moved.length, 1);
  assert.deepEqual(
    {
      from: comparison.moved[0].from,
      to: comparison.moved[0].to,
      contentUpdated: comparison.moved[0].contentUpdated,
    },
    {
      from: 'other/032869111-icon-service-Microsoft-Discovery.svg',
      to: 'new icons/032869111-icon-service-Microsoft-Discovery.svg',
      contentUpdated: true,
    },
  );
  assert.equal(comparison.moved[0].previousSha256, baseline.files.find(
    (file: { path: string }) => file.path === comparison.moved[0].from,
  ).sha256);
  assert.equal(comparison.moved[0].sha256, manifest.files.find(
    (file: { path: string }) => file.path === comparison.moved[0].to,
  ).sha256);
  assert.equal(comparison.updated.length, 1);
  assert.equal(comparison.updated[0].path, 'ai + machine learning/00800-icon-service-Translator-Text.svg');
  assert.equal(comparison.updated[0].previousSha256, baseline.files.find(
    (file: { path: string }) => file.path === comparison.updated[0].path,
  ).sha256);
  assert.equal(comparison.removed.length, 0);

  const oldCompatibilityPath = iconPath(comparison.moved[0].from);
  assert.equal(hash(readFileSync(oldCompatibilityPath)), comparison.moved[0].previousSha256);
});

test('the October 2023 Entra package is fully mapped and added assets are hash-verified', () => {
  const lock = readJson('scripts/entra-icon-sources.json');
  const manifest = readJson('Azure_Public_Service_Icons/entra-manifest.json');
  assert.equal(lock.packageVersion, 'October 2023');
  assert.equal(manifest.packageSha256, lock.packageSha256);
  assert.equal(lock.svgCount, 14);
  assert.equal(manifest.svgCount, 14);
  assert.equal(manifest.addedIconCount, 2);
  assert.equal(manifest.duplicateIconCount, 9);
  assert.equal(manifest.variantNotAddedCount, 3);
  assert.deepEqual(manifest.files.map((file: { sourcePath: string }) => file.sourcePath), lock.icons.map(
    (file: { sourcePath: string }) => file.sourcePath,
  ));

  for (const file of manifest.files) {
    assert.match(file.sha256, /^[a-f0-9]{64}$/);
    if (file.action === 'add') {
      assert.equal(hash(readFileSync(iconPath(file.targetPath))), file.sha256, file.targetPath);
    } else if (file.action === 'duplicate') {
      assert.match(file.targetSha256, /^[a-f0-9]{64}$/);
      assert.equal(hash(readFileSync(iconPath(file.targetPath))), file.targetSha256, file.targetPath);
    } else {
      assert.equal(file.action, 'variant-not-added');
      assert.ok(lock.icons.some((source: { sourcePath: string }) => source.sourcePath === file.counterpart));
    }
  }
});

test('all Microsoft product manifest files exist and match their hashes', () => {
  const manifest = readJson('Azure_Public_Service_Icons/microsoft-product-manifest.json');
  assert.equal(manifest.iconCount, 229);
  assert.equal(manifest.categories['microsoft 365'], 205);
  assert.equal(manifest.packages.find((pkg: { family: string }) => pkg.family === 'microsoft-365').released, '2024-04');
  const m365 = manifest.files.filter((file: { family: string }) => file.family === 'microsoft-365');
  assert.equal(m365.length, 205);
  assert.ok(m365.every((file: { kind: string }) => file.kind === 'symbol'));
  assert.equal(m365.filter((file: { includeInServiceMap: boolean }) => file.includeInServiceMap).length, 44);
  for (const file of manifest.files) {
    assert.equal(hash(readFileSync(iconPath(file.path))), file.sha256, file.path);
  }
});
