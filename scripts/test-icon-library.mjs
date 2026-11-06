#!/usr/bin/env node
// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

import { createHash } from 'node:crypto';
import { readFile } from 'node:fs/promises';
import { dirname, isAbsolute, relative, resolve } from 'node:path';
import { fileURLToPath } from 'node:url';
import { assertSafeSvg, parseCatalog } from './sync-fabric-icons.mjs';

const here = dirname(fileURLToPath(import.meta.url));
const repoRoot = resolve(here, '..');
const iconRoot = resolve(repoRoot, 'Azure_Public_Service_Icons', 'Icons');
const manifestPath = resolve(repoRoot, 'Azure_Public_Service_Icons', 'manifest.json');
const fabricManifestPath = resolve(
  repoRoot,
  'Azure_Public_Service_Icons',
  'fabric-manifest.json',
);
const fabricCatalogPath = resolve(repoRoot, 'src', 'data', 'fabricIconCatalog.ts');
const microsoftManifestPath = resolve(
  repoRoot,
  'Azure_Public_Service_Icons',
  'microsoft-product-manifest.json',
);
const microsoftCatalogPath = resolve(repoRoot, 'src', 'data', 'microsoftProductIconCatalog.ts');

// Must stay in step with `FAMILY_CATEGORY` in the catalog and in
// `scripts/sync-microsoft-icons.mjs`.
const MICROSOFT_PRODUCT_FAMILY_CATEGORY = {
  'power-platform': 'power platform',
  'dynamics-365': 'dynamics 365',
  'microsoft-365': 'microsoft 365',
};

function sha256(value) {
  return createHash('sha256').update(value).digest('hex');
}

function sha256NormalizedSvg(value) {
  // Git normalizes tracked SVG text to LF, while an older Windows worktree may
  // still contain CRLF bytes. Line endings do not alter the SVG, so hash the
  // canonical LF representation to keep integrity checks cross-platform.
  return sha256(Buffer.from(value.toString('utf8').replace(/\r\n/g, '\n')));
}

function resolveManifestIconPath(relativePath) {
  if (typeof relativePath !== 'string') {
    throw new Error('Icon manifest contains a non-string path');
  }
  const segments = relativePath.replaceAll('\\', '/').split('/');
  if (
    segments.length < 2
    || segments.some(segment => segment === '' || segment === '.' || segment === '..')
    || !/\.svg$/i.test(segments.at(-1))
  ) {
    throw new Error(`Icon manifest contains an unsafe path: ${relativePath}`);
  }

  const path = resolve(iconRoot, ...segments);
  const pathFromRoot = relative(iconRoot, path);
  if (pathFromRoot === '' || pathFromRoot.startsWith('..') || isAbsolute(pathFromRoot)) {
    throw new Error(`Icon manifest path escapes the library root: ${relativePath}`);
  }
  return path;
}

async function main() {
  const manifest = JSON.parse(await readFile(manifestPath, 'utf8'));
  if (!Array.isArray(manifest.files) || manifest.files.length !== manifest.officialIconCount) {
    throw new Error('Icon manifest count does not match its file inventory');
  }

  const missing = [];
  const changed = [];
  const seenPaths = new Set();
  for (const entry of manifest.files) {
    const normalizedPath = typeof entry.path === 'string'
      ? entry.path.toLocaleLowerCase()
      : '';
    if (seenPaths.has(normalizedPath)) {
      throw new Error(`Icon manifest contains a duplicate path: ${entry.path}`);
    }
    seenPaths.add(normalizedPath);
    if (!/^[a-f0-9]{64}$/i.test(entry.sha256)) {
      throw new Error(`Icon manifest contains an invalid SHA-256: ${entry.path}`);
    }

    const path = resolveManifestIconPath(entry.path);
    try {
      const content = await readFile(path);
      const actualHash = sha256NormalizedSvg(content);
      if (actualHash !== entry.normalizedSha256) changed.push(entry.path);
    } catch {
      missing.push(entry.path);
    }
  }

  if (missing.length > 0 || changed.length > 0) {
    const details = [
      missing.length > 0 ? `missing: ${missing.join(', ')}` : '',
      changed.length > 0 ? `changed: ${changed.join(', ')}` : '',
    ].filter(Boolean).join('\n');
    throw new Error(`Official Azure icon library is incomplete or modified\n${details}`);
  }

  console.log(
    `[test:icons] ${manifest.packageVersion}: all ${manifest.officialIconCount} official icons verified`,
  );

  const fabricManifest = JSON.parse(await readFile(fabricManifestPath, 'utf8'));
  if (
    !Array.isArray(fabricManifest.files)
    || fabricManifest.files.length !== fabricManifest.paletteIconCount
    || fabricManifest.paletteIconCount
      !== fabricManifest.officialFamilyCount + fabricManifest.localIconCount
        + fabricManifest.legacyIconCount + fabricManifest.uiIconCount
  ) {
    throw new Error('Fabric icon manifest count does not match its file inventory');
  }

  const fabricCatalog = await readFile(fabricCatalogPath, 'utf8');
  const catalogVersion =
    /FABRIC_ICON_PACKAGE_VERSION\s*=\s*'([^']+)'/.exec(fabricCatalog)?.[1];
  if (catalogVersion !== fabricManifest.packageVersion) {
    throw new Error('Fabric catalog and manifest package versions do not match');
  }
  const { entries: catalogEntries } = parseCatalog(fabricCatalog);
  if (catalogEntries.length !== fabricManifest.paletteIconCount) {
    throw new Error('Fabric catalog count does not match the manifest');
  }

  const catalogByPath = new Map();
  const catalogFileNames = new Set();
  const catalogServiceNames = new Set();
  for (const entry of catalogEntries) {
    const normalizedFileName = entry.fileName.toLocaleLowerCase();
    const normalizedServiceName = entry.serviceName.toLocaleLowerCase();
    if (catalogFileNames.has(normalizedFileName)) {
      throw new Error(`Fabric catalog contains a duplicate filename: ${entry.fileName}`);
    }
    if (catalogServiceNames.has(normalizedServiceName)) {
      throw new Error(`Fabric catalog contains a duplicate service: ${entry.serviceName}`);
    }
    catalogFileNames.add(normalizedFileName);
    catalogServiceNames.add(normalizedServiceName);
    catalogByPath.set(`fabric/${entry.fileName}.svg`, entry);
  }

  const fabricLock = JSON.parse(await readFile(resolve(here, 'fabric-icon-sources.json'), 'utf8'));
  if (JSON.stringify(fabricManifest.packages) !== JSON.stringify(fabricLock.packages)) {
    throw new Error('Fabric package provenance differs from its source lock');
  }
  const articleByPath = new Map(fabricLock.articleIcons.map(icon => [`fabric/${icon.fileName}.svg`, icon]));
  const fabricMissing = [];
  const fabricChanged = [];
  const manifestPaths = new Set();
  let articleCount = 0;
  for (const entry of fabricManifest.files) {
    if (manifestPaths.has(entry.path)) {
      throw new Error(`Fabric manifest contains a duplicate path: ${entry.path}`);
    }
    manifestPaths.add(entry.path);
    const catalogEntry = catalogByPath.get(entry.path);
    const sourceType = catalogEntry?.sourceVersion === 'article-2026' ? 'article'
      : catalogEntry?.sourceAsset ? 'npm' : 'local';
    if (
      !catalogEntry || entry.sourceAsset !== catalogEntry.sourceAsset
      || entry.sourceVersion !== catalogEntry.sourceVersion
      || entry.serviceName !== catalogEntry.serviceName
      || entry.sourceType !== sourceType
      || (sourceType === 'npm' && !fabricLock.packages.some(pin => pin.version === entry.sourceVersion))
    ) {
      throw new Error(`Fabric manifest mapping is invalid: ${entry.path}`);
    }
    const article = articleByPath.get(entry.path);
    if (article) {
      articleCount++;
      if (sourceType === 'article' && (
        entry.sha256 !== article.sha256 || entry.articleUrl !== article.url
        || entry.articleTitle !== article.title || entry.articleSection !== article.section
        || entry.sourceUrl !== article.url
      )) {
        throw new Error(`Fabric article provenance is invalid: ${entry.path}`);
      }
    }
    if (!/^[a-f0-9]{64}$/i.test(entry.sha256)) {
      throw new Error(`Fabric manifest contains an invalid SHA-256: ${entry.path}`);
    }

    const path = resolveManifestIconPath(entry.path);
    let content;
    try {
      content = await readFile(path);
    } catch (error) {
      if (error.code !== 'ENOENT') throw error;
      fabricMissing.push(entry.path);
      continue;
    }
    assertSafeSvg(content, entry.path);
    if (sha256(content) !== entry.sha256) fabricChanged.push(entry.path);
  }
  if (articleCount !== fabricLock.articleIcons.length || articleCount !== fabricManifest.articleIconCount) {
    throw new Error('Fabric manifest does not cover every article icon');
  }
  if (fabricMissing.length > 0 || fabricChanged.length > 0) {
    const details = [
      fabricMissing.length > 0 ? `missing: ${fabricMissing.join(', ')}` : '',
      fabricChanged.length > 0 ? `changed: ${fabricChanged.join(', ')}` : '',
    ].filter(Boolean).join('\n');
    throw new Error(`Official Fabric icon library is incomplete or modified\n${details}`);
  }

  console.log(
    `[test:icons] Fabric ${fabricManifest.packageVersion}: all `
    + `${fabricManifest.paletteIconCount} palette icons verified `
    + `(${fabricManifest.officialFamilyCount} current, ${fabricManifest.legacyIconCount} legacy, `
    + `${fabricManifest.uiIconCount} UI, ${fabricManifest.localIconCount} local; ${articleCount} article SVGs)`,
  );

  await verifyMicrosoftProductIcons();
}

async function verifyMicrosoftProductIcons() {
  const manifest = JSON.parse(await readFile(microsoftManifestPath, 'utf8'));
  if (!Array.isArray(manifest.files) || manifest.files.length !== manifest.iconCount) {
    throw new Error('Microsoft product icon manifest count does not match its file inventory');
  }

  const catalogText = await readFile(microsoftCatalogPath, 'utf8');
  const catalogVersion =
    /MICROSOFT_PRODUCT_ICON_PACKAGE_VERSION\s*=\s*'([^']+)'/.exec(catalogText)?.[1];
  if (catalogVersion !== manifest.packageVersion) {
    throw new Error('Microsoft product catalog and manifest package versions do not match');
  }

  const catalogEntries = [];
  const catalogPattern =
    /defineMicrosoftProductIcon\(\s*'([^']+)'\s*,\s*'([^']+)'\s*,\s*'([^']+)'\s*,\s*'([^']+\.svg)'\s*,\s*'([^']+)'\s*,/g;
  let catalogMatch;
  while ((catalogMatch = catalogPattern.exec(catalogText)) !== null) {
    catalogEntries.push({
      serviceName: catalogMatch[1],
      displayName: catalogMatch[2],
      fileName: catalogMatch[3],
      sourceAsset: catalogMatch[4],
      family: catalogMatch[5],
    });
  }
  if (catalogEntries.length !== manifest.iconCount) {
    throw new Error('Microsoft product catalog count does not match the manifest');
  }

  const catalogByPath = new Map();
  const serviceNames = new Set();
  for (const entry of catalogEntries) {
    const category = MICROSOFT_PRODUCT_FAMILY_CATEGORY[entry.family];
    if (!category) {
      throw new Error(`Microsoft product catalog has an unknown family: ${entry.family}`);
    }
    const normalizedServiceName = entry.serviceName.toLocaleLowerCase();
    if (serviceNames.has(normalizedServiceName)) {
      throw new Error(`Microsoft product catalog has a duplicate service: ${entry.serviceName}`);
    }
    serviceNames.add(normalizedServiceName);

    const path = `${category}/${entry.fileName}.svg`;
    if (catalogByPath.has(path)) {
      throw new Error(`Microsoft product catalog has a duplicate icon path: ${path}`);
    }
    catalogByPath.set(path, entry);
  }

  const missing = [];
  const changed = [];
  const seenPaths = new Set();
  for (const entry of manifest.files) {
    if (seenPaths.has(entry.path)) {
      throw new Error(`Microsoft product manifest has a duplicate path: ${entry.path}`);
    }
    seenPaths.add(entry.path);

    const catalogEntry = catalogByPath.get(entry.path);
    if (!catalogEntry) {
      throw new Error(`Microsoft product manifest path is not in the catalog: ${entry.path}`);
    }
    if (
      catalogEntry.serviceName !== entry.serviceName
      || catalogEntry.sourceAsset !== entry.sourceAsset
      || catalogEntry.family !== entry.family
    ) {
      throw new Error(`Microsoft product manifest mapping is invalid: ${entry.path}`);
    }
    if (!/^[a-f0-9]{64}$/i.test(entry.sha256)) {
      throw new Error(`Microsoft product manifest has an invalid SHA-256: ${entry.path}`);
    }

    const path = resolveManifestIconPath(entry.path);
    try {
      const content = await readFile(path);
      if (sha256NormalizedSvg(content) !== entry.sha256) changed.push(entry.path);
    } catch {
      missing.push(entry.path);
    }
  }

  if (seenPaths.size !== catalogByPath.size) {
    throw new Error('Microsoft product catalog has entries that the manifest does not cover');
  }
  if (missing.length > 0 || changed.length > 0) {
    const details = [
      missing.length > 0 ? `missing: ${missing.join(', ')}` : '',
      changed.length > 0 ? `changed: ${changed.join(', ')}` : '',
    ].filter(Boolean).join('\n');
    throw new Error(`Microsoft product icon library is incomplete or modified\n${details}`);
  }

  const counts = manifest.categories ?? {};
  console.log(
    `[test:icons] Microsoft product icons ${manifest.packageVersion}: all ${manifest.iconCount} `
    + `verified (${counts['power platform'] ?? 0} Power Platform, `
    + `${counts['dynamics 365'] ?? 0} Dynamics 365, `
    + `${counts['microsoft 365'] ?? 0} Microsoft 365)`,
  );
}

main().catch(error => {
  console.error(`[test:icons] ${error.stack || error.message}`);
  process.exitCode = 1;
});
