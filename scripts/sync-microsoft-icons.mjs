#!/usr/bin/env node
// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

/**
 * Syncs the official Microsoft Power Platform and Dynamics 365 icon packages
 * into the local icon library.
 *
 * The destination file names, folders, and display metadata are declared in
 * `src/data/microsoftProductIconCatalog.ts`, which is the single source of
 * truth. This script fails when the catalog references an asset that the
 * official package no longer ships, so upstream renames surface immediately.
 */

import { createHash } from 'node:crypto';
import { mkdir, readFile, writeFile } from 'node:fs/promises';
import { dirname, isAbsolute, relative, resolve } from 'node:path';
import { fileURLToPath } from 'node:url';
import JSZip from 'jszip';

/**
 * The official packages, pinned by URL *and* by SHA-256. Microsoft republishes
 * these archives in place, so the digest is what actually pins the content: a
 * substituted archive fails the check instead of being silently self-certified
 * into the manifest. When Microsoft ships a genuine update, verify the new
 * release on the `sourcePage`, then bump both `sha256` here and
 * `MICROSOFT_PRODUCT_ICON_PACKAGE_VERSION` in the catalog.
 */
const PACKAGES = [
  {
    family: 'power-platform',
    fileName: 'Power-Platform-icons-scalable.zip',
    sourcePage: 'https://learn.microsoft.com/en-us/power-platform/guidance/icons',
    url: 'https://download.microsoft.com/download/498606aa-6d27-4f13-aa5c-1401078c153b/Power-Platform-icons-scalable.zip',
    sha256: 'd5abafebbce553690caf7b42dd14b8335bf2c0dfc09f46bd9ac8041187da2c3a',
    released: '2025-12',
  },
  {
    family: 'dynamics-365',
    fileName: 'Dynamics-365-icons-scalable.zip',
    sourcePage: 'https://learn.microsoft.com/en-us/dynamics365/get-started/icons',
    url: 'https://download.microsoft.com/download/498606aa-6d27-4f13-aa5c-1401078c153b/Dynamics-365-icons-scalable.zip',
    sha256: 'f5f4d96aaa637b71e47136cc74e9a4a69b8a81f7e43c1c5756f00b5cea9ffe02',
    released: '2025-12',
  },
  {
    family: 'microsoft-365',
    fileName: 'Microsoft-365-architecture-icons.zip',
    sourcePage:
      'https://learn.microsoft.com/en-us/previous-versions/microsoft-365/solutions/architecture-icons-templates',
    // `go.microsoft.com/fwlink/?linkid=869455` is the download link published on
    // the source page. It redirects, so the digest below is what actually pins
    // the content.
    url: 'https://go.microsoft.com/fwlink/?linkid=869455',
    sha256: '522c4d43cf98a00380b0836f9ec6e06d81fb99758b258c49b6e5f385c05c547a',
    released: '2024-04',
  },
];

const FAMILY_CATEGORY = {
  'power-platform': 'power platform',
  'dynamics-365': 'dynamics 365',
  'microsoft-365': 'microsoft 365',
};

const MAX_PACKAGE_BYTES = 16 * 1024 * 1024;

const here = dirname(fileURLToPath(import.meta.url));
const repoRoot = resolve(here, '..');
const catalogPath = resolve(repoRoot, 'src', 'data', 'microsoftProductIconCatalog.ts');
const iconRoot = resolve(repoRoot, 'Azure_Public_Service_Icons', 'Icons');
const manifestPath = resolve(
  repoRoot,
  'Azure_Public_Service_Icons',
  'microsoft-product-manifest.json',
);

function sha256(value) {
  return createHash('sha256').update(value).digest('hex');
}

function resolveIconDestination(category, fileName) {
  if (!/^[a-z0-9 +-]+$/.test(category)) {
    throw new Error(`Unsafe icon category: ${category}`);
  }
  if (!/^[a-z0-9-]+$/.test(fileName)) {
    throw new Error(`Unsafe icon destination name: ${fileName}`);
  }

  const destination = resolve(iconRoot, category, `${fileName}.svg`);
  const pathFromRoot = relative(iconRoot, destination);
  if (pathFromRoot === '' || pathFromRoot.startsWith('..') || isAbsolute(pathFromRoot)) {
    throw new Error(`Icon path escapes the library root: ${category}/${fileName}.svg`);
  }
  return destination;
}

function parseCatalog(catalogText) {
  const version = /MICROSOFT_PRODUCT_ICON_PACKAGE_VERSION\s*=\s*'([^']+)'/.exec(catalogText)?.[1];
  if (!version) {
    throw new Error('The Microsoft product icon catalog does not declare its package version');
  }

  const entries = [];
  const entryPattern =
    /defineMicrosoftProductIcon\(\s*'([^']+)'\s*,\s*'([^']+)'\s*,\s*'([^']+)'\s*,\s*'([^']+\.svg)'\s*,\s*'([^']+)'\s*,/g;
  let match;
  while ((match = entryPattern.exec(catalogText)) !== null) {
    entries.push({
      serviceName: match[1],
      displayName: match[2],
      fileName: match[3],
      sourceAsset: match[4],
      family: match[5],
    });
  }
  if (entries.length === 0) {
    throw new Error('No Microsoft product icon catalog entries were found');
  }

  const knownFamilies = new Set(PACKAGES.map(entry => entry.family));
  const destinationNames = new Set();
  const sourceAssets = new Set();
  for (const entry of entries) {
    if (!knownFamilies.has(entry.family)) {
      throw new Error(`Unknown icon family in catalog: ${entry.family}`);
    }
    const destinationKey = entry.fileName.toLocaleLowerCase();
    if (destinationNames.has(destinationKey)) {
      throw new Error(`Duplicate icon destination name: ${entry.fileName}`);
    }
    destinationNames.add(destinationKey);

    const sourceKey = `${entry.family}:${entry.sourceAsset.toLocaleLowerCase()}`;
    if (sourceAssets.has(sourceKey)) {
      throw new Error(`Duplicate source asset: ${entry.sourceAsset}`);
    }
    sourceAssets.add(sourceKey);
  }

  return { version, entries };
}

/**
 * Read a response body with a running byte counter so the size cap actually
 * bounds memory. `arrayBuffer()` would buffer the whole response first, making
 * the check a post-hoc report rather than a guard.
 */
async function readCappedBody(response, label) {
  const declared = Number(response.headers.get('content-length'));
  if (Number.isFinite(declared) && declared > MAX_PACKAGE_BYTES) {
    throw new Error(`${label} declares ${declared} bytes, above the ${MAX_PACKAGE_BYTES} byte cap`);
  }

  const reader = response.body?.getReader();
  if (!reader) {
    throw new Error(`${label} returned an unreadable response body`);
  }

  const chunks = [];
  let total = 0;
  for (;;) {
    const { done, value } = await reader.read();
    if (done) break;
    total += value.byteLength;
    if (total > MAX_PACKAGE_BYTES) {
      await reader.cancel();
      throw new Error(`${label} exceeds the ${MAX_PACKAGE_BYTES} byte cap`);
    }
    chunks.push(Buffer.from(value));
  }
  return Buffer.concat(chunks, total);
}

async function downloadPackage(pkg) {
  const response = await fetch(pkg.url);
  if (!response.ok) {
    throw new Error(`Failed to download ${pkg.url}: HTTP ${response.status}`);
  }

  const archive = await readCappedBody(response, pkg.fileName);
  const digest = sha256(archive);
  if (digest !== pkg.sha256) {
    throw new Error(
      `${pkg.fileName} does not match its pinned SHA-256.\n`
      + `  expected: ${pkg.sha256}\n`
      + `  actual:   ${digest}\n`
      + `  Verify the release on ${pkg.sourcePage}, then update the pinned digest `
      + 'and MICROSOFT_PRODUCT_ICON_PACKAGE_VERSION.',
    );
  }

  const zip = await JSZip.loadAsync(archive);
  const assets = new Map();
  for (const entry of Object.values(zip.files)) {
    if (entry.dir || !/\.svg$/i.test(entry.name)) continue;
    const normalized = entry.name.replaceAll('\\', '/');
    if (normalized.split('/').some(segment => segment === '.' || segment === '..')) {
      throw new Error(`Unsafe entry in ${pkg.fileName}: ${entry.name}`);
    }
    assets.set(normalized.toLocaleLowerCase(), entry);
  }
  if (assets.size === 0) {
    throw new Error(`${pkg.fileName} contains no SVG assets`);
  }

  return { archive, assets };
}

async function main() {
  const checkOnly = process.argv.includes('--check');
  const catalogText = await readFile(catalogPath, 'utf8');
  const { version, entries } = parseCatalog(catalogText);

  const packages = new Map();
  for (const pkg of PACKAGES) {
    packages.set(pkg.family, { ...pkg, ...(await downloadPackage(pkg)) });
  }

  const files = [];
  const categories = new Map();
  const drift = [];
  let changed = 0;

  for (const entry of entries) {
    const pkg = packages.get(entry.family);
    const asset = pkg.assets.get(entry.sourceAsset.toLocaleLowerCase());
    if (!asset) {
      drift.push(`${entry.serviceName}: missing ${entry.family} asset "${entry.sourceAsset}"`);
      continue;
    }

    const category = FAMILY_CATEGORY[entry.family];
    if (!category) {
      throw new Error(`No icon category is defined for family ${entry.family}`);
    }
    const content = Buffer.from(await asset.async('uint8array'));
    if (!content.toString('utf8').trimStart().startsWith('<')) {
      throw new Error(`${entry.sourceAsset} is not an SVG document`);
    }

    const destination = resolveIconDestination(category, entry.fileName);
    let existing;
    try {
      existing = await readFile(destination);
    } catch {
      existing = null;
    }

    if (!existing || !existing.equals(content)) {
      if (checkOnly) {
        drift.push(`${entry.serviceName}: ${category}/${entry.fileName}.svg is out of date`);
      } else {
        await mkdir(dirname(destination), { recursive: true });
        await writeFile(destination, content);
        changed += 1;
      }
    }

    categories.set(category, (categories.get(category) ?? 0) + 1);
    files.push({
      path: `${category}/${entry.fileName}.svg`,
      serviceName: entry.serviceName,
      family: entry.family,
      sourceAsset: entry.sourceAsset,
      sha256: sha256(content),
    });
  }

  const unusedAssets = new Map();
  for (const pkg of packages.values()) {
    const used = new Set(
      entries
        .filter(entry => entry.family === pkg.family)
        .map(entry => entry.sourceAsset.toLocaleLowerCase()),
    );
    for (const key of pkg.assets.keys()) {
      if (!used.has(key)) {
        unusedAssets.set(pkg.family, (unusedAssets.get(pkg.family) ?? 0) + 1);
      }
    }
  }

  if (drift.length > 0) {
    throw new Error(`Microsoft product icon drift detected:\n  - ${drift.join('\n  - ')}`);
  }

  const manifest = {
    schemaVersion: 1,
    packageVersion: version,
    packages: PACKAGES.map(pkg => ({
      family: pkg.family,
      sourcePage: pkg.sourcePage,
      released: pkg.released,
      packageUrl: pkg.url,
      packageSha256: sha256(packages.get(pkg.family).archive),
    })),
    terms:
      'Microsoft permits these icons in architectural diagrams, training materials, or '
      + 'documentation. Do not crop, flip, rotate, or distort them, and do not use them to '
      + 'represent a non-Microsoft product.',
    iconCount: files.length,
    categories: Object.fromEntries(
      [...categories.entries()].sort(([left], [right]) => left.localeCompare(right)),
    ),
    files: files.sort((left, right) => left.path.localeCompare(right.path)),
  };

  if (!checkOnly) {
    await mkdir(dirname(manifestPath), { recursive: true });
    await writeFile(manifestPath, `${JSON.stringify(manifest, null, 2)}\n`, 'utf8');
  }

  if (unusedAssets.size > 0) {
    // The Microsoft 365 package ships every symbol in six colour treatments and
    // the catalog keeps one per symbol, so a large unused count is expected
    // there. A non-zero count for the logo packages means a real omission.
    const summary = [...unusedAssets.entries()]
      .sort(([left], [right]) => left.localeCompare(right))
      .map(([family, count]) => `${family}: ${count}`)
      .join(', ');
    console.warn(`[icons:sync:microsoft] package assets not in the catalog (${summary})`);
  }

  console.log(
    `[icons:sync:microsoft] ${manifest.packageVersion}: ${files.length} icons, `
    + `${checkOnly ? 'check only' : `${changed} files added or updated`}`,
  );
  for (const pkg of manifest.packages) {
    console.log(`[icons:sync:microsoft] ${pkg.family} SHA-256: ${pkg.packageSha256}`);
  }
}

main().catch(error => {
  console.error(`[icons:sync:microsoft] ${error.stack || error.message}`);
  process.exitCode = 1;
});
