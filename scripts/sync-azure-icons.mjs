#!/usr/bin/env node
// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

import { createHash } from 'node:crypto';
import { mkdir, readFile, writeFile } from 'node:fs/promises';
import { dirname, isAbsolute, relative, resolve } from 'node:path';
import { fileURLToPath } from 'node:url';
import JSZip from 'jszip';

const SOURCE_PAGE = 'https://learn.microsoft.com/en-us/azure/architecture/icons/';
const FALLBACK_PACKAGE_URL =
  'https://arch-center.azureedge.net/icons/Azure_Public_Service_Icons_V24.zip';

const here = dirname(fileURLToPath(import.meta.url));
const repoRoot = resolve(here, '..');
const iconRoot = resolve(repoRoot, 'Azure_Public_Service_Icons', 'Icons');
const manifestPath = resolve(repoRoot, 'Azure_Public_Service_Icons', 'manifest.json');

function sha256(value) {
  return createHash('sha256').update(value).digest('hex');
}

function packageVersion(packageUrl) {
  return /Azure_Public_Service_Icons_(V\d+)\.zip/i.exec(packageUrl)?.[1] ?? 'unknown';
}

function resolveIconDestination(relativePath) {
  const segments = relativePath.replaceAll('\\', '/').split('/');
  if (
    segments.length < 2
    || segments.some(segment => segment === '' || segment === '.' || segment === '..')
    || !/\.svg$/i.test(segments.at(-1))
  ) {
    throw new Error(`Unsafe icon path in package: ${relativePath}`);
  }

  const destination = resolve(iconRoot, ...segments);
  const pathFromRoot = relative(iconRoot, destination);
  if (pathFromRoot === '' || pathFromRoot.startsWith('..') || isAbsolute(pathFromRoot)) {
    throw new Error(`Icon path escapes the library root: ${relativePath}`);
  }
  return destination;
}

async function resolvePackageUrl() {
  if (process.env.AZURE_ICON_PACKAGE_URL) return process.env.AZURE_ICON_PACKAGE_URL;

  try {
    const response = await fetch(SOURCE_PAGE);
    if (!response.ok) throw new Error(`HTTP ${response.status}`);
    const html = await response.text();
    const urls = [...html.matchAll(
      /https:\/\/arch-center\.azureedge\.net\/icons\/Azure_Public_Service_Icons_V\d+\.zip/gi,
    )].map(match => match[0]);
    if (urls.length === 0) throw new Error('No Azure icon package link was found');
    return urls.sort((left, right) => {
      const leftVersion = Number(/V(\d+)/i.exec(left)?.[1] ?? 0);
      const rightVersion = Number(/V(\d+)/i.exec(right)?.[1] ?? 0);
      return rightVersion - leftVersion;
    })[0];
  } catch (error) {
    console.warn(`[icons:sync] using fallback package: ${error.message}`);
    return FALLBACK_PACKAGE_URL;
  }
}

async function main() {
  const packageUrl = await resolvePackageUrl();
  const response = await fetch(packageUrl);
  if (!response.ok) {
    throw new Error(`Failed to download ${packageUrl}: HTTP ${response.status}`);
  }

  const archive = Buffer.from(await response.arrayBuffer());
  if (archive.length > 100 * 1024 * 1024) {
    throw new Error(`Icon package is unexpectedly large: ${archive.length} bytes`);
  }
  const zip = await JSZip.loadAsync(archive);
  const firstIconPath = Object.keys(zip.files).find(path =>
    path.toLowerCase().includes('/icons/') && /\.svg$/i.test(path),
  );
  const iconsMarkerIndex = firstIconPath?.toLowerCase().indexOf('/icons/') ?? -1;
  const prefix = firstIconPath && iconsMarkerIndex >= 0
    ? firstIconPath.slice(0, iconsMarkerIndex + '/icons/'.length)
    : undefined;
  if (!prefix) throw new Error('The package does not contain the expected Icons directory');

  const entries = Object.values(zip.files)
    .filter(entry => !entry.dir && entry.name.startsWith(prefix) && /\.svg$/i.test(entry.name))
    .sort((left, right) => left.name.localeCompare(right.name));
  if (entries.length < 500) {
    throw new Error(`The package contains an unexpectedly small icon inventory: ${entries.length}`);
  }

  const files = [];
  const categories = new Map();
  const seenPaths = new Set();
  let changed = 0;

  for (const entry of entries) {
    const relativePath = entry.name.slice(prefix.length).replaceAll('\\', '/');
    const normalizedPath = relativePath.toLocaleLowerCase();
    if (seenPaths.has(normalizedPath)) {
      throw new Error(`The package contains a duplicate icon path: ${relativePath}`);
    }
    seenPaths.add(normalizedPath);

    const category = relativePath.split('/')[0];
    const content = Buffer.from(await entry.async('uint8array'));
    const destination = resolveIconDestination(relativePath);

    await mkdir(dirname(destination), { recursive: true });
    let existing;
    try {
      existing = await readFile(destination);
    } catch {
      existing = null;
    }
    if (!existing || !existing.equals(content)) {
      await writeFile(destination, content);
      changed += 1;
    }

    categories.set(category, (categories.get(category) ?? 0) + 1);
    files.push({
      path: relativePath,
      sha256: sha256(content),
    });
  }

  const manifest = {
    schemaVersion: 1,
    sourcePage: SOURCE_PAGE,
    packageUrl,
    packageVersion: packageVersion(packageUrl),
    packageSha256: sha256(archive),
    officialIconCount: files.length,
    categories: Object.fromEntries([...categories.entries()].sort(([left], [right]) =>
      left.localeCompare(right),
    )),
    files,
  };

  await mkdir(dirname(manifestPath), { recursive: true });
  await writeFile(manifestPath, `${JSON.stringify(manifest, null, 2)}\n`, 'utf8');

  console.log(
    `[icons:sync] ${manifest.packageVersion}: ${files.length} official icons, `
    + `${changed} files added or updated`,
  );
  console.log(`[icons:sync] package SHA-256: ${manifest.packageSha256}`);
}

main().catch(error => {
  console.error(`[icons:sync] ${error.stack || error.message}`);
  process.exitCode = 1;
});
