#!/usr/bin/env node
// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

import { createHash } from 'node:crypto';
import { mkdir, readFile, writeFile } from 'node:fs/promises';
import { dirname, isAbsolute, relative, resolve } from 'node:path';
import { fileURLToPath } from 'node:url';
import JSZip from 'jszip';

const SOURCE_PAGE = 'https://learn.microsoft.com/en-us/azure/architecture/icons/';
const PACKAGE_URL_PREFIX = 'https://arch-center.azureedge.net/icons/Azure_Public_Service_Icons_';
const MAX_PACKAGE_BYTES = 100 * 1024 * 1024;

const here = dirname(fileURLToPath(import.meta.url));
const repoRoot = resolve(here, '..');
const iconRoot = resolve(repoRoot, 'Azure_Public_Service_Icons', 'Icons');
const manifestPath = resolve(repoRoot, 'Azure_Public_Service_Icons', 'manifest.json');
const sourceLockPath = resolve(here, 'azure-icon-source.json');

function sha256(value) {
  return createHash('sha256').update(value).digest('hex');
}

function sha256NormalizedSvg(value) {
  return sha256(Buffer.from(value.toString('utf8').replace(/\r\n/g, '\n')));
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

function compareInventories(previous, current, currentVersion) {
  const previousByPath = new Map((previous?.files ?? []).map(file => [file.path, file.sha256]));
  const currentByPath = new Map(current.map(file => [file.path, file.sha256]));
  const unmatchedPrevious = [...previousByPath].filter(([path]) => !currentByPath.has(path));
  const unmatchedCurrent = [...currentByPath].filter(([path]) => !previousByPath.has(path));
  const moved = [];

  for (const [oldPath, oldHash] of unmatchedPrevious) {
    const oldName = oldPath.split('/').at(-1).toLocaleLowerCase();
    const candidates = unmatchedCurrent.filter(([newPath]) =>
      newPath.split('/').at(-1).toLocaleLowerCase() === oldName,
    );
    if (candidates.length === 1) {
      const [newPath, newHash] = candidates[0];
      moved.push({
        from: oldPath, to: newPath, previousSha256: oldHash, sha256: newHash,
        contentUpdated: oldHash !== newHash,
      });
    }
  }

  const movedFrom = new Set(moved.map(entry => entry.from));
  const movedTo = new Set(moved.map(entry => entry.to));
  return {
    fromVersion: previous?.packageVersion ?? null,
    toVersion: currentVersion,
    added: unmatchedCurrent
      .filter(([path]) => !movedTo.has(path))
      .map(([path, sha256]) => ({ path, sha256 })),
    moved: moved.sort((left, right) => left.from.localeCompare(right.from)),
    updated: [...currentByPath]
      .filter(([path, sha256]) => previousByPath.has(path) && previousByPath.get(path) !== sha256)
      .map(([path, sha256]) => ({ path, previousSha256: previousByPath.get(path), sha256 })),
    removed: unmatchedPrevious
      .filter(([path]) => !movedFrom.has(path))
      .map(([path, sha256]) => ({ path, sha256 })),
  };
}

async function readCappedBody(response) {
  const declaredBytes = Number(response.headers.get('content-length'));
  if (Number.isFinite(declaredBytes) && declaredBytes > MAX_PACKAGE_BYTES) {
    throw new Error(`Icon package declares an unsafe size: ${declaredBytes} bytes`);
  }
  const reader = response.body?.getReader();
  if (!reader) throw new Error('Icon package response has no readable body');

  const chunks = [];
  let total = 0;
  for (;;) {
    const { done, value } = await reader.read();
    if (done) break;
    total += value.byteLength;
    if (total > MAX_PACKAGE_BYTES) {
      await reader.cancel();
      throw new Error(`Icon package exceeds the ${MAX_PACKAGE_BYTES} byte safety limit`);
    }
    chunks.push(Buffer.from(value));
  }
  return Buffer.concat(chunks, total);
}

async function main() {
  const checkOnly = process.argv.includes('--check');
  const sourceLock = JSON.parse(await readFile(sourceLockPath, 'utf8'));
  const {
    packageUrl, packageVersion: lockedVersion, packageSha256, officialIconCount,
    comparisonBaselineFile, comparisonBaselineSha256,
  } = sourceLock;
  if (
    sourceLock.sourcePage !== SOURCE_PAGE
    || !packageUrl.startsWith(PACKAGE_URL_PREFIX)
    || packageUrl !== `${PACKAGE_URL_PREFIX}${lockedVersion}.zip`
    || !/^V\d+$/.test(lockedVersion)
    || !/^[a-f0-9]{64}$/.test(packageSha256)
    || !Number.isInteger(officialIconCount)
    || comparisonBaselineFile !== 'azure-icon-v24-baseline.json'
    || !/^[a-f0-9]{64}$/.test(comparisonBaselineSha256)
  ) {
    throw new Error('Azure icon source lock is invalid');
  }
  const baselineBytes = await readFile(resolve(here, comparisonBaselineFile));
  if (sha256(baselineBytes) !== comparisonBaselineSha256) {
    throw new Error('Azure V24 comparison baseline failed its SHA-256 pin');
  }
  const comparisonBaseline = JSON.parse(baselineBytes.toString('utf8'));
  if (comparisonBaseline.packageVersion !== 'V24' || comparisonBaseline.files.length !== 714) {
    throw new Error('Azure V24 comparison baseline does not cover the expected inventory');
  }

  const response = await fetch(packageUrl, { signal: AbortSignal.timeout(60_000) });
  if (!response.ok) {
    throw new Error(`Failed to download ${packageUrl}: HTTP ${response.status}`);
  }

  const archive = await readCappedBody(response);
  if (sha256(archive) !== packageSha256) {
    throw new Error(
      `Azure icon package SHA-256 does not match the reviewed source lock: ${packageUrl}`,
    );
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
  if (entries.length !== officialIconCount) {
    throw new Error(
      `Azure icon inventory differs from the source lock: expected ${officialIconCount}, found ${entries.length}`,
    );
  }

  const files = [];
  const categories = new Map();
  const seenPaths = new Set();
  const missing = [];
  const changedPaths = [];
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
    if (!existing || sha256NormalizedSvg(existing) !== sha256NormalizedSvg(content)) {
      if (checkOnly) {
        (existing ? changedPaths : missing).push(relativePath);
      } else {
        await writeFile(destination, content);
        changed += 1;
      }
    }

    categories.set(category, (categories.get(category) ?? 0) + 1);
    files.push({
      path: relativePath,
      sha256: sha256(content),
      normalizedSha256: sha256NormalizedSvg(content),
    });
  }

  const currentVersion = packageVersion(packageUrl);
  const releaseComparison = compareInventories(comparisonBaseline, files, currentVersion);
  const manifest = {
    schemaVersion: 1,
    sourcePage: SOURCE_PAGE,
    packageUrl,
    packageVersion: currentVersion,
    packageSha256,
    officialIconCount: files.length,
    categories: Object.fromEntries([...categories.entries()].sort(([left], [right]) =>
      left.localeCompare(right),
    )),
    ...(releaseComparison ? { releaseComparison } : {}),
    files,
  };

  if (checkOnly) {
    const currentManifest = JSON.parse(await readFile(manifestPath, 'utf8'));
    if (missing.length > 0 || changedPaths.length > 0) {
      throw new Error(
        `Azure icon files are out of date; missing: ${missing.length}, changed: ${changedPaths.length}`,
      );
    }
    if (JSON.stringify(currentManifest) !== JSON.stringify(manifest)) {
      throw new Error('Azure icon manifest differs from the pinned package inventory');
    }
  } else {
    await mkdir(dirname(manifestPath), { recursive: true });
    await writeFile(manifestPath, `${JSON.stringify(manifest, null, 2)}\n`, 'utf8');
  }

  console.log(
    `[icons:sync] ${manifest.packageVersion}: ${files.length} official icons, `
    + `${checkOnly ? 'check only' : `${changed} files added or updated`}`,
  );
  console.log(`[icons:sync] package SHA-256: ${manifest.packageSha256}`);
}

main().catch(error => {
  console.error(`[icons:sync] ${error.stack || error.message}`);
  process.exitCode = 1;
});
