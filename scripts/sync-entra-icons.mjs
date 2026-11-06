#!/usr/bin/env node
// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

import { createHash } from 'node:crypto';
import { mkdir, readFile, writeFile } from 'node:fs/promises';
import { dirname, isAbsolute, relative, resolve } from 'node:path';
import { fileURLToPath, pathToFileURL } from 'node:url';
import JSZip from 'jszip';
import { assertSafeSvg } from './sync-fabric-icons.mjs';

const SOURCE_PAGE = 'https://learn.microsoft.com/en-us/entra/architecture/architecture-icons';
const PACKAGE_URL =
  'https://download.microsoft.com/download/3/1/a/31a56038-856a-4489-88e4-ee5a1c4352be/Microsoft%20Entra%20architecture%20icons%20-%20Oct%202023.zip';
const PACKAGE_SHA256 = '4e07536706a2d092e6524e5417e2c861333fdbdb41c36f78d44dfa07ccc5eedc';
const MAX_PACKAGE_BYTES = 32 * 1024 * 1024;
const here = dirname(fileURLToPath(import.meta.url));
const repoRoot = resolve(here, '..');
const iconRoot = resolve(repoRoot, 'Azure_Public_Service_Icons', 'Icons');
const lockPath = resolve(here, 'entra-icon-sources.json');
const manifestPath = resolve(repoRoot, 'Azure_Public_Service_Icons', 'entra-manifest.json');
const sha256 = value => createHash('sha256').update(value).digest('hex');

function resolveIconPath(relativePath) {
  const segments = relativePath.replaceAll('\\', '/').split('/');
  if (
    segments.length < 2
    || segments.some(segment => segment === '' || segment === '.' || segment === '..')
    || !/\.svg$/i.test(segments.at(-1))
  ) {
    throw new Error(`Unsafe Entra icon path: ${relativePath}`);
  }
  const destination = resolve(iconRoot, ...segments);
  const pathFromRoot = relative(iconRoot, destination);
  if (pathFromRoot === '' || pathFromRoot.startsWith('..') || isAbsolute(pathFromRoot)) {
    throw new Error(`Entra icon path escapes the library: ${relativePath}`);
  }
  return destination;
}

async function download() {
  const response = await fetch(PACKAGE_URL, { signal: AbortSignal.timeout(60_000) });
  if (!response.ok) throw new Error(`Failed to download the official Entra package: HTTP ${response.status}`);
  const declaredBytes = Number(response.headers.get('content-length'));
  if (Number.isFinite(declaredBytes) && declaredBytes > MAX_PACKAGE_BYTES) {
    throw new Error(`Entra package declares an unsafe size: ${declaredBytes} bytes`);
  }
  const reader = response.body?.getReader();
  if (!reader) throw new Error('The official Entra package has no readable response body');
  const chunks = [];
  let total = 0;
  for (;;) {
    const { done, value } = await reader.read();
    if (done) break;
    total += value.byteLength;
    if (total > MAX_PACKAGE_BYTES) {
      await reader.cancel();
      throw new Error(`Entra package exceeds the ${MAX_PACKAGE_BYTES} byte safety limit`);
    }
    chunks.push(Buffer.from(value));
  }
  const archive = Buffer.concat(chunks, total);
  if (sha256(archive) !== PACKAGE_SHA256) throw new Error('Entra package failed its reviewed SHA-256 pin');
  return archive;
}

export async function main() {
  const checkOnly = process.argv.includes('--check');
  const lock = JSON.parse(await readFile(lockPath, 'utf8'));
  if (
    lock.sourcePage !== SOURCE_PAGE || lock.packageUrl !== PACKAGE_URL
    || lock.packageSha256 !== PACKAGE_SHA256 || lock.svgCount !== lock.icons.length
  ) {
    throw new Error('Entra icon source lock does not match the reviewed official source');
  }

  const zip = await JSZip.loadAsync(await download());
  const entries = Object.values(zip.files).filter(entry => !entry.dir && /\.svg$/i.test(entry.name));
  if (entries.length !== lock.svgCount) {
    throw new Error(`Expected ${lock.svgCount} Entra SVGs, found ${entries.length}`);
  }
  const sourceByPath = new Map(entries.map(entry => [entry.name, entry]));
  const sourcePaths = new Set(lock.icons.map(icon => icon.sourcePath));
  if (sourcePaths.size !== lock.svgCount || entries.some(entry => !sourcePaths.has(entry.name))) {
    throw new Error('Entra source map does not cover the complete official SVG inventory');
  }

  const files = [];
  const drift = [];
  for (const mapping of lock.icons) {
    const source = sourceByPath.get(mapping.sourcePath);
    if (!source || !['add', 'duplicate', 'variant-not-added'].includes(mapping.action)) {
      throw new Error(`Invalid Entra source mapping: ${mapping.sourcePath}`);
    }
    const content = Buffer.from(await source.async('uint8array'));
    assertSafeSvg(content, mapping.sourcePath);
    const record = {
      sourcePath: mapping.sourcePath,
      sha256: sha256(content),
      action: mapping.action,
      ...(mapping.targetPath ? { targetPath: mapping.targetPath } : {}),
      ...(mapping.counterpart ? { counterpart: mapping.counterpart } : {}),
    };

    if (mapping.action === 'variant-not-added') {
      if (!sourceByPath.has(mapping.counterpart)) {
        throw new Error(`Entra icon variant has no matching official asset: ${mapping.sourcePath}`);
      }
    } else {
      const destination = resolveIconPath(mapping.targetPath);
      let existing;
      try {
        existing = await readFile(destination);
      } catch (error) {
        if (error.code !== 'ENOENT') throw error;
      }
      if (mapping.action === 'duplicate' && !existing) {
        throw new Error(`Expected existing equivalent Entra service icon: ${mapping.targetPath}`);
      }
      if (mapping.action === 'add' && !existing?.equals(content)) {
        if (checkOnly) {
          drift.push(mapping.targetPath);
        } else {
          await mkdir(dirname(destination), { recursive: true });
          await writeFile(destination, content);
        }
      }
      if (mapping.action === 'duplicate') {
        record.targetSha256 = sha256(existing);
      }
    }
    files.push(record);
  }

  const manifest = {
    schemaVersion: 1,
    sourcePage: SOURCE_PAGE,
    packageUrl: PACKAGE_URL,
    packageVersion: lock.packageVersion,
    packageSha256: PACKAGE_SHA256,
    terms:
      'Microsoft permits these icons in architectural diagrams, training materials, or '
      + 'documentation; marketing communications are not permitted. Do not crop, flip, '
      + 'rotate, distort, or change the icon shape.',
    svgCount: files.length,
    addedIconCount: files.filter(file => file.action === 'add').length,
    duplicateIconCount: files.filter(file => file.action === 'duplicate').length,
    variantNotAddedCount: files.filter(file => file.action === 'variant-not-added').length,
    files,
  };

  if (checkOnly) {
    const existing = JSON.parse(await readFile(manifestPath, 'utf8'));
    if (drift.length > 0) throw new Error(`Entra icon files are out of date: ${drift.join(', ')}`);
    if (JSON.stringify(existing) !== JSON.stringify(manifest)) {
      throw new Error('Entra icon manifest differs from the pinned official package');
    }
  } else {
    await writeFile(manifestPath, `${JSON.stringify(manifest, null, 2)}\n`);
  }
  console.log(
    `[icons:sync:entra] ${files.length} official SVGs compared; `
    + `${manifest.addedIconCount} added, ${manifest.duplicateIconCount} mapped to existing services, `
    + `${manifest.variantNotAddedCount} monochrome variants not duplicated`,
  );
}

if (process.argv[1] && import.meta.url === pathToFileURL(resolve(process.argv[1])).href) {
  main().catch(error => {
    console.error(`[icons:sync:entra] ${error.stack || error.message}`);
    process.exitCode = 1;
  });
}
