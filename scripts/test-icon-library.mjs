#!/usr/bin/env node
// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

import { createHash } from 'node:crypto';
import { readFile } from 'node:fs/promises';
import { dirname, isAbsolute, relative, resolve } from 'node:path';
import { fileURLToPath } from 'node:url';

const here = dirname(fileURLToPath(import.meta.url));
const repoRoot = resolve(here, '..');
const iconRoot = resolve(repoRoot, 'Azure_Public_Service_Icons', 'Icons');
const manifestPath = resolve(repoRoot, 'Azure_Public_Service_Icons', 'manifest.json');

function sha256(value) {
  return createHash('sha256').update(value).digest('hex');
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
      if (sha256(content) !== entry.sha256) changed.push(entry.path);
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
}

main().catch(error => {
  console.error(`[test:icons] ${error.stack || error.message}`);
  process.exitCode = 1;
});
