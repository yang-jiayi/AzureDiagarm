#!/usr/bin/env node
// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

import { createHash } from 'node:crypto';
import { readFile, writeFile } from 'node:fs/promises';
import { dirname, join, resolve } from 'node:path';
import { fileURLToPath, pathToFileURL } from 'node:url';

const here = dirname(fileURLToPath(import.meta.url));
const repoRoot = resolve(here, '..');
const iconRoot = resolve(repoRoot, 'Azure_Public_Service_Icons', 'Icons');
const sourcePath = resolve(here, 'ignite-icon-sources.json');
const manifestPath = resolve(repoRoot, 'Azure_Public_Service_Icons', 'ignite-manifest.json');
const hash = bytes => createHash('sha256').update(bytes).digest('hex');

export async function main() {
  const checkOnly = process.argv.includes('--check');
  const source = JSON.parse(await readFile(sourcePath, 'utf8'));
  const paths = new Set();
  const files = [];

  for (const icon of source.icons) {
    if (
      !/^other\/[a-z0-9-]+\.svg$/.test(icon.path)
      || paths.has(icon.path)
      || !Array.isArray(icon.aliases)
      || !icon.serviceName
      || !icon.displayName
    ) {
      throw new Error(`Invalid or duplicate local Ignite symbol: ${icon.path}`);
    }
    paths.add(icon.path);
    const bytes = await readFile(join(iconRoot, icon.path));
    const svg = bytes.toString('utf8');
    if (
      !/<svg[\s>]/i.test(svg)
      || /<\s*(?:script|foreignObject|image|style|a)[\s/>]/i.test(svg)
      || /\bon[a-z]+\s*=/i.test(svg)
      || /\b(?:xlink:)?href\s*=/i.test(svg)
      || /(?:<!DOCTYPE|<!ENTITY|<\?xml-stylesheet|@import|url\()/i.test(svg)
    ) {
      throw new Error(`Unsafe or invalid local Ignite SVG: ${icon.path}`);
    }
    files.push({ ...icon, kind: 'unofficial-concept-substitute', sha256: hash(bytes) });
  }

  const manifest = {
    schemaVersion: 1,
    sourcePage: source.sourcePage,
    releaseDate: source.releaseDate,
    assetPolicy: source.assetPolicy,
    files,
  };
  const serialized = `${JSON.stringify(manifest, null, 2)}\n`;

  if (checkOnly) {
    const current = await readFile(manifestPath, 'utf8');
    if (current !== serialized) {
      throw new Error('Ignite icon manifest differs from the local source inventory');
    }
  } else {
    await writeFile(manifestPath, serialized);
  }
  console.log(`[icons:sync:ignite] ${files.length} local generic symbols; ${checkOnly ? 'check only' : 'manifest updated'}`);
}

if (process.argv[1] && import.meta.url === pathToFileURL(resolve(process.argv[1])).href) {
  main().catch(error => {
    console.error(`[icons:sync:ignite] ${error.message}`);
    process.exitCode = 1;
  });
}
