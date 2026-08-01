#!/usr/bin/env node
// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

import { execFile } from 'node:child_process';
import { createHash } from 'node:crypto';
import {
  mkdir,
  mkdtemp,
  readFile,
  readdir,
  rm,
  writeFile,
} from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { basename, dirname, isAbsolute, join, relative, resolve } from 'node:path';
import { promisify } from 'node:util';
import { fileURLToPath } from 'node:url';

const execFileAsync = promisify(execFile);
const PACKAGE_NAME = '@fabric-msft/svg-icons';
const REGISTRY_URL = 'https://registry.npmjs.org/@fabric-msft%2Fsvg-icons';
const SOURCE_PAGE = 'https://learn.microsoft.com/en-us/fabric/fundamentals/icons';
const OFFICIAL_ZIP_URL =
  'https://github.com/microsoft/fabric-samples/blob/main/docs-samples/Icons.zip';

const here = dirname(fileURLToPath(import.meta.url));
const repoRoot = resolve(here, '..');
const catalogPath = resolve(repoRoot, 'src', 'data', 'fabricIconCatalog.ts');
const iconRoot = resolve(repoRoot, 'Azure_Public_Service_Icons', 'Icons', 'fabric');
const manifestPath = resolve(repoRoot, 'Azure_Public_Service_Icons', 'fabric-manifest.json');

function sha256(value) {
  return createHash('sha256').update(value).digest('hex');
}

function isArchitectureAsset(fileName) {
  return (
    /_32_(?:color|item)\.svg$/i.test(fileName)
    || /_32_non-item\.svg$/i.test(fileName)
    || /_32\.svg$/i.test(fileName)
  );
}

function classifyAsset(fileName) {
  if (/_32_color\.svg$/i.test(fileName)) return 'workload';
  if (/_32_item\.svg$/i.test(fileName)) return 'item';
  if (/_32_non-item\.svg$/i.test(fileName)) return 'navigation';
  return 'item-special';
}

function parseCatalog(catalogText) {
  const version = /FABRIC_ICON_PACKAGE_VERSION\s*=\s*'([^']+)'/.exec(catalogText)?.[1];
  if (!version) throw new Error('Fabric catalog does not declare its package version');

  const entries = [];
  const entryPattern =
    /defineFabricIcon\(\s*'([^']+)'\s*,\s*'([^']+)'\s*,\s*'([^']+)'\s*,\s*(?:'([^']+\.svg)'|null)\s*,/g;
  let match;
  while ((match = entryPattern.exec(catalogText)) !== null) {
    entries.push({
      serviceName: match[1],
      displayName: match[2],
      fileName: match[3],
      sourceAsset: match[4] ?? null,
    });
  }
  if (entries.length === 0) throw new Error('No Fabric catalog entries were found');

  const sourceAssets = new Map();
  const destinationNames = new Set();
  for (const entry of entries) {
    if (!/^[a-z0-9-]+$/i.test(entry.fileName)) {
      throw new Error(`Unsafe Fabric icon destination name: ${entry.fileName}`);
    }
    if (destinationNames.has(entry.fileName.toLocaleLowerCase())) {
      throw new Error(`Duplicate Fabric icon destination name: ${entry.fileName}`);
    }
    destinationNames.add(entry.fileName.toLocaleLowerCase());
    if (!entry.sourceAsset) continue;
    const sourceKey = entry.sourceAsset.toLocaleLowerCase();
    if (sourceAssets.has(sourceKey)) {
      throw new Error(`Duplicate Fabric source asset: ${entry.sourceAsset}`);
    }
    sourceAssets.set(sourceKey, entry);
  }

  return { version, entries, sourceAssets };
}

function resolveDestination(fileName) {
  const destination = resolve(iconRoot, `${fileName}.svg`);
  const pathFromRoot = relative(iconRoot, destination);
  if (
    pathFromRoot === ''
    || pathFromRoot.startsWith('..')
    || isAbsolute(pathFromRoot)
  ) {
    throw new Error(`Fabric icon path escapes the library root: ${fileName}`);
  }
  return destination;
}

async function verifyIntegrity(archive, integrity) {
  const [algorithm, expected] = String(integrity ?? '').split('-', 2);
  if (!algorithm || !expected) throw new Error('The npm package has no integrity value');
  const actual = createHash(algorithm).update(archive).digest('base64');
  if (actual !== expected) throw new Error('The Fabric icon package failed integrity verification');
}

async function main() {
  const catalogText = await readFile(catalogPath, 'utf8');
  const catalog = parseCatalog(catalogText);
  const metadataResponse = await fetch(REGISTRY_URL);
  if (!metadataResponse.ok) {
    throw new Error(`Failed to query ${PACKAGE_NAME}: HTTP ${metadataResponse.status}`);
  }
  const metadata = await metadataResponse.json();
  const latestVersion = metadata['dist-tags']?.latest;
  if (latestVersion !== catalog.version) {
    throw new Error(
      `Fabric icon catalog is pinned to ${catalog.version}, but npm latest is ${latestVersion}. `
      + 'Review the new official families before syncing.',
    );
  }

  const packageMetadata = metadata.versions?.[catalog.version];
  const packageUrl = packageMetadata?.dist?.tarball;
  if (!packageUrl) throw new Error(`No tarball was found for ${PACKAGE_NAME} ${catalog.version}`);

  const packageResponse = await fetch(packageUrl);
  if (!packageResponse.ok) {
    throw new Error(`Failed to download ${packageUrl}: HTTP ${packageResponse.status}`);
  }
  const archive = Buffer.from(await packageResponse.arrayBuffer());
  if (archive.length > 50 * 1024 * 1024) {
    throw new Error(`Fabric icon package is unexpectedly large: ${archive.length} bytes`);
  }
  await verifyIntegrity(archive, packageMetadata.dist.integrity);

  const tempRoot = await mkdtemp(join(tmpdir(), 'azurediagarm-fabric-icons-'));
  try {
    const archivePath = join(tempRoot, 'package.tgz');
    await writeFile(archivePath, archive);
    await execFileAsync('tar', ['-xzf', archivePath, '-C', tempRoot]);

    const extractedPackage = resolve(tempRoot, 'package');
    const extractedPackageJson = JSON.parse(
      await readFile(resolve(extractedPackage, 'package.json'), 'utf8'),
    );
    if (
      extractedPackageJson.name !== PACKAGE_NAME
      || extractedPackageJson.version !== catalog.version
    ) {
      throw new Error('The extracted Fabric icon package identity is invalid');
    }

    const sourceRoot = resolve(extractedPackage, 'svg');
    const packageSvgNames = (await readdir(sourceRoot))
      .filter(fileName => /\.svg$/i.test(fileName))
      .sort((left, right) => left.localeCompare(right));
    const architectureAssets = packageSvgNames.filter(isArchitectureAsset);
    const architectureAssetSet = new Set(
      architectureAssets.map(fileName => fileName.toLocaleLowerCase()),
    );
    const missingCatalogAssets = architectureAssets.filter(
      fileName => !catalog.sourceAssets.has(fileName.toLocaleLowerCase()),
    );
    const staleCatalogAssets = [...catalog.sourceAssets.values()].filter(
      entry => !architectureAssetSet.has(entry.sourceAsset.toLocaleLowerCase()),
    );
    if (missingCatalogAssets.length > 0 || staleCatalogAssets.length > 0) {
      throw new Error([
        'Fabric catalog does not exactly cover the official architecture icon families.',
        missingCatalogAssets.length > 0
          ? `Missing catalog entries: ${missingCatalogAssets.join(', ')}`
          : '',
        staleCatalogAssets.length > 0
          ? `Removed package assets: ${staleCatalogAssets.map(entry => entry.sourceAsset).join(', ')}`
          : '',
      ].filter(Boolean).join('\n'));
    }

    await mkdir(iconRoot, { recursive: true });
    let changed = 0;
    const files = [];
    const categories = new Map();
    for (const sourceAsset of architectureAssets) {
      const definition = catalog.sourceAssets.get(sourceAsset.toLocaleLowerCase());
      const sourcePath = resolve(sourceRoot, sourceAsset);
      const destination = resolveDestination(definition.fileName);
      const content = await readFile(sourcePath);
      let existing = null;
      try {
        existing = await readFile(destination);
      } catch {
        // Missing files are created below.
      }
      if (!existing || !existing.equals(content)) {
        await writeFile(destination, content);
        changed += 1;
      }

      const category = classifyAsset(sourceAsset);
      categories.set(category, (categories.get(category) ?? 0) + 1);
      files.push({
        sourceAsset,
        path: `fabric/${basename(destination)}`,
        sha256: sha256(content),
      });
    }

    const localIconCount = catalog.entries.filter(entry => !entry.sourceAsset).length;
    const manifest = {
      schemaVersion: 1,
      sourcePage: SOURCE_PAGE,
      officialZipUrl: OFFICIAL_ZIP_URL,
      packageName: PACKAGE_NAME,
      packageVersion: catalog.version,
      packageUrl,
      packageIntegrity: packageMetadata.dist.integrity,
      packageSha256: sha256(archive),
      packageSvgCount: packageSvgNames.length,
      officialFamilyCount: files.length,
      localIconCount,
      paletteIconCount: files.length + localIconCount,
      excludedGeneralUiSvgCount: packageSvgNames.length - files.length,
      categories: Object.fromEntries(
        [...categories.entries()].sort(([left], [right]) => left.localeCompare(right)),
      ),
      files,
    };
    await writeFile(manifestPath, `${JSON.stringify(manifest, null, 2)}\n`, 'utf8');

    console.log(
      `[icons:sync:fabric] ${catalog.version}: ${files.length} official families, `
      + `${localIconCount} local symbol, ${changed} files added or updated`,
    );
    console.log(`[icons:sync:fabric] package SHA-256: ${manifest.packageSha256}`);
  } finally {
    await rm(tempRoot, { recursive: true, force: true });
  }
}

main().catch(error => {
  console.error(`[icons:sync:fabric] ${error.stack || error.message}`);
  process.exitCode = 1;
});
