#!/usr/bin/env node
// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

import { execFile } from 'node:child_process';
import { createHash } from 'node:crypto';
import { mkdir, mkdtemp, readFile, readdir, rm, writeFile } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { dirname, join, resolve } from 'node:path';
import { promisify } from 'node:util';
import { fileURLToPath, pathToFileURL } from 'node:url';

const execFileAsync = promisify(execFile);
const PACKAGE_NAME = '@fabric-msft/svg-icons';
const REGISTRY_URL = 'https://registry.npmjs.org/@fabric-msft%2Fsvg-icons';
const SOURCE_PAGE = 'https://learn.microsoft.com/en-us/fabric/fundamentals/icons';
const ARTICLE_PAGE = 'https://debruyn.dev/2026/all-microsoft-fabric-icons-for-diagramming-2026-edition/';
const here = dirname(fileURLToPath(import.meta.url));
const repoRoot = resolve(here, '..');
const iconRoot = resolve(repoRoot, 'Azure_Public_Service_Icons', 'Icons', 'fabric');
const manifestPath = resolve(repoRoot, 'Azure_Public_Service_Icons', 'fabric-manifest.json');
const lockPath = resolve(here, 'fabric-icon-sources.json');
const sha256 = value => createHash('sha256').update(value).digest('hex');

export function parseCatalog(text) {
  const version = /FABRIC_ICON_PACKAGE_VERSION\s*=\s*'([^']+)'/.exec(text)?.[1];
  if (!version) throw new Error('Fabric catalog does not declare its package version');
  const pattern = /^\s*defineFabricIcon\('([^']+)',\s*'([^']+)',\s*'([^']+)',\s*(?:'([^']+\.svg)'|null),\s*'([^']+)',\s*'([^']+)',\s*\[[^\]]*\]\s*(?:,\s*\{([^}]*)\})?\s*\)/gm;
  const entries = [...text.matchAll(pattern)].map(([, serviceName, displayName, fileName, sourceAsset, group, kind, options = '']) => ({
    serviceName, displayName, fileName, sourceAsset: sourceAsset ?? null, group, kind,
    sourceVersion: /sourceVersion:\s*'([^']+)'/.exec(options)?.[1] ?? (sourceAsset ? version : 'local'),
  }));
  if (!entries.length) throw new Error('No Fabric catalog entries were found');
  const destinations = new Set();
  const sources = new Set();
  for (const entry of entries) {
    if (!/^[a-z0-9-]+$/.test(entry.fileName) || destinations.has(entry.fileName)) {
      throw new Error(`Unsafe or duplicate Fabric destination: ${entry.fileName}`);
    }
    destinations.add(entry.fileName);
    if (entry.sourceAsset) {
      if (!/^[a-z0-9_ .-]+\.svg$/i.test(entry.sourceAsset) || entry.sourceAsset.includes('..')) {
        throw new Error(`Unsafe Fabric source: ${entry.sourceAsset}`);
      }
      const key = `${entry.sourceVersion}/${entry.sourceAsset}`.toLowerCase();
      if (sources.has(key)) throw new Error(`Duplicate Fabric source: ${key}`);
      sources.add(key);
    }
  }
  return { version, entries };
}

export function assertSafeSvg(content, name) {
  const text = content.toString('utf8');
  if (
    !/<svg[\s>]/i.test(text)
    || /<\s*(?:script|foreignObject|image|use|style|a)[\s/>]/i.test(text)
    || /\bon[a-z]+\s*=/i.test(text)
    || /(?:href\s*=|<!DOCTYPE|<!ENTITY|<\?xml-stylesheet|@import)/i.test(text)
    || /url\(\s*['"]?(?!#)[^)]/i.test(text)
  ) {
    throw new Error(`Unsafe or invalid Fabric SVG: ${name}`);
  }
}

async function download(url) {
  const response = await fetch(url, { signal: AbortSignal.timeout(60_000) });
  if (!response.ok) throw new Error(`Failed to download ${url}: HTTP ${response.status}`);
  const buffer = Buffer.from(await response.arrayBuffer());
  if (buffer.length > 50 * 1024 * 1024) throw new Error(`Oversized Fabric download: ${url}`);
  return buffer;
}

async function readOptional(path) {
  try {
    return await readFile(path);
  } catch (error) {
    if (error.code !== 'ENOENT') throw error;
    return null;
  }
}

export async function main() {
  const catalog = parseCatalog(await readFile(resolve(repoRoot, 'src', 'data', 'fabricIconCatalog.ts'), 'utf8'));
  const lock = JSON.parse(await readFile(lockPath, 'utf8'));
  const metadata = JSON.parse((await download(REGISTRY_URL)).toString('utf8'));
  const latest = metadata['dist-tags']?.latest;
  if (latest !== catalog.version) {
    console.warn(`[icons:sync:fabric] npm latest is ${latest}; retaining reviewed pin ${catalog.version}. Review upgrades explicitly.`);
  }
  const tempRoot = await mkdtemp(join(tmpdir(), 'azurediagarm-fabric-icons-'));
  try {
    const packages = new Map();
    for (const pin of lock.packages) {
      const dist = metadata.versions?.[pin.version]?.dist;
      if (dist?.tarball !== pin.url || dist?.integrity !== pin.integrity) {
        throw new Error(`Fabric registry metadata differs from the source lock: ${pin.version}`);
      }
      const archive = await download(pin.url);
      if (
        `sha512-${createHash('sha512').update(archive).digest('base64')}` !== pin.integrity
        || sha256(archive) !== pin.sha256
      ) throw new Error(`Fabric package failed integrity verification: ${pin.version}`);
      const directory = join(tempRoot, pin.version);
      await mkdir(directory);
      const archivePath = join(directory, 'package.tgz');
      await writeFile(archivePath, archive);
      await execFileAsync('tar', ['-xzf', 'package.tgz'], { cwd: directory });
      const packageRoot = join(directory, 'package');
      const identity = JSON.parse(await readFile(join(packageRoot, 'package.json'), 'utf8'));
      if (identity.name !== PACKAGE_NAME || identity.version !== pin.version) {
        throw new Error(`Invalid extracted Fabric package: ${pin.version}`);
      }
      const sourceRoot = join(packageRoot, pin.svgDirectory);
      packages.set(pin.version, { sourceRoot, names: await readdir(sourceRoot) });
    }
    const current = packages.get(catalog.version);
    if (!current) throw new Error('The current Fabric package is not locked');
    const architectureNames = current.names.filter(name =>
      /_48(?:_color|_item|_non-item)?\.svg$/i.test(name) || name === 'sample_workload_32_color.svg');
    const catalogNames = catalog.entries.filter(entry => entry.sourceVersion === catalog.version).map(entry => entry.sourceAsset);
    if (
      architectureNames.length !== catalogNames.length
      || architectureNames.some(name => !catalogNames.includes(name))
    ) throw new Error('Fabric catalog must cover every current 48px architecture family exactly once');

    const articleByDestination = new Map(lock.articleIcons.map(icon => [icon.fileName, icon]));
    if (articleByDestination.size !== lock.articleIcons.length) throw new Error('Duplicate article mapping');
    for (const icon of lock.articleIcons) {
      if (!catalog.entries.some(entry => entry.fileName === icon.fileName)) {
        throw new Error(`Article icon is missing from the catalog: ${icon.title}`);
      }
      if (!icon.url.startsWith(ARTICLE_PAGE) || !/^[a-f0-9]{64}$/.test(icon.sha256)) {
        throw new Error(`Invalid article source pin: ${icon.title}`);
      }
    }

    // Validate the entire inventory before changing any shipped artwork.
    const prepared = [];
    for (const entry of catalog.entries) {
      const destination = join(iconRoot, `${entry.fileName}.svg`);
      const article = articleByDestination.get(entry.fileName);
      let content;
      let source;
      if (entry.sourceVersion === 'article-2026') {
        if (!article || entry.sourceAsset !== new URL(article.url).pathname.split('/').at(-1)) {
          throw new Error(`Missing article source pin: ${entry.fileName}`);
        }
        content = await readOptional(destination);
        if (!content || sha256(content) !== article.sha256) content = await download(article.url);
        source = { sourceType: 'article', sourceUrl: article.url, sourceVersion: entry.sourceVersion };
      } else if (entry.sourceAsset) {
        const pkg = packages.get(entry.sourceVersion);
        if (!pkg || !pkg.names.includes(entry.sourceAsset)) throw new Error(`Missing Fabric source: ${entry.sourceAsset}`);
        content = await readFile(join(pkg.sourceRoot, entry.sourceAsset));
        source = { sourceType: 'npm', sourceVersion: entry.sourceVersion, packageName: PACKAGE_NAME };
      } else {
        content = await readFile(destination);
        source = { sourceType: 'local', sourceVersion: 'local' };
      }
      assertSafeSvg(content, entry.fileName);
      if (article && sha256(content) !== article.sha256) throw new Error(`Article SVG hash mismatch: ${article.title}`);
      prepared.push({
        content, destination,
        record: {
          serviceName: entry.serviceName, sourceAsset: entry.sourceAsset, ...source,
          path: `fabric/${entry.fileName}.svg`, sha256: sha256(content),
          ...(article ? { articleTitle: article.title, articleSection: article.section, articleUrl: article.url } : {}),
        },
      });
    }
    await mkdir(iconRoot, { recursive: true });
    let changed = 0;
    for (const item of prepared) {
      const existing = await readOptional(item.destination);
      if (!existing?.equals(item.content)) {
        await writeFile(item.destination, item.content);
        changed++;
      }
    }
    const manifest = {
      schemaVersion: 2, sourcePage: SOURCE_PAGE, articlePage: ARTICLE_PAGE,
      officialZipUrl: 'https://github.com/microsoft/fabric-samples/blob/main/docs-samples/Icons.zip',
      attribution: 'Microsoft graphics and trademarks; article compiled by Sam Debruyn. See FABRIC-ICONS.md for permitted diagramming use.',
      packageName: PACKAGE_NAME, packageVersion: catalog.version, packages: lock.packages,
      officialFamilyCount: catalogNames.length,
      legacyIconCount: catalog.entries.filter(entry => entry.kind === 'legacy').length,
      uiIconCount: catalog.entries.filter(entry => entry.kind === 'ui').length,
      localIconCount: catalog.entries.filter(entry => !entry.sourceAsset).length,
      articleIconCount: lock.articleIcons.length, paletteIconCount: prepared.length,
      files: prepared.map(item => item.record),
    };
    await writeFile(manifestPath, `${JSON.stringify(manifest, null, 2)}\n`);
    console.log(`[icons:sync:fabric] ${manifest.paletteIconCount} icons (${manifest.officialFamilyCount} current, ${manifest.legacyIconCount} legacy, ${manifest.uiIconCount} UI, ${manifest.localIconCount} local); ${changed} SVGs updated`);
  } finally {
    await rm(tempRoot, { recursive: true, force: true });
  }
}

if (process.argv[1] && import.meta.url === pathToFileURL(resolve(process.argv[1])).href) {
  main().catch(error => {
    console.error(`[icons:sync:fabric] ${error.stack || error.message}`);
    process.exitCode = 1;
  });
}
