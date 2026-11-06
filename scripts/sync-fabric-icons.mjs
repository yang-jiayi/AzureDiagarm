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
const articleArchiveRoot = resolve(repoRoot, 'Azure_Public_Service_Icons', 'Article_Fabric_Icons', '2026');
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
  const hrefAttributeCount = (text.match(/\b(?:xlink:)?href\s*=/gi) ?? []).length;
  const hrefs = [...text.matchAll(/\b(?:xlink:)?href\s*=\s*(['"])(.*?)\1/gi)];
  if (
    !/<svg[\s>]/i.test(text)
    || /<\s*(?:script|foreignObject|image|style|a)[\s/>]/i.test(text)
    || /\bon[a-z]+\s*=/i.test(text)
    || hrefs.length !== hrefAttributeCount
    || hrefs.some(([, , reference]) => !/^#[\w:.-]+$/.test(reference.trim()))
    || (/<\s*use\b/i.test(text) && hrefs.length < (text.match(/<\s*use\b/gi) ?? []).length)
    || /(?:<!DOCTYPE|<!ENTITY|<\?xml-stylesheet|@import)/i.test(text)
    || /url\(\s*['"]?(?!#)[^)]/i.test(text)
  ) {
    throw new Error(`Unsafe or invalid Fabric SVG: ${name}`);
  }
}

async function download(url) {
  const response = await fetch(url, { signal: AbortSignal.timeout(60_000) });
  if (!response.ok) throw new Error(`Failed to download ${url}: HTTP ${response.status}`);
  const maxBytes = 50 * 1024 * 1024;
  const declaredBytes = Number(response.headers.get('content-length'));
  if (Number.isFinite(declaredBytes) && declaredBytes > maxBytes) {
    throw new Error(`Oversized Fabric download: ${url}`);
  }
  const reader = response.body?.getReader();
  if (!reader) throw new Error(`Unreadable Fabric download: ${url}`);
  const chunks = [];
  let total = 0;
  for (;;) {
    const { done, value } = await reader.read();
    if (done) break;
    total += value.byteLength;
    if (total > maxBytes) {
      await reader.cancel();
      throw new Error(`Oversized Fabric download: ${url}`);
    }
    chunks.push(Buffer.from(value));
  }
  return Buffer.concat(chunks, total);
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
  const checkOnly = process.argv.includes('--check');
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
      if (pin.publishedAt && metadata.time?.[pin.version] !== pin.publishedAt) {
        throw new Error(`Fabric package publication timestamp differs from the source lock: ${pin.version}`);
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

    const packageComparisonPin = lock.packageComparisons?.find(
      comparison => comparison.toVersion === catalog.version,
    );
    const previousPackage = packageComparisonPin
      ? packages.get(packageComparisonPin.fromVersion)
      : null;
    if (!packageComparisonPin || !previousPackage) {
      throw new Error(`Fabric package comparison is not locked for ${catalog.version}`);
    }
    const previousArchitectureNames = previousPackage.names.filter(name =>
      /_48(?:_color|_item|_non-item)?\.svg$/i.test(name) || name === 'sample_workload_32_color.svg');
    const previousHashes = new Map();
    const currentHashes = new Map();
    for (const name of previousArchitectureNames) {
      previousHashes.set(name, sha256(await readFile(join(previousPackage.sourceRoot, name))));
    }
    for (const name of architectureNames) {
      currentHashes.set(name, sha256(await readFile(join(current.sourceRoot, name))));
    }
    const renamedFrom = new Set();
    const renamedTo = new Set();
    const renamed = packageComparisonPin.renamed.map(item => {
      if (
        !previousHashes.has(item.from) || !currentHashes.has(item.to)
        || renamedFrom.has(item.from) || renamedTo.has(item.to)
      ) throw new Error(`Invalid Fabric package rename mapping: ${item.from} -> ${item.to}`);
      renamedFrom.add(item.from);
      renamedTo.add(item.to);
      return {
        from: item.from, to: item.to,
        previousSha256: previousHashes.get(item.from), sha256: currentHashes.get(item.to),
      };
    });
    const packageComparison = {
      fromVersion: packageComparisonPin.fromVersion,
      toVersion: catalog.version,
      previousFamilyCount: previousArchitectureNames.length,
      currentFamilyCount: architectureNames.length,
      added: architectureNames
        .filter(name => !previousHashes.has(name) && !renamedTo.has(name))
        .map(asset => ({ asset, sha256: currentHashes.get(asset) })),
      renamed,
      removed: previousArchitectureNames
        .filter(name => !currentHashes.has(name) && !renamedFrom.has(name))
        .map(asset => ({ asset, sha256: previousHashes.get(asset) })),
      updated: architectureNames
        .filter(name => previousHashes.has(name) && previousHashes.get(name) !== currentHashes.get(name))
        .map(asset => ({
          asset, previousSha256: previousHashes.get(asset), sha256: currentHashes.get(asset),
        })),
      unchangedCount: architectureNames.filter(
        name => previousHashes.get(name) === currentHashes.get(name),
      ).length,
    };

    const articleByDestination = new Map(lock.articleIcons.map(icon => [icon.fileName, icon]));
    if (articleByDestination.size !== lock.articleIcons.length) throw new Error('Duplicate article mapping');
    if (lock.articleIcons.length !== 101) {
      throw new Error(`Expected 101 pinned article icons, found ${lock.articleIcons.length}`);
    }
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
    const preparedByFileName = new Map();
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
      const item = {
        content, destination,
        record: {
          serviceName: entry.serviceName, sourceAsset: entry.sourceAsset, ...source,
          path: `fabric/${entry.fileName}.svg`, sha256: sha256(content),
        },
      };
      prepared.push(item);
      preparedByFileName.set(entry.fileName, item);
    }

    const preparedArticleAssets = [];
    for (const article of lock.articleIcons) {
      const entry = catalog.entries.find(candidate => candidate.fileName === article.fileName);
      const activeItem = preparedByFileName.get(article.fileName);
      let destination = activeItem.destination;
      let content = activeItem.content;
      if (sha256(content) !== article.sha256) {
        content = await download(article.url);
        assertSafeSvg(content, article.title);
        if (sha256(content) !== article.sha256) {
          throw new Error(`Article SVG hash mismatch: ${article.title}`);
        }
        destination = join(articleArchiveRoot, `${article.fileName}.svg`);
      }
      preparedArticleAssets.push({
        content,
        destination,
        record: {
          title: article.title, section: article.section, sourceUrl: article.url,
          path: destination.startsWith(articleArchiveRoot)
            ? `Article_Fabric_Icons/2026/${article.fileName}.svg`
            : `Icons/fabric/${article.fileName}.svg`,
          sha256: article.sha256,
          activePath: `Icons/fabric/${entry.fileName}.svg`,
        },
      });
      if (sha256(activeItem.content) === article.sha256) {
        activeItem.record.articleTitle = article.title;
        activeItem.record.articleSection = article.section;
        activeItem.record.articleUrl = article.url;
      }
    }
    await mkdir(iconRoot, { recursive: true });
    if (preparedArticleAssets.some(item => item.destination.startsWith(articleArchiveRoot))) {
      await mkdir(articleArchiveRoot, { recursive: true });
    }
    let changed = 0;
    const drift = [];
    for (const item of [...prepared, ...preparedArticleAssets.filter(item => item.destination.startsWith(articleArchiveRoot))]) {
      const existing = await readOptional(item.destination);
      if (!existing?.equals(item.content)) {
        if (checkOnly) {
          drift.push(item.record.path);
        } else {
          await writeFile(item.destination, item.content);
          changed++;
        }
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
      officialAssets: architectureNames.slice().sort((left, right) => left.localeCompare(right)),
      packageComparison,
      files: prepared.map(item => item.record),
      articleAssets: preparedArticleAssets.map(item => item.record),
    };
    if (checkOnly) {
      const currentManifest = JSON.parse(await readFile(manifestPath, 'utf8'));
      if (drift.length > 0) {
        throw new Error(`Fabric SVG files are out of date: ${drift.join(', ')}`);
      }
      if (JSON.stringify(currentManifest) !== JSON.stringify(manifest)) {
        throw new Error('Fabric icon manifest differs from the pinned source inventory');
      }
    } else {
      await writeFile(manifestPath, `${JSON.stringify(manifest, null, 2)}\n`);
    }
    console.log(`[icons:sync:fabric] ${manifest.paletteIconCount} icons (${manifest.officialFamilyCount} current, ${manifest.legacyIconCount} legacy, ${manifest.uiIconCount} UI, ${manifest.localIconCount} local); ${checkOnly ? 'check only' : `${changed} SVGs updated`}`);
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
