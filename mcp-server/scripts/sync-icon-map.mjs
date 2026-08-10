#!/usr/bin/env node
// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

/**
 * Build-time helper: extract icons, aliases, and pricing metadata from the web
 * app's src/data/serviceIconMapping.ts and emit a JSON sidecar consumable by
 * the MCP server (avoids duplicating the application catalog).
 *
 * Run: node mcp-server/scripts/sync-icon-map.mjs
 */
import { readFileSync, writeFileSync } from 'node:fs';
import { fileURLToPath } from 'node:url';
import { dirname, resolve } from 'node:path';

const here = dirname(fileURLToPath(import.meta.url));
const repoRoot = resolve(here, '..', '..');
const sourcePath = resolve(repoRoot, 'src', 'data', 'serviceIconMapping.ts');
const fabricCatalogPath = resolve(repoRoot, 'src', 'data', 'fabricIconCatalog.ts');
const microsoftCatalogPath = resolve(repoRoot, 'src', 'data', 'microsoftProductIconCatalog.ts');
const outPath = resolve(here, '..', 'src', 'iconMap.generated.json');

const text = readFileSync(sourcePath, 'utf8');

// Naive but reliable parser: each entry is `'<Key>': { ... iconFile: 'xxx', category: 'yyy', ... }`.
// We capture every block until the matching closing brace at the same depth.
const map = {};
const entryRe = /'([^']+)'\s*:\s*\{/g;
let match;
while ((match = entryRe.exec(text)) !== null) {
  const key = match[1];
  // Walk forward to find the matching closing brace
  let depth = 1;
  let i = entryRe.lastIndex;
  while (i < text.length && depth > 0) {
    const ch = text[i];
    if (ch === '{') depth++;
    else if (ch === '}') depth--;
    i++;
  }
  const block = text.slice(match.index, i);
  const displayNameMatch = block.match(/displayName:\s*'([^']+)'/);
  const iconFileMatch = block.match(/iconFile:\s*'([^']+)'/);
  const categoryMatch = block.match(/category:\s*'([^']+)'/);
  const hasPricingDataMatch = block.match(/hasPricingData:\s*(true|false)/);
  const pricingServiceNameMatch = block.match(/pricingServiceName:\s*'([^']+)'/);
  const isUsageBasedMatch = block.match(/isUsageBased:\s*(true|false)/);
  const costRangeMatch = block.match(/costRange:\s*'([^']+)'/);
  if (iconFileMatch && categoryMatch) {
    // Capture aliases so the MCP renderer can resolve real-world type variants
    // (e.g. "Blob Storage", "Azure Cache for Redis") to the right icon.
    const aliasesMatch = block.match(/aliases:\s*\[([^\]]*)\]/);
    const aliases = aliasesMatch
      ? [...aliasesMatch[1].matchAll(/'([^']+)'/g)].map(m => m[1])
      : [];
    map[key] = {
      displayName: displayNameMatch?.[1] ?? key,
      iconFile: iconFileMatch[1],
      category: categoryMatch[1],
      aliases,
      hasPricingData: hasPricingDataMatch?.[1] === 'true',
      ...(pricingServiceNameMatch ? { pricingServiceName: pricingServiceNameMatch[1] } : {}),
      ...(isUsageBasedMatch ? { isUsageBased: isUsageBasedMatch[1] === 'true' } : {}),
      ...(costRangeMatch ? { costRange: costRangeMatch[1] } : {}),
    };
  }
}

// Fabric icons are defined through a typed catalog and spread into the runtime
// mapping, so merge those entries explicitly for the standalone MCP renderer.
const fabricText = readFileSync(fabricCatalogPath, 'utf8');
const fabricEntryRe = /^\s*defineFabricIcon\('([^']+)',\s*'([^']+)',\s*'([^']+)',\s*(?:null|'[^']+'),\s*'[^']+',\s*'([^']+)',\s*\[([^\]]*)\]\s*(?:,\s*\{([^}]*)\})?\s*\)/gm;
let fabricCount = 0;
while ((match = fabricEntryRe.exec(fabricText)) !== null) {
  const [, key, displayName, iconFile, kind, aliasSource, optionsSource = ''] = match;
  const aliases = [...aliasSource.matchAll(/'([^']+)'/g)].map(alias => alias[1]);
  const hasPricingDataMatch = optionsSource.match(/hasPricingData:\s*(true|false)/);
  const pricingServiceNameMatch = optionsSource.match(/pricingServiceName:\s*'([^']+)'/);
  const isUsageBasedMatch = optionsSource.match(/isUsageBased:\s*(true|false)/);
  const costRangeMatch = optionsSource.match(/costRange:\s*'([^']+)'/);
  const includeInServiceMapMatch = optionsSource.match(/includeInServiceMap:\s*(true|false)/);
  if (includeInServiceMapMatch?.[1] === 'false') continue;
  const consumesCapacity = kind === 'workload' || kind === 'item' || kind === 'state';
  if (displayName !== key && !aliases.includes(displayName)) aliases.unshift(displayName);
  fabricCount++;
  map[key] = {
    displayName,
    iconFile,
    category: 'fabric',
    aliases,
    hasPricingData: hasPricingDataMatch?.[1] === 'true',
    ...(pricingServiceNameMatch ? { pricingServiceName: pricingServiceNameMatch[1] } : {}),
    ...(isUsageBasedMatch ? { isUsageBased: isUsageBasedMatch[1] === 'true' } : {}),
    ...(costRangeMatch
      ? { costRange: costRangeMatch[1] }
      : consumesCapacity
        ? { costRange: '$0 (consumes Fabric capacity)' }
        : {}),
  };
}
if (fabricCount === 0) {
  console.error(`[sync-icon-map] no entries extracted from ${fabricCatalogPath}`);
  process.exit(1);
}

// Power Platform / Dynamics 365 icons follow the same catalog-and-spread shape
// as Fabric, so they need their own parse pass or the MCP server silently ships
// a catalog that is 24 services behind the web app.
const microsoftText = readFileSync(microsoftCatalogPath, 'utf8');

/** Parse a `const NAME: Record<...> = { 'key': 'value', ... };` literal. */
function parseFamilyRecord(source, name) {
  const block = source.match(new RegExp(`const ${name}:[^=]*=\\s*\\{([\\s\\S]*?)\\n\\};`));
  if (!block) {
    console.error(`[sync-icon-map] could not parse ${name} from ${microsoftCatalogPath}`);
    process.exit(1);
  }
  return Object.fromEntries(
    [...block[1].matchAll(/'([^']+)'\s*:\s*'([^']+)'/g)].map(entry => [entry[1], entry[2]]),
  );
}

const microsoftFamilyCategory = parseFamilyRecord(microsoftText, 'FAMILY_CATEGORY');
const microsoftFamilyCostRange = parseFamilyRecord(microsoftText, 'FAMILY_COST_RANGE');

const microsoftEntryRe = /^\s*defineMicrosoftProductIcon\('([^']+)',\s*'([^']+)',\s*'([^']+)',\s*'[^']+',\s*'([^']+)',\s*'[^']+',\s*'[^']+',\s*\[([^\]]*)\]\s*\)/gm;
let microsoftCount = 0;
while ((match = microsoftEntryRe.exec(microsoftText)) !== null) {
  const [, key, displayName, iconFile, family, aliasSource] = match;
  const category = microsoftFamilyCategory[family];
  if (!category) {
    console.error(`[sync-icon-map] unknown Microsoft product icon family: ${family}`);
    process.exit(1);
  }
  const aliases = [...aliasSource.matchAll(/'([^']+)'/g)].map(alias => alias[1]);
  if (displayName !== key && !aliases.includes(displayName)) aliases.unshift(displayName);
  // These products are licensed per user or per app, never through an Azure
  // meter, so they always report as having no Azure pricing data.
  map[key] = {
    displayName,
    iconFile,
    category,
    aliases,
    hasPricingData: false,
    isUsageBased: false,
    ...(microsoftFamilyCostRange[family] ? { costRange: microsoftFamilyCostRange[family] } : {}),
  };
  microsoftCount++;
}
if (microsoftCount === 0) {
  console.error(`[sync-icon-map] no entries extracted from ${microsoftCatalogPath}`);
  process.exit(1);
}

const count = Object.keys(map).length;
if (count === 0) {
  console.error(`[sync-icon-map] no entries extracted from ${sourcePath}`);
  process.exit(1);
}

writeFileSync(outPath, JSON.stringify(map, null, 2) + '\n', 'utf8');
console.log(`[sync-icon-map] wrote ${count} entries to ${outPath}`);

// ── Embed real Azure icon SVGs as data URIs ────────────────────────────
// For each referenced icon file, read the SVG, lightly minify, and base64
// encode into a data URI. The renderer inlines these via <image> so diagrams
// use the official Azure glyphs instead of emoji. <image> data URIs avoid the
// gradient-id collisions that inlining raw <svg> would cause.
const iconsRoot = resolve(repoRoot, 'Azure_Public_Service_Icons', 'Icons');
const svgOutPath = resolve(here, '..', 'src', 'iconSvgs.generated.json');
const svgs = {};
let embedded = 0;
let missing = 0;
const seen = new Set();
for (const entry of Object.values(map)) {
  const { iconFile, category } = entry;
  if (seen.has(iconFile)) continue;
  seen.add(iconFile);
  const svgPath = resolve(iconsRoot, category, `${iconFile}.svg`);
  try {
    let svg = readFileSync(svgPath, 'utf8');
    svg = svg
      .replace(/<\?xml[^>]*\?>/g, '')
      .replace(/<!--[\s\S]*?-->/g, '')
      .replace(/>\s+</g, '><')
      .trim();
    svgs[iconFile] = `data:image/svg+xml;base64,${Buffer.from(svg, 'utf8').toString('base64')}`;
    embedded++;
  } catch {
    missing++;
  }
}

writeFileSync(svgOutPath, JSON.stringify(svgs) + '\n', 'utf8');
console.log(`[sync-icon-map] embedded ${embedded} icon SVGs (${missing} missing) to ${svgOutPath}`);
