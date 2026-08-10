#!/usr/bin/env node
// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

/**
 * Build-time helper: distill the web app's Azure Retail Prices JSON
 * (public/pricing/regions/<region>/<service>.json) into a compact per-region
 * sidecar the MCP server can bundle and query at runtime WITHOUT shipping the
 * full dataset.
 *
 * Those files are compacted in place by scripts/prep-pricing-data.mjs (hoisted
 * ServiceName/type, dropped read-never fields), so they are expanded here with
 * the very same `expandPricingData` the browser runtime uses. Reading them raw
 * would silently distil zero entries, because `type` no longer exists per item.
 *
 * This mirrors sync-icon-map.mjs: the source of truth stays in the web app;
 * we generate a small derived artifact (src/pricing.generated.json, a few KB).
 *
 * Parity note: the monthly-cost derivation replicates the web app's
 * regionalPricingService.parsePricingTiers() unit-of-measure handling
 * (/Month as-is, /Day ×30, per-1K ×100, else hourly ×730) so numbers line up.
 *
 * AI (foundry_*) and Microsoft Fabric files need special per-product / F-SKU
 * handling and are intentionally skipped in this first cut — those services
 * fall back to the catalog costRange in estimate_costs. (Tracked for P0-1b.)
 *
 * Run: node mcp-server/scripts/sync-pricing-data.mjs
 */
import { readFileSync, writeFileSync, readdirSync, existsSync } from 'node:fs';
import { fileURLToPath } from 'node:url';
import { dirname, resolve, basename } from 'node:path';
import { expandPricingData } from '../../scripts/prep-pricing-data.mjs';

const here = dirname(fileURLToPath(import.meta.url));
const repoRoot = resolve(here, '..', '..');
const regionsRoot = resolve(repoRoot, 'public', 'pricing', 'regions');
const outPath = resolve(here, '..', 'src', 'pricing.generated.json');

/**
 * Read one compacted pricing file and restore the shape the distillers expect.
 * `PricesAsOf` is the file-level vintage hoisted by the compaction step, since
 * per-item `effectiveStartDate` is dropped there.
 */
function loadPricingFile(filePath) {
  let compact;
  try {
    compact = JSON.parse(readFileSync(filePath, 'utf8'));
  } catch {
    return null;
  }
  const expanded = expandPricingData(compact);
  const pricesAsOf = typeof compact?.PricesAsOf === 'string' ? compact.PricesAsOf : null;
  return { ...expanded, pricesAsOf };
}

// Files kept as honest catalog ranges. Foundry (AI) meters are usage-based
// (per-token / per-transaction) so a representative monthly is misleading; they
// stay as catalog ranges. Microsoft Fabric IS distilled specially below (F-SKU
// capacity is a real fixed monthly reservation). All other non-instance /
// usage-based / composite services are handled by the "default-SKU-only" rule.
const SKIP_FILES = new Set(['foundry_models', 'foundry_tools']);

const HOURS_PER_MONTH = 730;

/**
 * Representative-SKU map (P0-1a).
 *
 * "expected" should approximate a TYPICAL production deployment. We only emit a
 * numeric estimate for services whose per-SKU meter equals a deployable unit
 * (compute/db/cache/search/APIM/AKS) — and only when one of these preferred
 * default SKUs matches (case-insensitive substring, priority order). Everything
 * else (usage/consumption/per-GB/composite billing) is NOT distilled and falls
 * back to the curated catalog range in estimate_costs, which is far more honest
 * than a median/percentile of per-unit meters. low/high still span the full
 * cheapest→most-expensive SKU range for the distilled services.
 *
 * Keyed by file stem (serviceName lowercased + underscores).
 */
const REPRESENTATIVE_SKUS = {
  azure_app_service: ['p1 v3', 'p1v3', 'p1 v2', 'p1v2', 's1', 'b1'],
  redis_cache: ['c1', 'c2', 'basic c1', 'standard c1'],
  sql_database: ['s3', 's2', '100 dtu', 'general purpose', 's1'],
  virtual_machines: ['d2s v5', 'd2s v4', 'd2as v5', 'd2 v3', 'd2s v3', 'b2ms'],
  azure_kubernetes_service: ['standard', 'base'],
  azure_cognitive_search: ['standard s1', 'standard', 'basic'],
  api_management: ['developer', 'basic', 'standard'],
};

/** Compute a monthly price from a single retail-price item (parity with web app). */
function monthlyFromItem(item) {
  const unit = (item.unitOfMeasure || '1 Hour');
  const rate = item.retailPrice || item.unitPrice || 0;
  if (unit.includes('/Month') || unit.includes('1/Month')) return rate;
  if (unit.includes('/Day') || unit.includes('1/Day')) return rate * 30;
  if (unit === '1K' || unit.includes('1000')) return rate * 100;
  return rate * HOURS_PER_MONTH;
}

/** Median of a numeric array (sorted copy). */
function median(nums) {
  if (nums.length === 0) return 0;
  const s = [...nums].sort((a, b) => a - b);
  const mid = Math.floor(s.length / 2);
  return s.length % 2 ? s[mid] : (s[mid - 1] + s[mid]) / 2;
}

const round2 = (n) => Math.round(n * 100) / 100;

/** Most-frequent value (mode) of a numeric array, else fallback. */
function modeOrDefault(nums, fallback) {
  if (nums.length === 0) return fallback;
  const counts = new Map();
  let best = fallback;
  let bestCount = -1;
  for (const v of nums) {
    const c = (counts.get(v) || 0) + 1;
    counts.set(v, c);
    if (c > bestCount) { bestCount = c; best = v; }
  }
  return best;
}

/**
 * Distill Microsoft Fabric capacity (P0-1b). Unlike usage-based AI/storage,
 * Fabric F-SKU capacity is a fixed monthly reservation: monthly = per-CU-hour
 * rate x CUs x 730. Rate = mode of the "Capacity Usage CU" 1-Hour consumption
 * meters (~$0.18). Band: F2 (low) -> F8 (expected) -> F64 (high). Mirrors the
 * web app's getFabricRegionalPricing.
 */
function distillFabric(filePath) {
  const parsed = loadPricingFile(filePath);
  if (!parsed) return null;
  const items = Array.isArray(parsed.Items) ? parsed.Items : [];
  const currency = parsed.BillingCurrency || 'USD';
  const rates = [];
  for (const i of items) {
    if (
      i.type === 'Consumption' &&
      (i.unitOfMeasure || '') === '1 Hour' &&
      /Capacity Usage CU/i.test(i.meterName || '')
    ) {
      const r = i.retailPrice || i.unitPrice || 0;
      if (r > 0) rates.push(r);
    }
  }
  if (rates.length === 0) return null;
  const rate = modeOrDefault(rates, 0.18);
  const monthly = (cu) => round2(rate * cu * HOURS_PER_MONTH);
  return {
    low: monthly(2), // F2
    expected: monthly(8), // F8
    high: monthly(64), // F64
    reservedExpected: null,
    reservedRatio: null,
    currency,
    sampleSku: 'F8',
    expectedBasis: 'fabric-capacity:F8',
    tierCount: 3,
    pricesAsOf: parsed.pricesAsOf,
  };
}

/**
 * Distill one service JSON file into representative monthly costs.
 * Returns null when there is no usable Consumption pricing.
 */
function distillFile(filePath, stem) {
  const parsed = loadPricingFile(filePath);
  if (!parsed) return null;
  const items = Array.isArray(parsed.Items) ? parsed.Items : [];
  const currency = parsed.BillingCurrency || 'USD';

  // Consumption meters only (exclude reservations/spot rows; those are handled
  // via each item's savingsPlan for the reserved-term estimate).
  const consumption = items.filter((i) => i.type === 'Consumption');
  if (consumption.length === 0) return null;

  // Dedupe to the cheapest monthly per SKU (matches web app tier parsing).
  const perSku = new Map();
  const savingsRatios = [];

  for (const item of consumption) {
    const sku = item.skuName || item.armSkuName;
    if (!sku) continue;
    const monthly = monthlyFromItem(item);
    if (!perSku.has(sku) || monthly < perSku.get(sku)) perSku.set(sku, monthly);

    // 1-Year savings-plan ratio (reserved-term discount) when present.
    const plans = Array.isArray(item.savingsPlan) ? item.savingsPlan : [];
    const oneYear = plans.find((p) => /1\s*year/i.test(p.term || ''));
    const retail = item.retailPrice || item.unitPrice || 0;
    if (oneYear && retail > 0) {
      const rp = oneYear.unitPrice || oneYear.retailPrice || 0;
      if (rp > 0) savingsRatios.push(rp / retail);
    }
  }

  const monthlies = [...perSku.values()].filter((n) => Number.isFinite(n));
  if (monthlies.length === 0) return null;

  const sorted = [...monthlies].sort((a, b) => a - b);
  const low = sorted[0];
  const high = sorted[sorted.length - 1];

  const paidSorted = sorted.filter((m) => m > 0);
  if (paidSorted.length === 0) return null;

  // "expected" = a typical-deployment SKU. We ONLY emit a numeric estimate when
  // a trusted default SKU matches; usage/consumption/composite services (no
  // default) return null and fall back to the curated catalog range.
  let expected = null;
  let sampleSku = '';
  let expectedBasis = null;

  const patterns = REPRESENTATIVE_SKUS[stem];
  if (patterns) {
    for (const pat of patterns) {
      let match = null;
      for (const [sku, m] of perSku.entries()) {
        // Skip geo-replica / failover SKUs — not a representative primary cost.
        if (/secondary|failover|passive/i.test(sku)) continue;
        if (m > 0 && sku.toLowerCase().includes(pat)) {
          if (!match || m < match.m) match = { sku, m };
        }
      }
      if (match) {
        expected = match.m;
        sampleSku = match.sku;
        expectedBasis = `default-sku:${pat}`;
        break;
      }
    }
  }

  if (expected == null) return null;

  const reservedRatio = savingsRatios.length ? median(savingsRatios) : null;

  return {
    low: round2(low),
    expected: round2(expected),
    high: round2(high),
    reservedExpected: reservedRatio != null ? round2(expected * reservedRatio) : null,
    reservedRatio: reservedRatio != null ? round2(reservedRatio) : null,
    currency,
    sampleSku,
    expectedBasis,
    tierCount: monthlies.length,
    pricesAsOf: parsed.pricesAsOf,
  };
}

function main() {
  if (!existsSync(regionsRoot)) {
    console.error(`[sync-pricing] regions dir not found: ${regionsRoot}`);
    process.exit(1);
  }

  const regions = readdirSync(regionsRoot, { withFileTypes: true })
    .filter((d) => d.isDirectory())
    .map((d) => d.name);

  const out = {
    generatedAt: null,
    source: 'public/pricing/regions (Azure Retail Prices snapshot)',
    hoursPerMonth: HOURS_PER_MONTH,
    currency: 'USD',
    regions: {},
  };

  let totalEntries = 0;
  let newestPriceDate = '';
  for (const region of regions) {
    const dir = resolve(regionsRoot, region);
    const files = readdirSync(dir).filter((f) => f.endsWith('.json'));
    const regionMap = {};
    for (const file of files) {
      const stem = basename(file, '.json');
      if (SKIP_FILES.has(stem)) continue;
      const distilled = stem === 'microsoft_fabric'
        ? distillFabric(resolve(dir, file))
        : distillFile(resolve(dir, file), stem);
      if (distilled) {
        regionMap[stem] = distilled;
        if (distilled.pricesAsOf && distilled.pricesAsOf > newestPriceDate) {
          newestPriceDate = distilled.pricesAsOf;
        }
        totalEntries++;
      }
    }
    out.regions[region] = regionMap;
  }
  out.generatedAt = newestPriceDate
    ? `${newestPriceDate}T00:00:00.000Z`
    : '1970-01-01T00:00:00.000Z';

  writeFileSync(outPath, JSON.stringify(out, null, 2) + '\n', 'utf8');
  console.log(
    `[sync-pricing] wrote ${outPath} — ${regions.length} regions, ${totalEntries} service entries`,
  );
}

main();
