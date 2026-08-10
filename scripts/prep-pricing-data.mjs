// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

/**
 * Pricing data preparation / compaction.
 *
 * The Azure Retail Prices dumps under `public/pricing/regions/**` are fetched
 * verbatim from the API and carry ~11 fields per meter that the runtime never
 * reads, plus redundant values (serviceName/type repeated on every item,
 * unitPrice duplicating retailPrice, armSkuName duplicating skuName). This
 * module compacts each file *in place* so the browser downloads far less data,
 * then the runtime `expandPricingData()` in regionalPricingService.ts restores
 * exactly the fields the pricing parser consumes.
 *
 * The transform is loss-less with respect to the parsed pricing tiers: only
 * fields that `filterPricingItems` / `parsePricingTiers` / the Fabric parser
 * never read are dropped, and every dropped-but-conditionally-read field is
 * restored on expand (see tests/pricing-prep.test.ts).
 *
 * Usage:
 *   node scripts/prep-pricing-data.mjs [--dir <regionsDir>] [--check]
 */

import { readFileSync, writeFileSync, readdirSync, statSync } from 'node:fs';
import { fileURLToPath } from 'node:url';
import { dirname, join, basename } from 'node:path';

// Fields present on the raw Retail-Prices meters that the runtime never reads.
// Dropping them is safe; they are NOT restored on expand.
const DROPPED_ITEM_FIELDS = [
  'currencyCode',
  'tierMinimumUnits',
  'armRegionName',
  'location',
  'effectiveStartDate',
  'meterId',
  'productId',
  'skuId',
  'serviceId',
  'serviceFamily',
  'isPrimaryMeterRegion',
];

// Top-level response fields the runtime never reads (only BillingCurrency +
// Items are used, plus the hoisted ServiceName we add).
const DROPPED_TOP_FIELDS = ['CustomerEntityId', 'CustomerEntityType', 'NextPageLink', 'Count'];

// Files whose items carry a productName the runtime actually reads
// (App Service tier filtering + Foundry AI productName matching).
const KEEP_PRODUCT_FILES = new Set([
  'azure_app_service',
  'foundry_models',
  'foundry_tools',
]);

/**
 * Compact a raw Retail-Prices response object.
 *
 * @param {any} data Parsed raw file: { BillingCurrency, Items: [...] , ... }
 * @param {{ keepProductName?: boolean }} [options]
 * @returns {any} Compacted object safe to JSON.stringify.
 */
export function compactPricingData(data, options = {}) {
  const keepProductName = options.keepProductName === true;
  const items = Array.isArray(data?.Items) ? data.Items : [];

  // Hoist serviceName to the top level only when every item agrees; otherwise
  // keep it per-item so nothing is lost.
  let commonServiceName = null;
  let hoistServiceName = items.length > 0;
  for (const item of items) {
    const name = item?.serviceName ?? null;
    if (commonServiceName === null) {
      commonServiceName = name;
    } else if (commonServiceName !== name) {
      hoistServiceName = false;
      break;
    }
  }

  // Hoist the constant `type` (always "Consumption" for these dumps) the same
  // way — drop it per-item only when every item shares it.
  let hoistType = items.length > 0;
  for (const item of items) {
    if (item?.type !== 'Consumption') { hoistType = false; break; }
  }

  const out = { BillingCurrency: data?.BillingCurrency ?? 'USD' };
  if (hoistServiceName && commonServiceName != null) {
    out.ServiceName = commonServiceName;
  }

  // `effectiveStartDate` is dropped per item, but the newest one is the file's
  // data vintage and is read by mcp-server/scripts/sync-pricing-data.mjs. Hoist
  // it once (10 bytes per file) instead of losing it.
  let newestStartDate = '';
  for (const item of items) {
    const date = typeof item?.effectiveStartDate === 'string' ? item.effectiveStartDate : '';
    if (date > newestStartDate) newestStartDate = date;
  }
  if (newestStartDate) {
    out.PricesAsOf = newestStartDate.slice(0, 10);
  } else if (typeof data?.PricesAsOf === 'string') {
    // Already-compacted input: keep the vintage through a re-compaction.
    out.PricesAsOf = data.PricesAsOf;
  }

  out.Items = items.map((item) => {
    const next = {};
    for (const [key, value] of Object.entries(item)) {
      if (DROPPED_ITEM_FIELDS.includes(key)) continue;
      if (key === 'serviceName' && hoistServiceName) continue;
      if (key === 'type' && hoistType) continue;
      // armSkuName is only read when skuName is absent, so drop the duplicate.
      if (key === 'armSkuName' && item.skuName) continue;
      // unitPrice is only read when retailPrice is falsy.
      if (key === 'unitPrice' && item.retailPrice) continue;
      // productName is only read for App Service + Foundry files.
      if (key === 'productName' && !keepProductName) continue;
      if (key === 'savingsPlan' && Array.isArray(value)) {
        next.savingsPlan = value.map((plan) => {
          const p = {};
          for (const [pk, pv] of Object.entries(plan)) {
            if (pk === 'unitPrice' && plan.retailPrice) continue;
            p[pk] = pv;
          }
          return p;
        });
        continue;
      }
      next[key] = value;
    }
    return next;
  });

  // Preserve BillingCurrency ordering; explicitly ignore dropped top fields.
  void DROPPED_TOP_FIELDS;
  return out;
}

/**
 * Restore a compacted response to the shape the pricing parser expects. Only
 * the conditionally-read fields are reconstructed; permanently-dropped fields
 * stay absent because nothing reads them.
 *
 * @param {any} compact Compacted object from compactPricingData.
 * @returns {{ BillingCurrency: string, Items: any[] }}
 */
export function expandPricingData(compact) {
  const serviceName = compact?.ServiceName;
  const items = Array.isArray(compact?.Items) ? compact.Items : [];
  return {
    BillingCurrency: compact?.BillingCurrency ?? 'USD',
    Items: items.map((item) => {
      const next = { ...item };
      if (next.type === undefined) next.type = 'Consumption';
      if (next.serviceName === undefined && serviceName !== undefined) {
        next.serviceName = serviceName;
      }
      if (next.unitPrice === undefined) next.unitPrice = next.retailPrice;
      if (next.armSkuName === undefined) next.armSkuName = '';
      if (next.productName === undefined) next.productName = '';
      if (Array.isArray(next.savingsPlan)) {
        next.savingsPlan = next.savingsPlan.map((plan) => (
          plan.unitPrice === undefined ? { ...plan, unitPrice: plan.retailPrice } : plan
        ));
      }
      return next;
    }),
  };
}

// ── CLI ─────────────────────────────────────────────────────────────────────

function* walkJsonFiles(dir) {
  for (const entry of readdirSync(dir)) {
    const full = join(dir, entry);
    const st = statSync(full);
    if (st.isDirectory()) {
      yield* walkJsonFiles(full);
    } else if (entry.endsWith('.json')) {
      yield full;
    }
  }
}

function runCli() {
  const args = process.argv.slice(2);
  const here = dirname(fileURLToPath(import.meta.url));
  let regionsDir = join(here, '..', 'public', 'pricing', 'regions');
  const checkOnly = args.includes('--check');
  const dirFlag = args.indexOf('--dir');
  if (dirFlag !== -1 && args[dirFlag + 1]) {
    regionsDir = args[dirFlag + 1];
  }

  let files = 0;
  let bytesBefore = 0;
  let bytesAfter = 0;
  for (const file of walkJsonFiles(regionsDir)) {
    const rawText = readFileSync(file, 'utf8');
    const data = JSON.parse(rawText);
    const fileKey = basename(file, '.json');
    const keepProductName = KEEP_PRODUCT_FILES.has(fileKey);
    const compact = compactPricingData(data, { keepProductName });
    const compactText = JSON.stringify(compact);
    files += 1;
    bytesBefore += Buffer.byteLength(rawText);
    bytesAfter += Buffer.byteLength(compactText);
    if (!checkOnly) {
      writeFileSync(file, compactText);
    }
  }

  const mb = (n) => (n / (1024 * 1024)).toFixed(2);
  const ratio = bytesAfter > 0 ? (bytesBefore / bytesAfter).toFixed(2) : 'n/a';
  console.log(
    `${checkOnly ? '[check] ' : ''}Pricing prep: ${files} files, ` +
    `${mb(bytesBefore)} MB -> ${mb(bytesAfter)} MB (${ratio}x smaller)`
  );
}

// Only run the CLI when executed directly, not when imported by tests.
if (process.argv[1] && fileURLToPath(import.meta.url) === process.argv[1]) {
  runCli();
}
