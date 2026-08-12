// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

import assert from 'node:assert/strict';
import test from 'node:test';
import { readFileSync, readdirSync } from 'node:fs';
import { fileURLToPath } from 'node:url';
import { dirname, join } from 'node:path';
import { compactPricingData } from '../scripts/prep-pricing-data.mjs';
import {
  expandPricingData,
  filterPricingItems,
  getMeterVintage,
  oldestMeterVintage,
  parsePricingTiers,
  recordMeterVintage,
  resetMeterVintages,
} from '../src/services/regionalPricingService';
import type { AzureRetailPrice } from '../src/types/pricing';

const here = dirname(fileURLToPath(import.meta.url));
const REGIONS_DIR = join(here, '..', 'public', 'pricing', 'regions');

// Build a full raw Retail-Prices meter with every field the API emits, so the
// compaction has real fields to drop and the round-trip is realistic.
function rawItem(overrides: Partial<AzureRetailPrice>): AzureRetailPrice {
  return {
    currencyCode: 'USD',
    tierMinimumUnits: 0,
    retailPrice: 0.1,
    unitPrice: 0.1,
    armRegionName: 'japaneast',
    location: 'JA East',
    effectiveStartDate: '2021-11-01T00:00:00Z',
    meterId: '00000000-0000-0000-0000-000000000000',
    meterName: 'Sample Meter',
    productId: 'DZHXXXXXXXXX',
    skuId: 'DZHXXXXXXXXX/0001',
    productName: 'Sample Product',
    skuName: 'Standard_Sample',
    serviceName: 'Virtual Machines',
    serviceId: 'DZHYYYYYYYYY',
    serviceFamily: 'Compute',
    unitOfMeasure: '1 Hour',
    type: 'Consumption',
    isPrimaryMeterRegion: true,
    armSkuName: 'Standard_Sample',
    ...overrides,
  };
}

// Compaction → expansion must preserve the parsed pricing tiers exactly.
function assertTiersPreserved(
  raw: { BillingCurrency: string; Items: AzureRetailPrice[] },
  serviceName: string,
  keepProductName: boolean,
) {
  const rawTiers = parsePricingTiers(filterPricingItems(raw.Items, serviceName), serviceName);
  const compact = compactPricingData(raw, { keepProductName });
  const expanded = expandPricingData(compact as never);
  const expandedTiers = parsePricingTiers(
    filterPricingItems(expanded.Items, serviceName),
    serviceName,
  );
  assert.deepEqual(expandedTiers, rawTiers);
  return { rawTiers, compact };
}

test('compaction preserves Virtual Machines tiers (skuName/armSkuName/savingsPlan/zero-retail)', () => {
  const raw = {
    BillingCurrency: 'USD',
    Items: [
      rawItem({
        meterName: 'D2 v5',
        skuName: 'Standard_D2_v5',
        armSkuName: 'Standard_D2_v5',
        retailPrice: 0.096,
        unitPrice: 0.096,
        savingsPlan: [
          { term: '1 Year', retailPrice: 0.06, unitPrice: 0.06 },
          { term: '3 Years', retailPrice: 0.04, unitPrice: 0.04 },
        ],
      }),
      // retailPrice 0 → unitPrice must survive compaction and drive the tier.
      rawItem({
        meterName: 'D4 v5 Spot',
        skuName: 'Standard_D4_v5 Spot',
        retailPrice: 0,
        unitPrice: 0.031,
      }),
      // skuName empty → armSkuName must survive and be used as the tier key.
      rawItem({
        meterName: 'F8s v2',
        skuName: '',
        armSkuName: 'Standard_F8s_v2',
        retailPrice: 0.338,
        unitPrice: 0.338,
      }),
      // Duplicate SKU at higher price → parser keeps the cheaper one.
      rawItem({
        meterName: 'D2 v5 Dup',
        skuName: 'Standard_D2_v5',
        armSkuName: 'Standard_D2_v5',
        retailPrice: 0.15,
        unitPrice: 0.15,
      }),
    ],
  };

  const { rawTiers, compact } = assertTiersPreserved(raw, 'Virtual Machines', false);
  assert.ok(rawTiers.length >= 3, 'expected VM tiers to be parsed');
  // serviceName/type were hoisted; verify per-item copies were dropped.
  assert.equal((compact as { ServiceName?: string }).ServiceName, 'Virtual Machines');
  assert.equal((compact as { Items: Array<Record<string, unknown>> }).Items[0].serviceName, undefined);
  assert.equal((compact as { Items: Array<Record<string, unknown>> }).Items[0].type, undefined);
});

test('compaction preserves Azure App Service tiers and keeps productName', () => {
  const raw = {
    BillingCurrency: 'USD',
    Items: [
      rawItem({
        serviceName: 'Azure App Service',
        productName: 'Azure App Service Standard Plan',
        skuName: 'S1',
        armSkuName: 'S1',
        meterName: 'S1',
        unitOfMeasure: '1 Hour',
        retailPrice: 0.1,
        unitPrice: 0.1,
      }),
      rawItem({
        serviceName: 'Azure App Service',
        productName: 'Azure App Service Premium v3 Plan - Linux',
        skuName: 'P1V3',
        armSkuName: 'P1V3',
        meterName: 'P1 v3',
        retailPrice: 0.2,
        unitPrice: 0.2,
      }),
      // Non-plan product (SSL) must be filtered out both before and after.
      rawItem({
        serviceName: 'Azure App Service',
        productName: 'Azure App Service SSL Connections',
        skuName: 'SNI SSL',
        armSkuName: '',
        meterName: 'SNI SSL',
        retailPrice: 0,
        unitPrice: 9,
      }),
    ],
  };

  const { rawTiers, compact } = assertTiersPreserved(raw, 'Azure App Service', true);
  assert.equal(rawTiers.length, 2, 'SSL add-on must be excluded, two plan tiers kept');
  assert.equal(
    (compact as { Items: Array<Record<string, unknown>> }).Items[0].productName,
    'Azure App Service Standard Plan',
    'productName must be retained for App Service files',
  );
});

test('real compacted regional files expand and parse to non-empty VM tiers', () => {
  // Smoke test on the actual public data: compacted files must remain usable.
  const regions = ['japaneast', 'eastus2', 'westeurope'];
  for (const region of regions) {
    const file = join(REGIONS_DIR, region, 'virtual_machines.json');
    const compact = JSON.parse(readFileSync(file, 'utf8'));
    const expanded = expandPricingData(compact);
    const tiers = parsePricingTiers(
      filterPricingItems(expanded.Items, 'Virtual Machines'),
      'Virtual Machines',
    );
    assert.ok(tiers.length > 0, `expected VM tiers for ${region}`);
    // Re-compacting the expanded data and expanding again must be idempotent.
    const reExpanded = expandPricingData(compactPricingData(expanded, { keepProductName: false }) as never);
    const reTiers = parsePricingTiers(
      filterPricingItems(reExpanded.Items, 'Virtual Machines'),
      'Virtual Machines',
    );
    assert.deepEqual(reTiers, tiers, `round-trip must be idempotent for ${region}`);
  }
});

// ---------------------------------------------------------------------------
// Meter vintage. `PricesAsOf` is the newest meter date in a file, which is not
// the date the file was downloaded: Azure restates a meter only when it
// changes, so a refresh run today still returns 2018 meters for services it has
// not repriced. The shipped corpus spans 2018-02 to 2026-07, so quoting the
// download date alone tells a customer a deck is current when part of it is
// eight years old.
// ---------------------------------------------------------------------------

test('expansion carries the meter vintage off the file', () => {
  const compact = { BillingCurrency: 'USD', ServiceName: 'Virtual Machines', PricesAsOf: '2018-02-01', Items: [] };
  assert.equal(expandPricingData(compact as never).pricesAsOf, '2018-02-01');
});

test('a malformed or missing vintage is dropped rather than reported', () => {
  // A bad date is worse than no date: it would be printed on a customer deck
  // verbatim, and it would sort against real dates in the oldest-wins compare.
  for (const bad of [undefined, '', 'last Tuesday', '2018-2-1', '2018-02-01T00:00:00Z', 42]) {
    const compact = { BillingCurrency: 'USD', Items: [], PricesAsOf: bad };
    assert.equal(expandPricingData(compact as never).pricesAsOf, undefined, `rejected: ${String(bad)}`);
  }
});

test('the registry keeps the oldest vintage seen for a service', () => {
  resetMeterVintages();
  assert.equal(getMeterVintage('Virtual Machines'), undefined);
  recordMeterVintage('Virtual Machines', '2024-10-01');
  assert.equal(getMeterVintage('Virtual Machines'), '2024-10-01');
  // The same service is loaded once per region, and a number is only as
  // trustworthy as the stalest meter behind it.
  recordMeterVintage('Virtual Machines', '2018-02-01');
  assert.equal(getMeterVintage('Virtual Machines'), '2018-02-01');
  recordMeterVintage('Virtual Machines', '2026-07-01');
  assert.equal(getMeterVintage('Virtual Machines'), '2018-02-01', 'a newer region must not refresh an older meter');
});

test('the registry refuses a vintage it cannot trust', () => {
  // The guard in `recordMeterVintage` is not redundant with the one in
  // `expandPricingData`: this is an exported entry point, and the values it
  // stores are compared with `<` and printed verbatim on a customer deck. A
  // string that is not an ISO date would sort arbitrarily against real ones and
  // could win the oldest-wins compare on every estimate.
  resetMeterVintages();
  for (const bad of [undefined, '', 'last Tuesday', '2018-2-1', '2018-02-01T00:00:00Z']) {
    recordMeterVintage('Virtual Machines', bad as string | undefined);
    assert.equal(getMeterVintage('Virtual Machines'), undefined, `accepted: ${String(bad)}`);
  }
  recordMeterVintage('Virtual Machines', '2018-02-01');
  assert.equal(getMeterVintage('Virtual Machines'), '2018-02-01');
});

test('the oldest vintage is taken across only the services asked about', () => {
  resetMeterVintages();
  recordMeterVintage('Virtual Machines', '2018-02-01');
  recordMeterVintage('Storage', '2024-10-01');
  recordMeterVintage('SQL Database', '2026-07-01');
  assert.equal(oldestMeterVintage(['Storage', 'SQL Database']), '2024-10-01');
  assert.equal(oldestMeterVintage(['Virtual Machines', 'Storage']), '2018-02-01');
  assert.equal(oldestMeterVintage(['Unpriced Thing']), undefined);
  assert.equal(oldestMeterVintage([]), undefined);
});

test('every shipped pricing file states a vintage that parses', () => {
  // The whole feature rests on this field being present and well formed in the
  // data as shipped; a silent regression in the prep script would otherwise
  // just make the warning quietly stop appearing.
  let checked = 0;
  for (const region of ['japaneast', 'eastus2', 'westeurope']) {
    for (const entry of readdirSync(join(REGIONS_DIR, region))) {
      if (!entry.endsWith('.json')) continue;
      const compact = JSON.parse(readFileSync(join(REGIONS_DIR, region, entry), 'utf8'));
      if (!Array.isArray(compact.Items) || compact.Items.length === 0) continue;
      checked += 1;
      const asOf = expandPricingData(compact).pricesAsOf;
      assert.ok(asOf, `${region}/${entry} ships meters but states no vintage`);
      assert.ok(!Number.isNaN(Date.parse(asOf!)), `${region}/${entry} states an unparseable vintage ${asOf}`);
    }
  }
  assert.ok(checked > 50, `expected a real corpus, checked only ${checked} files`);
});

test('the estimate reports how long its longest-standing price has held', async () => {
  // Azure's retail API returns current prices, so a 2018 meter is genuinely
  // today's price — it is one Azure has not repriced since. That is not a
  // staleness warning, it is the answer to "how firm is this number?", and it
  // is only worth putting on a slide when the price has actually held for a
  // while: every meter predates the download by some margin, so reporting any
  // gap at all would put a second date on every deck that never means anything.
  const { calculateCostBreakdown } = await import('../src/services/costEstimationService');
  const { PRICING_DATA_AS_OF } = await import('../src/data/azurePricing');
  const node = (id: string, serviceName: string, extra: Record<string, unknown> = {}) => ({
    id,
    position: { x: 0, y: 0 },
    data: {
      label: serviceName,
      serviceName,
      category: 'compute',
      pricing: { estimatedCost: 10, quantity: 1, tier: 'Standard', region: 'japaneast', ...extra },
    },
  });

  resetMeterVintages();
  recordMeterVintage('Virtual Machines', '2018-02-01');
  recordMeterVintage('Storage', '2026-07-01');

  const stable = calculateCostBreakdown([node('a', 'Virtual Machines'), node('b', 'Storage')] as never);
  assert.equal(stable.pricesAsOf, PRICING_DATA_AS_OF, 'the refresh date still says when the data was fetched');
  assert.equal(stable.oldestMeterAsOf, '2018-02-01', 'the oldest of the two, not the newest');

  // A fortnight is not news. Reporting it would make the line permanent
  // furniture and therefore invisible when it does matter.
  const recent = calculateCostBreakdown([node('b', 'Storage')] as never);
  assert.equal(recent.oldestMeterAsOf, undefined);

  // A custom price is a number the user typed; no Azure meter stands behind it,
  // so it must not drag the reported vintage backwards.
  resetMeterVintages();
  recordMeterVintage('Virtual Machines', '2018-02-01');
  const typed = calculateCostBreakdown([node('a', 'Virtual Machines', { isCustom: true })] as never);
  assert.equal(typed.oldestMeterAsOf, undefined);

  // The bar is a year, so a price that moved eleven months ago is still recent.
  const elevenMonths = new Date(Date.parse(`${PRICING_DATA_AS_OF}T00:00:00Z`) - 334 * 86_400_000)
    .toISOString().slice(0, 10);
  resetMeterVintages();
  recordMeterVintage('Virtual Machines', elevenMonths);
  assert.equal(calculateCostBreakdown([node('a', 'Virtual Machines')] as never).oldestMeterAsOf, undefined);
});
