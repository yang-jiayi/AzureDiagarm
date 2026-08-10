// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

import assert from 'node:assert/strict';
import test from 'node:test';
import { readFileSync } from 'node:fs';
import { fileURLToPath } from 'node:url';
import { dirname, join } from 'node:path';
import { compactPricingData } from '../scripts/prep-pricing-data.mjs';
import {
  expandPricingData,
  filterPricingItems,
  parsePricingTiers,
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
