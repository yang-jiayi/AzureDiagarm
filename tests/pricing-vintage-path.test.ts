// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

/**
 * The vintage feature driven through the real pricing path, against the real
 * shipped data.
 *
 * The hand-built-node tests next door check what `calculateCostBreakdown` does
 * with a pricing config. They cannot check the half of the feature that was
 * actually broken, because that bug was in how the config gets its date: the
 * vintage was recorded under the *mapped* Azure product name and read back
 * under the raw canvas name, so `Storage Accounts`, `Function Apps`,
 * `Machine Learning` and about half the catalogue silently reported nothing —
 * and the ones it lost included the oldest meters in the corpus. Nothing short
 * of running the real fetch → parse → initialise chain would have noticed.
 */

import assert from 'node:assert/strict';
import test, { before } from 'node:test';
import { readFileSync, existsSync } from 'node:fs';
import { fileURLToPath } from 'node:url';
import { dirname, join } from 'node:path';

const here = dirname(fileURLToPath(import.meta.url));
const REGIONS_DIR = join(here, '..', 'public', 'pricing', 'regions');

/** Serve `public/pricing/regions/**` the way the dev server does. */
function serveFromDisk(): void {
  (globalThis as { fetch?: unknown }).fetch = async (input: unknown) => {
    const url = String(input);
    const match = /\/pricing\/regions\/([^/]+)\/([^/]+)\.json/.exec(url);
    const file = match ? join(REGIONS_DIR, match[1], `${match[2]}.json`) : '';
    if (!file || !existsSync(file)) {
      return { ok: false, status: 404, json: async () => ({}) } as unknown as Response;
    }
    const body = JSON.parse(readFileSync(file, 'utf8'));
    return { ok: true, status: 200, json: async () => body } as unknown as Response;
  };
}

before(serveFromDisk);

test('a canvas name that has to be mapped still reports its vintage', async () => {
  // Every one of these differs from the Azure retail product name it is priced
  // from, and every one prices from a real meter — which is the combination the
  // old registry answered "nothing" for. `Azure Key Vault` and `Azure IoT Hub`
  // are the point of the whole feature: both were last repriced in 2022 and
  // both are on almost every diagram.
  const { initializeNodePricing } = await import('../src/services/costEstimationService');
  const { getAzureServiceName } = await import('../src/data/azurePricing');
  const mapped = ['Azure Key Vault', 'Azure IoT Hub', 'Azure SQL Database', 'AKS', 'Container Instance', 'Express Route'];
  for (const serviceType of mapped) {
    assert.notEqual(
      getAzureServiceName(serviceType),
      serviceType,
      `${serviceType} no longer needs mapping, so it no longer tests anything`,
    );
    const pricing = await initializeNodePricing(serviceType, 'eastus2');
    assert.ok(pricing, `${serviceType} must price from the shipped data`);
    assert.ok(pricing.meterAsOf, `${serviceType} priced from a meter but reported no vintage`);
    assert.match(pricing.meterAsOf!, /^\d{4}-\d{2}-\d{2}$/, `${serviceType} carries a malformed vintage`);
  }
});

test('the long-stable services the feature exists for reach the deck', async () => {
  // The end-to-end claim, against the real corpus: a diagram of ordinary
  // building blocks says out loud that its oldest price has held since 2022.
  const { initializeNodePricing, calculateCostBreakdown } = await import('../src/services/costEstimationService');
  const node = async (id: string, serviceType: string) => ({
    id,
    position: { x: 0, y: 0 },
    data: {
      label: serviceType,
      serviceName: serviceType,
      category: 'compute',
      pricing: await initializeNodePricing(serviceType, 'eastus2'),
    },
  });
  const breakdown = calculateCostBreakdown([
    await node('a', 'Azure Key Vault'),
    await node('b', 'Virtual Machine'),
    await node('c', 'Azure SQL Database'),
  ] as never);
  assert.equal(breakdown.oldestMeterAsOf, '2022-03-01', 'the oldest of the three, and past the two-year bar');
});

test('the vintage the node carries is the one in its own region', async () => {
  // 30 of the 48 priced services have a region-dependent vintage, spreads
  // reaching six years. The old registry took the oldest across every region
  // ever loaded, so pricing in one region and then another printed the first
  // one's date beside the second one's money. Loading both here in one process
  // is the exact sequence that used to poison it.
  const { initializeNodePricing } = await import('../src/services/costEstimationService');
  const east = await initializeNodePricing('Container Instances', 'eastus2');
  const aus = await initializeNodePricing('Container Instances', 'australiaeast');
  assert.equal(east?.region, 'eastus2');
  assert.equal(aus?.region, 'australiaeast');
  assert.equal(east?.meterAsOf, '2024-09-01');
  assert.equal(aus?.meterAsOf, '2026-07-01', 'australiaeast repriced 14 days before the download');
  // And the earlier region's answer is unchanged by the later one having run.
  const again = await initializeNodePricing('Container Instances', 'eastus2');
  assert.equal(again?.meterAsOf, '2024-09-01', 'a later region changed an earlier one\u2019s answer');
});

test('a node priced from the static fallback claims no meter', async () => {
  // `initializeNodePricing` falls back to a hand-maintained constant whenever
  // the parsed tier is $0 or no tier parses at all. The number that reaches the
  // slide then came from no Azure meter, so "Azure last changed this price on
  // X" is false in both of its clauses. The old code recorded the vintage when
  // the *file* loaded, so it attested to it anyway.
  const { initializeNodePricing } = await import('../src/services/costEstimationService');
  // All five of these are priced from the hand-maintained table because their
  // parsed retail tier is $0 — the reviewer's own list, and the reason their
  // costs looked convincing ($14.60, $159.35, $50.00) while standing behind no
  // meter at all.
  for (const serviceType of ['Storage Accounts', 'Function Apps', 'Machine Learning', 'Event Hubs', 'API Management Services']) {
    const pricing = await initializeNodePricing(serviceType, 'japaneast');
    assert.ok(pricing, `${serviceType} must still price`);
    assert.ok(pricing.estimatedCost > 0, `${serviceType} must still carry a usable number`);
    assert.equal(pricing.isUsageBased, true, `${serviceType} no longer takes the fallback path`);
    assert.equal(
      pricing.meterAsOf,
      undefined,
      `${serviceType} is fallback-priced but claimed Azure set its price on a date`,
    );
  }
});

test('the vintage survives a round trip through saved node data', async () => {
  // The whole point of moving it off session state: a restored diagram exports
  // the same provenance as the session that saved it.
  const { initializeNodePricing, calculateCostBreakdown } = await import('../src/services/costEstimationService');
  const pricing = await initializeNodePricing('Virtual Machines', 'eastus2');
  assert.ok(pricing, 'Virtual Machines must price from the shipped data');
  const saved = JSON.parse(JSON.stringify({
    id: 'a',
    position: { x: 0, y: 0 },
    data: { label: 'Virtual Machines', serviceName: 'Virtual Machines', category: 'compute', pricing },
  }));
  const breakdown = calculateCostBreakdown([saved] as never);
  assert.equal(breakdown.byService.length, 1);
  // Equal whether or not the bar lets it through: what matters is that the
  // restored node answers identically to the live one.
  const live = calculateCostBreakdown([{
    id: 'a',
    position: { x: 0, y: 0 },
    data: { label: 'Virtual Machines', serviceName: 'Virtual Machines', category: 'compute', pricing },
  }] as never);
  assert.equal(breakdown.oldestMeterAsOf, live.oldestMeterAsOf);
});
