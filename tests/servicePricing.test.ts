import test from 'node:test';
import assert from 'node:assert/strict';
import { readFile } from 'node:fs/promises';
import type { Node } from 'reactflow';
import type { AzureRetailPrice, NodePricingConfig, PricingTier } from '../src/types/pricing';
import { parsePricingTiers } from '../src/services/pricingMeters';
import {
  buildServiceInspectorData, getFallbackPricingTiers, getPricingProvenance,
  isPricingProvenance, isPricingUsage, pricingFromTier, selectPricingTier, withUsageAmount,
} from '../src/services/pricingConfiguration';
import {
  calculateCostBreakdown, getCostSummaryMarkdown, getCostSummaryText, exportCostBreakdownCSV, initializeNodePricing,
  refreshAllNodePricing, setCustomPricing, updateNodePricing,
} from '../src/services/costEstimationService';
import { calculateMonthlyCost } from '../src/services/azurePricingService';

const meter = (overrides: Partial<AzureRetailPrice> = {}): AzureRetailPrice => ({
  skuName: 'B1', meterName: 'B1 instance', meterId: 'meter-1',
  retailPrice: 0.1, unitPrice: 0.2, unitOfMeasure: '1 Hour',
  effectiveStartDate: '2026-07-01', tierMinimumUnits: 0, ...overrides,
} as AzureRetailPrice);
const instanceTier = (): PricingTier => parsePricingTiers([meter()])[0];
const config = (): NodePricingConfig => pricingFromTier(instanceTier(), 1, 'eastus2');
const node = (pricing?: NodePricingConfig, serviceName = 'App Service'): Node => ({
  id: 'node-1', type: 'azureNode', position: { x: 0, y: 0 },
  data: { label: 'Customer-facing label', serviceName, pricing },
});

test('official meters carry effective date, units, source and supported monthly assumptions', () => {
  const tier = instanceTier();
  assert.equal(tier.monthlyPrice, 73);
  assert.equal(tier.provenance?.kind, 'official-meter');
  assert.equal(tier.provenance?.asOf, '2026-07-01');
  assert.equal(tier.provenance?.meterId, 'meter-1');
  assert.deepEqual(tier.provenance?.assumptions[0], { label: 'Monthly operation', value: 730, unit: 'hours/month' });
  assert.equal(parsePricingTiers([meter({ unitOfMeasure: '1/Month', retailPrice: 12 })])[0].monthlyPrice, 12);
  assert.equal(parsePricingTiers([meter({ unitOfMeasure: '1/Day', retailPrice: 2 })])[0].monthlyPrice, 60);
  assert.equal(parsePricingTiers([meter({ unitOfMeasure: '1/Year', retailPrice: 120 })])[0].monthlyPrice, 10);
});

test('unknown usage has no invented hourly/token/GB consumption; explicit meter usage works', () => {
  for (const unit of ['1 GB/Month', '1K', '1M Tokens', '1 GB']) {
    const tier = parsePricingTiers([meter({ unitOfMeasure: unit, retailPrice: 0.5 })])[0];
    assert.equal(tier.monthlyPrice, null);
    assert.equal(tier.provenance?.kind, 'unpriced');
    assert.equal(tier.usage?.amount, null);
    const priced = pricingFromTier(tier, 3, 'eastus2', 10);
    assert.equal(priced.estimatedCost, 5);
    assert.equal(priced.provenance?.kind, 'usage-estimate');
    assert.equal(calculateCostBreakdown([node(priced)]).totalMonthlyCost, 15);
    assert.equal(withUsageAmount(priced, null).estimatedCost, null);
  }
});

test('true zero retail rates are preserved; nonfinite or negative rates are rejected', () => {
  assert.equal(parsePricingTiers([meter({ retailPrice: 0, unitPrice: 9 })])[0].monthlyPrice, 0);
  assert.equal(parsePricingTiers([meter({ retailPrice: Infinity })]).length, 0);
  assert.equal(parsePricingTiers([meter({ retailPrice: -1 })]).length, 0);
});

test('tiered usage meters do not extrapolate the zero-price band into a free service', () => {
  const tiers = parsePricingTiers([
    meter({ unitOfMeasure: '1 GB', retailPrice: 0 }),
    meter({ unitOfMeasure: '1 GB', retailPrice: 2, tierMinimumUnits: 5 }),
  ]);
  assert.equal(tiers[0].monthlyPrice, null);
  assert.equal(tiers[0].usage, undefined);
  assert.match(tiers[0].provenance!.note!, /Tiered meter/);
  const compact = parsePricingTiers([
    meter({ unitOfMeasure: '1 GB', retailPrice: 0, tierMinimumUnits: undefined }),
    meter({ unitOfMeasure: '1 GB', retailPrice: 2, tierMinimumUnits: undefined }),
  ]);
  assert.equal(compact[0].monthlyPrice, null);
  assert.equal(compact[0].usage, undefined, 'lost band thresholds cannot become a linear free meter');
});

test('fallback consumption retains usage classification when switching inspector tiers', () => {
  const tier = getFallbackPricingTiers('Azure OpenAI', 'eastus2')[0];
  assert.ok(tier);
  assert.equal(pricingFromTier(tier, 1, 'eastus2').isUsageBased, true);
});

test('quantity updates preserve PER UNIT cost and aggregate quantity exactly once', async () => {
  const updated = await updateNodePricing('App Service', config(), undefined, 4);
  assert.equal(updated.estimatedCost, 73);
  assert.equal(updated.quantity, 4);
  assert.equal(calculateCostBreakdown([node(updated)]).totalMonthlyCost, 292);
  const fallback = await updateNodePricing('App Service', updated, 'Premium', 3, 'unbundled-region');
  assert.equal(fallback.estimatedCost, 146);
  assert.equal(calculateCostBreakdown([node(fallback)]).totalMonthlyCost, 438);
  assert.equal(fallback.provenance?.kind, 'fallback-estimate');
});

test('production quantities up to 100000 stay priced in totals and regional refresh', async () => {
  for (const quantity of [10_001, 100_000]) {
    const pricing = await updateNodePricing('App Service', config(), undefined, quantity);
    const breakdown = calculateCostBreakdown([node(pricing)]);
    assert.equal(breakdown.estimateCompleteness, 'complete');
    assert.equal(breakdown.totalMonthlyCost, 73 * quantity);
    assert.equal(breakdown.byService[0].quantity, quantity);
    assert.deepEqual(breakdown.unpricedServices, []);

    const [refreshed] = await refreshAllNodePricing([node(setCustomPricing(pricing, 25))], 'japaneast');
    assert.equal(refreshed.data.pricing.quantity, quantity);
    assert.equal(calculateCostBreakdown([refreshed]).totalMonthlyCost, 25 * quantity);
  }
  const invalid = calculateCostBreakdown([node({ ...config(), quantity: 100_001 })]);
  assert.equal(invalid.estimateCompleteness, 'unpriced');
  assert.equal(invalid.byService.length, 0);
});

test('unavailable imported pricing keeps a valid 100000-unit quantity', async () => {
  const original = node({ ...config(), quantity: 100_000 });
  original.data.serviceName = { legacy: 'App Service' };
  const [refreshed] = await refreshAllNodePricing([original], 'japaneast');
  assert.equal(refreshed.data.pricing.quantity, 100_000);
  assert.equal(refreshed.data.pricing.estimatedCost, null);
  assert.equal(refreshed.data.pricing.provenance.kind, 'unpriced');
  assert.match(refreshed.data.pricing.provenance.note, /Imported pricing is incomplete or invalid/);
  assert.equal(original.data.pricing.estimatedCost, 73);
});

test('invalid quantities, nonfinite estimates and unavailable SKUs are rejected', async () => {
  for (const value of [0, -1, 1.5, 100001, Infinity, NaN]) {
    await assert.rejects(updateNodePricing('App Service', config(), undefined, value), /Quantity/);
  }
  for (const value of [-1, NaN, Infinity, 1e13]) {
    assert.throws(() => setCustomPricing(config(), value), /Amount/);
  }
  await assert.rejects(updateNodePricing('App Service', config(), 'Invented SKU'), /SKU/);
  assert.equal(selectPricingTier([instanceTier()], 'unknown'), undefined);
});

test('custom prices are per-unit, survive region changes and are never automatically discounted', async () => {
  const custom = setCustomPricing({ ...config(), quantity: 3 }, 25);
  const updated = await updateNodePricing('App Service', custom, undefined, 4, 'japaneast');
  assert.equal(updated.estimatedCost, 25);
  assert.equal(updated.region, 'japaneast');
  assert.equal(updated.quantity, 4);
  assert.equal(updated.isCustom, true);
  assert.equal(updated.provenance?.kind, 'custom');
  assert.equal(calculateCostBreakdown([node(updated)], undefined, 'reserved1yr').totalMonthlyCost, 100);
});

test('region refresh preserves available tier, quantity and stable service identity after relabeling', async () => {
  const standard = (await initializeNodePricing('App Service', 'unbundled-region'))!;
  const premium = await updateNodePricing('App Service', standard, 'Premium', 3);
  const [updated] = await refreshAllNodePricing([node(premium)], 'japaneast');
  assert.equal(updated.data.pricing.tier, 'Premium');
  assert.equal(updated.data.pricing.quantity, 3);
  assert.equal(updated.data.pricing.estimatedCost, 146 * 1.12);
  assert.equal(updated.data.pricing.provenance.kind, 'fallback-estimate');
});

test('a regional unavailable manual SKU is retained but unpriced, not switched to standard', async () => {
  const old = { ...config(), tier: 'P999', skuName: 'P999', quantity: 2 };
  const updated = await updateNodePricing('App Service', old, undefined, undefined, 'unbundled-region');
  assert.equal(updated.tier, 'P999');
  assert.equal(updated.quantity, 2);
  assert.equal(updated.estimatedCost, null);
  assert.equal(updated.provenance?.kind, 'unpriced');
  const namedLikeFallback = { ...config(), tierId: 'Standard', tier: 'Standard', skuName: 'Standard' };
  const unavailable = await updateNodePricing('App Service', namedLikeFallback, undefined, undefined, 'unbundled-region');
  assert.equal(unavailable.estimatedCost, null, 'fallback Standard is not a substantiated regional Standard SKU');
  assert.equal(unavailable.tierId, 'Standard');
});

test('legacy pricing source and assumptions remain unknown after quantity-only edits', async () => {
  const { provenance: _, ...legacy } = config();
  const updated = await updateNodePricing('App Service', legacy, undefined, 2);
  assert.equal(getPricingProvenance(updated).kind, 'unknown');
  assert.equal(getPricingProvenance(updated).asOf, undefined);
  assert.deepEqual(getPricingProvenance(updated).assumptions, []);
  assert.equal(getPricingProvenance({ ...legacy, provenance: {} as never }).kind, 'unknown');
  assert.equal(isPricingProvenance({ kind: 'official-meter' }), false);
  assert.equal(isPricingUsage({ amount: Infinity, unit: '1 GB', unitPrice: 1 }), false);
  assert.equal(isPricingUsage({ amount: 1, unit: '', unitPrice: 1 }), false);
  assert.throws(() => withUsageAmount({
    ...legacy, usage: { amount: 5, unit: '1 GB', unitPrice: 1 },
  }, 10), /substantiated/);
});

test('capacity ladder preserves F-SKU selection and consumed workloads are not free estimates', async () => {
  const capacity = (await initializeNodePricing('Microsoft Fabric Capacity', 'unbundled-region'))!;
  assert.equal(capacity.skuName, 'F2');
  const larger = await updateNodePricing('Microsoft Fabric Capacity', capacity, 'F64', 2);
  assert.equal(larger.estimatedCost, 8409.60);
  assert.equal(larger.quantity, 2);
  assert.equal(calculateCostBreakdown([node(larger, 'Microsoft Fabric Capacity')]).totalMonthlyCost, 16819.2);
  const workload = (await initializeNodePricing('Lakehouse', 'unbundled-region'))!;
  assert.equal(workload.provenance?.kind, 'capacity');
  assert.equal(workload.estimatedCost, null);
  const breakdown = calculateCostBreakdown([node(workload, 'Lakehouse')]);
  assert.equal(breakdown.capacityServices?.length, 1);
  assert.equal(breakdown.byService.length, 0);
});

test('missing estimates are explicitly excluded and exports mark a partial subtotal', () => {
  const breakdown = calculateCostBreakdown([
    node({ ...config(), estimatedCost: null }),
    { ...node(undefined, 'Unknown service'), id: 'unknown' },
    { ...node(config()), id: 'priced' },
  ]);
  assert.equal(breakdown.totalMonthlyCost, 73);
  assert.equal(breakdown.unpricedServices?.length, 2);
  assert.equal(breakdown.byService.length, 1);
  assert.equal(breakdown.pricedMonthlySubtotal, 73);
  assert.equal(breakdown.estimateCompleteness, 'partial');
  assert.match(getCostSummaryMarkdown(breakdown), /Partial subtotal/);
  assert.deepEqual(getFallbackPricingTiers('Unsupported service', 'eastus2'), []);
});

test('unknown-only subtotal differs from known zero, including partial legacy pricing', () => {
  const unknown = calculateCostBreakdown([
    node({ estimatedCost: null, quantity: 1 } as NodePricingConfig),
    { ...node({ estimatedCost: 73 } as NodePricingConfig), id: 'missing-quantity' },
    { ...node({ quantity: 1 } as NodePricingConfig), id: 'missing-price' },
  ]);
  assert.equal(unknown.estimateCompleteness, 'unpriced');
  assert.equal(unknown.pricedMonthlySubtotal, 0);
  assert.equal(unknown.unpricedServices?.length, 3);
  assert.match(getCostSummaryText(unknown), /Priced Monthly Subtotal/);
  assert.match(getCostSummaryMarkdown(unknown), /No usable total estimate/);
  assert.match(exportCostBreakdownCSV(unknown), /Estimate Completeness,unpriced/);
  const free = calculateCostBreakdown([node({ ...config(), estimatedCost: 0 })]);
  assert.equal(free.estimateCompleteness, 'complete');
  assert.equal(free.byService[0].cost, 0);
  assert.equal(free.unpricedServices?.length, 0);
});

test('valid legacy amounts without tier/unit/provenance metadata remain exportable and source-unknown', () => {
  const legacy = { estimatedCost: 10, quantity: 2 } as NodePricingConfig;
  const breakdown = calculateCostBreakdown([node(legacy)], undefined, 'reserved1yr');
  assert.equal(breakdown.totalMonthlyCost, 20);
  assert.equal(breakdown.byService[0].tier, 'Unspecified');
  assert.equal(getPricingProvenance(legacy).kind, 'unknown');
  assert.doesNotThrow(() => getCostSummaryMarkdown(breakdown));
  assert.doesNotThrow(() => exportCostBreakdownCSV(breakdown));
  assert.equal(getPricingProvenance({ ...legacy, unit: {} as never }).unit, 'USD/unit/month');
});

test('one invalid imported quantity does not abort region refresh or become a guessed price', async () => {
  const invalid = { ...node({ estimatedCost: 10, quantity: Infinity } as NodePricingConfig), id: 'invalid' };
  const custom = { ...node(setCustomPricing({ ...config(), quantity: 3 }, 25)), id: 'custom' };
  const standard = (await initializeNodePricing('App Service', 'unbundled-region'))!;
  const priced = { ...node({ ...standard, quantity: 2 }), id: 'priced' };
  const refreshed = await refreshAllNodePricing([invalid, custom, priced], 'japaneast');
  assert.equal(refreshed[0].data.pricing.estimatedCost, null);
  assert.equal(refreshed[0].data.pricing.provenance.kind, 'unpriced');
  assert.equal(refreshed[0].data.pricing.region, 'japaneast');
  assert.equal(invalid.data.pricing.quantity, Infinity);
  assert.equal(refreshed[1].data.pricing.estimatedCost, 25);
  assert.equal(refreshed[1].data.pricing.quantity, 3);
  assert.equal(refreshed[2].data.pricing.quantity, 2);
  assert.equal(calculateCostBreakdown(refreshed).estimateCompleteness, 'partial');
});

test('shared-capacity workloads without a priced capacity keep the subtotal incomplete', async () => {
  const workload = (await initializeNodePricing('Lakehouse', 'unbundled-region'))!;
  const missing = calculateCostBreakdown([node(workload, 'Lakehouse')]);
  assert.equal(missing.estimateCompleteness, 'unpriced');
  assert.equal(missing.missingCapacityEstimate, true);
  assert.match(getCostSummaryMarkdown(missing), /no shared capacity estimate/);
  const capacity = (await initializeNodePricing('Microsoft Fabric Capacity', 'unbundled-region'))!;
  const complete = calculateCostBreakdown([
    node(workload, 'Lakehouse'),
    { ...node(capacity, 'Microsoft Fabric Capacity'), id: 'capacity' },
  ]);
  assert.equal(complete.estimateCompleteness, 'complete');
  assert.equal(complete.missingCapacityEstimate, false);
});

test('Savings Plan cost is per-unit and usage/custom prices never get fabricated discounts', () => {
  const tier = parsePricingTiers([meter({
    savingsPlan: [{ term: '1 Year', retailPrice: 0.05, unitPrice: 0.05 }],
  })])[0];
  assert.equal(calculateCostBreakdown([node(pricingFromTier(tier, 4, 'eastus2'))], undefined, 'reserved1yr').totalMonthlyCost, 146);
  const consumption = pricingFromTier(parsePricingTiers([meter({ unitOfMeasure: '1K', retailPrice: 1 })])[0], 2, 'eastus2', 10);
  assert.equal(calculateCostBreakdown([node(consumption)], undefined, 'reserved1yr').totalMonthlyCost, 20);
  const legacy = { ...config(), provenance: undefined, reserved1yrCost: 0 };
  assert.equal(calculateCostBreakdown([node(legacy)], undefined, 'reserved1yr').totalMonthlyCost, 73);
  assert.equal(calculateCostBreakdown([node({
    ...consumption, isUsageBased: undefined, reserved1yrCost: 0,
  })], undefined, 'reserved1yr').totalMonthlyCost, 20);
});

test('inspector creates one merged immutable payload, retaining arbitrary diagram data', () => {
  const original = { label: 'App Service', description: 'Old', icon: 'icon.svg', groupId: 'group', appearance: { color: 'blue' } };
  const merged = buildServiceInspectorData(original, {
    label: ' Customer API ', description: 'Changed', pricing: config(),
  }, 'App Service');
  assert.equal(merged.label, 'Customer API');
  assert.equal(merged.serviceName, 'App Service');
  assert.equal(merged.icon, original.icon);
  assert.deepEqual(merged.appearance, original.appearance);
  assert.equal(original.label, 'App Service');
  assert.throws(() => buildServiceInspectorData(original, { label: ' ', description: '', pricing: config() }, 'App Service'));
  const usagePricing = pricingFromTier(parsePricingTiers([meter({ unitOfMeasure: '1 GB' })])[0], 1, 'eastus2', 5);
  const payload = buildServiceInspectorData(original, {
    label: 'Customer API', description: '', pricing: usagePricing,
  }, 'App Service').pricing as NodePricingConfig;
  payload.usage!.amount = 999;
  payload.provenance!.assumptions[0].value = 999;
  assert.equal(usagePricing.usage!.amount, 5);
  assert.equal(usagePricing.provenance!.assumptions[0].value, 5);
});

test('mixed-region subtotals and CSV text remain accurate and injection-safe', () => {
  const first = node(config());
  const second = { ...node({ ...config(), quantity: 2, region: 'japaneast' }), id: 'japan' };
  const partial = calculateCostBreakdown([first, second, { ...node(), id: 'unknown' }]);
  assert.equal(partial.totalMonthlyCost, 219);
  assert.equal(partial.region, 'Mixed (eastus2, japaneast)');
  assert.equal(partial.estimateCompleteness, 'partial');
  const csv = exportCostBreakdownCSV({
    ...partial, region: '=HYPERLINK("bad")',
    byService: [{ ...partial.byService[0], serviceName: '+SUM(1,2)' }],
  });
  assert.match(csv, /Priced Monthly Subtotal,\$219\.00/);
  assert.match(csv, /'=HYPERLINK/);
  assert.match(csv, /'\+SUM\(1,2\)/);
});

test('low-level monthly calculator does not turn missing SKUs or unknown consumption into zero', () => {
  const service = { serviceType: 'test', serviceName: 'test', defaultTier: 'B1',
    tiers: [instanceTier()], calculationType: 'hourly' as const, lastUpdated: '' };
  assert.equal(calculateMonthlyCost(service, 'B1', 3), 219);
  assert.equal(calculateMonthlyCost(service, 'missing', 3), null);
  assert.equal(calculateMonthlyCost({ ...service, tiers: [{ ...instanceTier(), monthlyPrice: null }] }, 'B1'), null);
});

test('Vite loads production compact assets with honest Fabric/OneLake and platform provenance (no listener)', async t => {
  const requests: string[] = [];
  t.mock.method(globalThis, 'fetch', async (input: unknown) => {
    const url = String(input);
    requests.push(url);
    const match = /^\/pricing\/regions\/([a-z0-9]+)\/([a-z0-9_]+)\.json$/.exec(url);
    if (!match) return { ok: false, status: 404 };
    try {
      const body = await readFile(new URL(`../public/pricing/regions/${match[1]}/${match[2]}.json`, import.meta.url), 'utf8');
      return { ok: true, status: 200, json: async () => JSON.parse(body) };
    } catch {
      return { ok: false, status: 404 };
    }
  });
  const { createServer } = await import('vite');
  const server = await createServer({
    configFile: false, logLevel: 'silent',
    server: { middlewareMode: true, watch: null, hmr: false },
    optimizeDeps: { noDiscovery: true, include: [] },
  });
  try {
    assert.equal(server.httpServer, null);
    const module = await server.ssrLoadModule('/src/services/costEstimationService.ts');
    const capacity = await module.initializeNodePricing('Microsoft Fabric Capacity', 'eastus2');
    assert.equal(capacity.provenance.kind, 'capacity');
    assert.equal(capacity.provenance.source, 'azure-retail-prices');
    assert.equal(capacity.provenance.meterId, undefined);
    assert.equal(capacity.provenance.asOf, undefined);
    assert.match(capacity.meterAsOf, /^\d{4}-\d{2}-\d{2}$/);
    assert.equal(capacity.provenance.snapshotAsOf, capacity.meterAsOf);
    assert.ok(capacity.estimatedCost > 0);
    assert.equal(capacity.skuName, 'F2');
    const larger = await module.updateNodePricing('Microsoft Fabric Capacity', capacity, 'F4', 3);
    assert.ok(Math.abs(larger.estimatedCost - capacity.estimatedCost * 2) < 0.01);
    const lake = await module.initializeNodePricing('OneLake Storage', 'eastus2');
    assert.equal(lake.provenance.kind, 'usage-estimate');
    assert.equal(lake.provenance.source, 'azure-retail-prices');
    assert.ok(lake.usage.amount > 0);
    assert.equal(lake.provenance.asOf, undefined);
    assert.ok(lake.provenance.snapshotAsOf);
    assert.ok(requests.includes('/pricing/regions/eastus2/microsoft_fabric.json'));

    const options = await module.getNodePricingTiers('App Service', 'eastus2');
    const linux = options.find((tier: PricingTier) => tier.name === 'S1 (Linux)');
    const windows = options.find((tier: PricingTier) => tier.name === 'S1 (Windows)');
    assert.ok(linux && windows);
    assert.notEqual(linux.id, windows.id);
    assert.equal(selectPricingTier(options, 'S1'), undefined, 'legacy platform must not be guessed');
    const selected = pricingFromTier(windows, 3, 'eastus2');
    const before = JSON.stringify(selected);
    const [regional] = await module.refreshAllNodePricing([node(selected)], 'japaneast');
    assert.equal(JSON.stringify(selected), before);
    assert.equal(regional.data.pricing.tierId, windows.id);
    assert.equal(regional.data.pricing.quantity, 3);
    const target = (await module.getNodePricingTiers('App Service', 'japaneast'))
      .find((tier: PricingTier) => tier.id === windows.id);
    assert.equal(regional.data.pricing.estimatedCost, target.monthlyPrice);
    assert.equal(regional.data.pricing.meterAsOf, target.provenance.snapshotAsOf);
    assert.equal(module.calculateCostBreakdown([regional]).totalMonthlyCost, target.monthlyPrice * 3);

    const legacy = { ...selected, tierId: undefined, tier: 'S1' };
    const recovered = await module.updateNodePricing('App Service', legacy, undefined, 2, 'japaneast');
    assert.equal(recovered.tierId, windows.id, 'source-region price uniquely recovers the saved platform');
    const ambiguous = { ...legacy, estimatedCost: 123456 };
    const missing = await module.updateNodePricing('App Service', ambiguous, undefined, undefined, 'japaneast');
    assert.equal(missing.skuName, 'S1');
    assert.equal(missing.estimatedCost, null);
    assert.equal(missing.meterAsOf, undefined);
    assert.equal(module.calculateCostBreakdown([node(missing)]).estimateCompleteness, 'unpriced');
    const resetCustom = await module.updateNodePricing('App Service', setCustomPricing(selected, 20), windows.id);
    assert.equal(resetCustom.isCustom, false);
    assert.equal(resetCustom.estimatedCost, windows.monthlyPrice);

    const ai = await module.initializeNodePricing('Azure AI Document Intelligence', 'eastus2');
    assert.equal(ai.provenance.source, 'azure-retail-prices');
    assert.ok(requests.includes('/pricing/regions/eastus2/foundry_tools.json'));
  } finally {
    await server.close();
  }
});
