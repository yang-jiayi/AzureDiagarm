import type { NodePricingConfig, PricingProvenance, PricingTier, PricingUsage } from '../types/pricing';
import { FALLBACK_PRICING, FABRIC_CAPACITY_SKUS, USAGE_BASED_SERVICES, getAzureServiceName } from '../data/azurePricing';
import { isCapacityConsumed } from '../data/serviceIconMapping';
import { applyRegionalPricing, getRegionalMultiplier } from '../utils/pricingHelpers';

export const MAX_PRICING_QUANTITY = 100_000;
export const MAX_PRICING_AMOUNT = 1e12;

export function validatePricingQuantity(quantity: number): void {
  if (!Number.isInteger(quantity) || quantity < 1 || quantity > MAX_PRICING_QUANTITY) {
    throw new Error(`Quantity must be an integer between 1 and ${MAX_PRICING_QUANTITY}`);
  }
}

export function validatePricingAmount(amount: number): void {
  if (!Number.isFinite(amount) || amount < 0 || amount > MAX_PRICING_AMOUNT) {
    throw new Error('Amount must be finite and between 0 and 1000000000000');
  }
}

export function isPricingProvenance(value: unknown): value is PricingProvenance {
  if (!value || typeof value !== 'object' || Array.isArray(value)) return false;
  const provenance = value as Record<string, unknown>;
  return typeof provenance.kind === 'string' &&
    ['official-meter', 'usage-estimate', 'fallback-estimate', 'custom', 'unpriced', 'capacity', 'unknown'].includes(provenance.kind) &&
    typeof provenance.unit === 'string' &&
    (provenance.source === undefined || (typeof provenance.source === 'string' &&
      ['azure-retail-prices', 'bundled-fallback', 'user', 'legacy'].includes(provenance.source))) &&
    ['asOf', 'snapshotAsOf', 'meterId', 'meterName', 'note'].every(key => provenance[key] === undefined || typeof provenance[key] === 'string') &&
    Array.isArray(provenance.assumptions) && provenance.assumptions.every(assumption =>
      assumption && typeof assumption === 'object' && typeof assumption.label === 'string' &&
      typeof assumption.unit === 'string' && (typeof assumption.value === 'string' ||
        (typeof assumption.value === 'number' && Number.isFinite(assumption.value))));
}

export function isPricingUsage(value: unknown): value is PricingUsage {
  if (!value || typeof value !== 'object' || Array.isArray(value)) return false;
  const usage = value as Record<string, unknown>;
  const amount = (candidate: unknown) => typeof candidate === 'number' &&
    Number.isFinite(candidate) && candidate >= 0 && candidate <= MAX_PRICING_AMOUNT;
  return typeof usage.unit === 'string' && !!usage.unit.trim() && amount(usage.unitPrice) &&
    (usage.amount === null || amount(usage.amount));
}

export function getPricingProvenance(config?: NodePricingConfig | null): PricingProvenance {
  const provenance = config?.provenance;
  return isPricingProvenance(provenance) ? {
    ...provenance, assumptions: provenance.assumptions.map(assumption => ({ ...assumption })),
  } : {
    kind: config ? 'unknown' : 'unpriced',
    source: config ? 'legacy' : undefined,
    unit: typeof config?.unit === 'string' && config.unit ? config.unit : 'USD/unit/month',
    assumptions: [],
    note: config ? 'Imported price; source, assumptions and data date are unknown.' : 'No price is available.',
  };
}

export function isUsageBasedService(serviceType: string): boolean {
  const mapped = getAzureServiceName(serviceType).toLowerCase();
  return USAGE_BASED_SERVICES.some(service => getAzureServiceName(service).toLowerCase() === mapped);
}

export function getFallbackPricingTiers(serviceType: string, region: string): PricingTier[] {
  if (isCapacityConsumed(serviceType)) {
    return [{
      name: 'Included in capacity', skuName: 'Included in capacity', monthlyPrice: null,
      unit: 'shared capacity',
      provenance: {
        kind: 'capacity', source: 'bundled-fallback', unit: 'shared capacity',
        assumptions: [], note: 'Consumes shared Fabric capacity; price the capacity separately. Not a free service.',
      },
    }];
  }
  const normalized = (value: string) => value.trim().replace(/^(?:Azure|Microsoft)\s+/i, '').toLowerCase();
  const key = Object.prototype.hasOwnProperty.call(FALLBACK_PRICING, serviceType) ? serviceType :
    Object.keys(FALLBACK_PRICING).find(candidate => normalized(candidate) === normalized(serviceType)) ??
    Object.keys(FALLBACK_PRICING).find(candidate =>
      getAzureServiceName(candidate).toLowerCase() === getAzureServiceName(serviceType).toLowerCase());
  const table = key ? FALLBACK_PRICING[key] : undefined;
  if (!table) return [];
  const entries = key === 'Microsoft Fabric Capacity'
    ? Object.entries(FABRIC_CAPACITY_SKUS).map(([sku, value]) => [sku, value.paygMonthly] as const)
    : (['basic', 'standard', 'premium'] as const).map(level =>
      [level[0].toUpperCase() + level.slice(1), table[level]] as const);
  return entries.map(([name, price]) => ({
    name, skuName: name,
    // A zero fallback often means consumption/no fixed charge, not free usage.
    monthlyPrice: price > 0 ? applyRegionalPricing(price, region) : null,
    unit: table.unit,
    isUsageBased: isUsageBasedService(serviceType),
    provenance: {
      kind: price > 0 ? 'fallback-estimate' : 'unpriced',
      source: 'bundled-fallback',
      asOf: '2026-01',
      unit: table.unit,
      assumptions: [
        { label: 'Bundled estimate level', value: name, unit: 'level (not a verified regional SKU)' },
        { label: 'Regional estimate multiplier', value: getRegionalMultiplier(region), unit: '× baseline' },
      ],
      note: 'Static estimate, not an official regional quote. Detailed usage assumptions are unavailable.',
    },
  }));
}

export function pricingFromTier(
  tier: PricingTier,
  quantity: number,
  region: string,
  usageAmount?: number | null,
): NodePricingConfig {
  validatePricingQuantity(quantity);
  let estimatedCost = tier.monthlyPrice;
  let provenance = tier.provenance
    ? { ...tier.provenance, assumptions: tier.provenance.assumptions.map(value => ({ ...value })) }
    : {
      kind: 'unknown' as const, source: 'legacy' as const, unit: tier.unit, assumptions: [],
      note: 'Price source, assumptions and data date are unknown.',
    };
  const usage = tier.usage ? { ...tier.usage } : undefined;
  if (usage) {
    validatePricingAmount(usage.unitPrice);
    usage.amount = usageAmount === undefined ? usage.amount : usageAmount;
    if (usage.amount !== null) {
      validatePricingAmount(usage.amount);
      estimatedCost = usage.amount * usage.unitPrice;
      provenance = {
        ...provenance, kind: 'usage-estimate',
        assumptions: [{ label: 'Monthly usage', value: usage.amount, unit: usage.unit }],
      };
    } else {
      estimatedCost = null;
      provenance = { ...provenance, kind: 'unpriced', assumptions: [] };
    }
  }
  if (estimatedCost !== null) validatePricingAmount(estimatedCost);
  const reserved = !usage && !tier.isUsageBased && tier.reserved1yrMonthly !== undefined &&
    Number.isFinite(tier.reserved1yrMonthly) && tier.reserved1yrMonthly >= 0 &&
    tier.reserved1yrMonthly <= MAX_PRICING_AMOUNT ? tier.reserved1yrMonthly : undefined;
  return {
    estimatedCost, quantity, region,
    tier: tier.name, tierId: tier.id, skuName: tier.skuName, unit: tier.unit,
    isCustom: false,
    isUsageBased: !!tier.isUsageBased || !!usage || provenance.kind === 'usage-estimate',
    lastUpdated: new Date().toISOString(), provenance, usage,
    reserved1yrCost: reserved,
    reservedIsSavingsPlan: reserved !== undefined,
    meterAsOf: estimatedCost !== null && provenance.source === 'azure-retail-prices'
      ? provenance.snapshotAsOf : undefined,
  };
}

/** Reuse the exact selected SKU; a region change must not silently switch tiers. */
export function selectPricingTier(
  tiers: PricingTier[], selection: string,
): PricingTier | undefined {
  const exactId = tiers.find(tier => tier.id === selection);
  if (exactId) return exactId;
  const named = tiers.filter(tier => tier.name === selection);
  if (named.length) return named.length === 1 ? named[0] : undefined;
  const candidates = tiers.filter(tier => tier.skuName === selection);
  return candidates.length === 1 ? candidates[0] : undefined;
}

export function withUsageAmount(config: NodePricingConfig, amount: number | null): NodePricingConfig {
  if (!isPricingUsage(config.usage) || getPricingProvenance(config).source !== 'azure-retail-prices') {
    throw new Error('No substantiated usage meter is available');
  }
  return pricingFromTier({
    id: config.tierId, name: config.tier, skuName: config.skuName, monthlyPrice: config.estimatedCost,
    unit: config.unit, provenance: config.provenance, usage: config.usage,
  }, config.quantity, config.region, amount);
}

export function buildServiceInspectorData(
  data: Record<string, unknown>,
  edits: { label: string; description: string; pricing: NodePricingConfig },
  serviceType: string,
): Record<string, unknown> {
  const label = edits.label.trim();
  if (!label || label.length > 200 || edits.description.length > 4000) {
    throw new Error('Invalid service label or description');
  }
  validatePricingQuantity(edits.pricing.quantity);
  if (edits.pricing.estimatedCost !== null) validatePricingAmount(edits.pricing.estimatedCost);
  return {
    ...data, label, description: edits.description,
    // Preserve service identity when its user-facing label changes.
    serviceName: data.serviceName || serviceType,
    pricing: {
      ...edits.pricing,
      provenance: getPricingProvenance(edits.pricing),
      ...(edits.pricing.usage ? { usage: { ...edits.pricing.usage } } : {}),
      ...(edits.pricing.usageEstimate ? { usageEstimate: { ...edits.pricing.usageEstimate } } : {}),
    },
  };
}
