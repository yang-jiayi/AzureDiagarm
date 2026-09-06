// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

/**
 * Cost Estimation Service
 * Core logic for calculating architecture costs
 */

import { Node } from 'reactflow';
import { 
  NodePricingConfig, 
  CostBreakdown,
  PricingTier
} from '../types/pricing';
import { 
  getServicePricing
} from './azurePricingService';
import { 
  getActiveRegion
} from './regionalPricingService';
import { 
  getAzureServiceName, 
  getDefaultTier, 
  getFallbackDefaultSku,
  getReserved1yrDiscount,
  PRICING_DATA_AS_OF
} from '../data/azurePricing';
import { 
  getPricingFreshness
} from '../utils/pricingHelpers';
import {
  getFallbackPricingTiers, getPricingProvenance, pricingFromTier,
  selectPricingTier, validatePricingAmount, validatePricingQuantity, isPricingUsage,
  isUsageBasedService, MAX_PRICING_QUANTITY,
} from './pricingConfiguration';
import { isCapacityConsumed } from '../data/serviceIconMapping';
import { csvTextCell } from '../utils/csv';

export async function getNodePricingTiers(serviceType: string, region: string): Promise<PricingTier[]> {
  if (isCapacityConsumed(serviceType)) return getFallbackPricingTiers(serviceType, region);
  try {
    const pricing = await getServicePricing(serviceType, getAzureServiceName(serviceType), region);
    if (pricing?.tiers.length) return pricing.tiers.map(tier => ({
      ...tier,
      provenance: tier.provenance ? {
        ...tier.provenance,
        assumptions: tier.provenance.assumptions.map(assumption => ({ ...assumption })),
        snapshotAsOf: pricing.meterAsOf,
      } : undefined,
      ...(tier.usage ? { usage: { ...tier.usage } } : {}),
    }));
  } catch (error) {
    console.warn(`Regional pricing unavailable for ${serviceType}`, error);
  }
  return getFallbackPricingTiers(serviceType, region);
}

/**
 * Initialize pricing for a new node
 */
export async function initializeNodePricing(
  serviceType: string,
  region?: string
): Promise<NodePricingConfig | null> {
  const targetRegion = region || getActiveRegion();
  const tiers = await getNodePricingTiers(serviceType, targetRegion);
  const defaultTier = getDefaultTier(serviceType);
  const tier = selectPricingTier(tiers, defaultTier) ??
    tiers.find(candidate => candidate.skuName === defaultTier) ??
    selectPricingTier(tiers, getFallbackDefaultSku(serviceType)) ?? tiers[0];
  if (!tier) return null;
  const config = pricingFromTier(tier, 1, targetRegion);
  return { ...config, isUsageBased: config.isUsageBased || isUsageBasedService(serviceType) };
}

/**
 * Update pricing when tier or quantity changes
 */
function pricesMatch(left: number | null, right: number | null): boolean {
  return typeof left === 'number' && Number.isFinite(left) &&
    typeof right === 'number' && Number.isFinite(right) &&
    Math.abs(left - right) <= Math.max(0.01, Math.abs(right) * 0.001);
}

function findTierCandidates(tiers: PricingTier[], selector: string): PricingTier[] {
  const exactIdMatch = tiers.find(candidate => candidate.id === selector);
  if (exactIdMatch) return [exactIdMatch];

  const exactNameMatches = tiers.filter(candidate => candidate.name === selector);
  if (exactNameMatches.length > 0) return exactNameMatches;

  return tiers.filter(candidate => candidate.skuName === selector);
}

async function resolvePricingTier(
  serviceType: string,
  currentConfig: NodePricingConfig,
  selector: string,
  region: string,
  tiers: PricingTier[],
): Promise<PricingTier> {
  const matchingTiers = findTierCandidates(tiers, selector);
  if (matchingTiers.length === 0) {
    throw new Error(`Selected SKU ${selector} is not available for ${serviceType} in ${region}`);
  }
  if (matchingTiers.length === 1) return matchingTiers[0];
  const identifiedTier = matchingTiers.find(candidate => candidate.id === currentConfig.tierId);
  if (identifiedTier) return identifiedTier;

  // New App Service tiers have platform-specific ids, but older saved diagrams
  // only stored a shared SKU such as "S1". Recover the original platform from
  // its source-region unit price, then select that same product in the target
  // region. If the old estimate cannot identify one platform uniquely, keep
  // treating it as ambiguous instead of silently changing the workload.
  const namedTier = matchingTiers.filter(candidate => candidate.name === currentConfig.tier);
  if (namedTier.length === 1) return namedTier[0];

  const sourceTiers = currentConfig.region === region
    ? tiers
    : await getNodePricingTiers(serviceType, currentConfig.region);
  if (!currentConfig.isCustom) {
    const sourceCandidates = findTierCandidates(sourceTiers, selector);
    const sourcePriceMatches = sourceCandidates.filter(candidate =>
      pricesMatch(candidate.monthlyPrice, currentConfig.estimatedCost)
    );
    if (sourcePriceMatches.length === 1 && sourcePriceMatches[0].id) {
      const targetTier = tiers.find(candidate => candidate.id === sourcePriceMatches[0].id);
      if (targetTier) return targetTier;
    }
  }

  throw new Error(`Selected SKU ${selector} is ambiguous for ${serviceType}; select a platform-specific SKU`);
}

export async function updateNodePricing(
  serviceType: string,
  currentConfig: NodePricingConfig,
  newTier?: string,
  newQuantity?: number,
  newRegion?: string
): Promise<NodePricingConfig> {
  const tier = newTier ?? currentConfig.tierId ?? currentConfig.skuName ?? currentConfig.tier;
  const quantity = newQuantity ?? currentConfig.quantity;
  const region = newRegion ?? currentConfig.region;
  validatePricingQuantity(quantity);
  const sameTier = tier === currentConfig.tierId || tier === currentConfig.tier || tier === currentConfig.skuName;
  if ((newTier === undefined && region === currentConfig.region) ||
      (sameTier && currentConfig.isCustom && (newTier === undefined || region !== currentConfig.region))) {
    return {
      ...currentConfig, quantity, region, provenance: getPricingProvenance(currentConfig),
      ...(currentConfig.usage ? { usage: { ...currentConfig.usage } } : {}),
      ...(currentConfig.usageEstimate ? { usageEstimate: { ...currentConfig.usageEstimate } } : {}),
    };
  }
  const currentSource = getPricingProvenance(currentConfig).source;
  const available = sameTier && currentSource === 'bundled-fallback'
    ? getFallbackPricingTiers(serviceType, region)
    : await getNodePricingTiers(serviceType, region);
  // A generic fallback level with the same spelling is not a verified match
  // for an unavailable regional SKU.
  const tiers = sameTier && region !== currentConfig.region && currentSource !== 'bundled-fallback'
    ? available.filter(candidate => candidate.provenance?.source !== 'bundled-fallback')
    : available;
  let selected: PricingTier;
  try {
    selected = await resolvePricingTier(serviceType, currentConfig, tier, region, tiers);
  } catch (error) {
    if (!sameTier || region === currentConfig.region) throw error;
    return {
      ...currentConfig, quantity, region, estimatedCost: null,
      reserved1yrCost: undefined, reservedIsSavingsPlan: false,
      meterAsOf: undefined, lastUpdated: new Date().toISOString(),
      usage: isPricingUsage(currentConfig.usage) ? { ...currentConfig.usage } : undefined,
      provenance: {
        kind: 'unpriced', unit: currentConfig.unit, assumptions: [],
        meterName: getPricingProvenance(currentConfig).meterName,
        note: 'Selected SKU is unavailable or ambiguous in this region. Choose an available tier or supply a custom estimate.',
      },
    };
  }
  const amount = sameTier && isPricingUsage(currentConfig.usage) &&
    selected.usage?.unit === currentConfig.usage.unit &&
    selected.provenance?.meterName === currentConfig.provenance?.meterName
    ? currentConfig.usage.amount : undefined;
  const updated = pricingFromTier(selected, quantity, region, amount);
  return {
    ...updated,
    isUsageBased: updated.isUsageBased || isUsageBasedService(serviceType),
    ...(sameTier && currentConfig.usageEstimate &&
      (amount !== undefined || currentSource === 'bundled-fallback')
      ? { usageEstimate: { ...currentConfig.usageEstimate } } : {}),
  };
}

/**
 * Set custom pricing for a node
 */
export function setCustomPricing(
  currentConfig: NodePricingConfig,
  customPrice: number
): NodePricingConfig {
  validatePricingAmount(customPrice);
  validatePricingQuantity(currentConfig.quantity);
  return {
    ...currentConfig,
    estimatedCost: customPrice,
    customPrice: customPrice,
    isCustom: true,
    lastUpdated: new Date().toISOString(),
    reserved1yrCost: undefined,
    reservedIsSavingsPlan: false,
    meterAsOf: undefined,
    usage: undefined,
    provenance: {
      kind: 'custom', source: 'user', asOf: new Date().toISOString(),
      unit: 'USD/unit/month', assumptions: [],
      note: 'User-provided per-unit monthly estimate; not an official quote.',
    },
  };
}

/**
 * Tiers/SKUs a service can be switched to, for the per-node cost editor.
 *
 * Kept for the existing per-node editor; uses the same catalog/fallback
 * semantics as the inspector and preserves unpriced consumption meters.
 */
export async function getAvailableTiers(
  serviceType: string,
  region?: string
): Promise<PricingTier[]> {
  return getNodePricingTiers(serviceType, region || getActiveRegion());
}

/**
 * Billing term used for cost estimates.
 * - 'payg': pay-as-you-go list price
 * - 'reserved1yr': 1-year commitment. Uses the meter's real 1-year Savings
 *   Plan rate when available; otherwise falls back to a representative discount
 *   on reservation-eligible, non-usage-based services. Usage-based/consumption
 *   services always stay at PAYG.
 */
export type PricingMode = 'payg' | 'reserved1yr';

/**
 * How long a price must have held before its age is worth reporting.
 *
 * Set from the shipped corpus, not from intuition: at a year, 48% of services
 * qualify and the line appears on nearly every deck, which makes it furniture.
 * At two years it is 33%, and the services it picks out are the ones the claim
 * is actually interesting for — Log Analytics, Application Insights, IoT Hub,
 * Azure ML, Data Factory, all repriced more than five years ago.
 */
const STABLE_PRICE_DAYS = 730;

/** Whole days between two ISO calendar dates, or 0 if either is unusable. */
function ageInDays(from: string, to: string): number {
  const a = Date.parse(`${from}T00:00:00Z`);
  const b = Date.parse(`${to}T00:00:00Z`);
  if (Number.isNaN(a) || Number.isNaN(b)) return 0;
  return Math.floor((b - a) / 86_400_000);
}

/**
 * Calculate total cost breakdown for all nodes
 */
export function calculateCostBreakdown(
  nodes: Node[],
  region?: string,
  pricingMode: PricingMode = 'payg'
): CostBreakdown {
  const targetRegion = region || getActiveRegion();
  // Initialize breakdown
  const breakdown: CostBreakdown = {
    totalMonthlyCost: 0,
    byService: [],
    byGroup: [],
    byCategory: [],
    region: targetRegion,
    currency: 'USD',
    lastCalculated: new Date().toISOString(),
    pricesAsOf: PRICING_DATA_AS_OF,
    pricingTerm: pricingMode === 'reserved1yr' ? 'Savings Plan (1-year)' : 'Pay-as-you-go',
    unpricedServices: [],
    capacityServices: [],
  };

  // Track costs by group and category
  const groupCosts = new Map<string, { label: string; cost: number; count: number }>();
  const categoryCosts = new Map<string, number>();
  let hasPricedCapacity = false;
  const pricingRegions = new Set<string>();
  // The oldest meter behind any number on this page. Read off the nodes, not
  // from a name lookup: the date was stamped on the pricing config by whichever
  // load produced the figure, so it is already the right region's, it is
  // already absent for anything the static fallback priced, and it survives
  // save and restore along with the estimate it belongs to.
  let oldestMeterAsOf: string | undefined;

  // Calculate per-service costs
  nodes.forEach(node => {
    if (node.type === 'groupNode') return;
    const pricing = node.data.pricing as NodePricingConfig | undefined;
    const provenance = getPricingProvenance(pricing);
    const serviceName = node.data.label || 'Unnamed Service';
    const serviceType = node.data.serviceName || node.data.serviceType || node.data.label || '';
    if (isCapacityConsumed(serviceType) && (!pricing || !pricing.isCustom)) {
      breakdown.capacityServices!.push({ nodeId: node.id, serviceName });
      return;
    }
    if (pricing) {
      const pricingRegion = typeof pricing.region === 'string' ? pricing.region.trim() : '';
      pricingRegions.add(pricingRegion || 'Unknown');
    }
    if (!pricing || typeof pricing.estimatedCost !== 'number' || !Number.isFinite(pricing.estimatedCost) ||
        pricing.estimatedCost < 0 || pricing.estimatedCost > 1e12 || !Number.isInteger(pricing.quantity) ||
        pricing.quantity < 1 || pricing.quantity > MAX_PRICING_QUANTITY || provenance.kind === 'unpriced') {
      breakdown.unpricedServices!.push({
        nodeId: node.id, serviceName,
        reason: provenance.kind === 'unpriced' && provenance.note
          ? provenance.note : 'No valid monthly estimate is available.',
      });
      return;
    }
    // A custom price is a number the user typed, so no Azure meter stands
    // behind it and it must not drag the reported vintage backwards.
    if (!pricing.isCustom && provenance.source !== 'bundled-fallback' &&
        typeof pricing.meterAsOf === 'string' && /^\d{4}-\d{2}-\d{2}$/.test(pricing.meterAsOf)) {
      if (oldestMeterAsOf === undefined || pricing.meterAsOf < oldestMeterAsOf) {
        oldestMeterAsOf = pricing.meterAsOf;
      }
    }

    let cost = pricing.estimatedCost * pricing.quantity;
    // Apply the 1-year commitment to reservation-eligible, non-usage-based
    // services. Prefer the meter's REAL 1-year Savings Plan rate; only fall
    // back to the representative discount table when no savings-plan rate is
    // known. Custom, legacy-unknown and usage-based prices are not discounted.
    if (pricingMode === 'reserved1yr' && !pricing.isUsageBased && !pricing.isCustom &&
        !isPricingUsage(pricing.usage) && provenance.kind !== 'usage-estimate' &&
        provenance.kind !== 'unknown' && provenance.kind !== 'custom') {
      if (pricing.reserved1yrCost != null && Number.isFinite(pricing.reserved1yrCost) &&
          pricing.reserved1yrCost >= 0 && pricing.reserved1yrCost <= 1e12) {
        cost = pricing.reserved1yrCost * pricing.quantity;
      } else {
        const discount = getReserved1yrDiscount(serviceType);
        if (discount > 0) cost = cost * (1 - discount);
      }
    }
    breakdown.totalMonthlyCost += cost;
    if (serviceType === 'Microsoft Fabric Capacity' || provenance.kind === 'capacity') hasPricedCapacity = true;

    // Add to service breakdown
    breakdown.byService.push({
      serviceName: node.data.label || 'Unnamed Service',
      serviceType: node.data.category || 'Other',
      nodeId: node.id,
      cost: cost,
      quantity: pricing.quantity,
      tier: typeof pricing.tier === 'string' && pricing.tier
        ? pricing.tier : typeof pricing.skuName === 'string' && pricing.skuName ? pricing.skuName : 'Unspecified'
    });

    // Track by group
    const groupId = node.data.groupId || 'ungrouped';
    if (!groupCosts.has(groupId)) {
      groupCosts.set(groupId, {
        label: node.data.groupLabel || 'Ungrouped',
        cost: 0,
        count: 0
      });
    }
    const groupData = groupCosts.get(groupId)!;
    groupData.cost += cost;
    groupData.count += 1;

    // Track by category
    const category = node.data.category || 'Other';
    categoryCosts.set(category, (categoryCosts.get(category) || 0) + cost);
  });

  if (pricingRegions.size === 1) {
    breakdown.region = [...pricingRegions][0];
  } else if (pricingRegions.size > 1) {
    breakdown.region = `Mixed (${[...pricingRegions].sort().join(', ')})`;
  }

  // Convert group costs to array
  breakdown.byGroup = Array.from(groupCosts.entries()).map(([groupId, data]) => ({
    groupId,
    groupLabel: data.label,
    cost: data.cost,
    serviceCount: data.count
  }));

  // Convert category costs to array with percentages
  breakdown.byCategory = Array.from(categoryCosts.entries()).map(([category, cost]) => ({
    category,
    cost,
    percentage: breakdown.totalMonthlyCost > 0 ? (cost / breakdown.totalMonthlyCost) * 100 : 0
  }));

  // Sort all arrays by cost (descending)
  breakdown.byService.sort((a, b) => b.cost - a.cost);
  breakdown.byGroup.sort((a, b) => b.cost - a.cost);
  breakdown.byCategory.sort((a, b) => b.cost - a.cost);

  breakdown.pricedMonthlySubtotal = breakdown.totalMonthlyCost;
  breakdown.missingCapacityEstimate = !!breakdown.capacityServices?.length && !hasPricedCapacity;
  const incomplete = !!breakdown.unpricedServices?.length || breakdown.missingCapacityEstimate;
  breakdown.estimateCompleteness = incomplete
    ? breakdown.byService.length ? 'partial' : 'unpriced'
    : 'complete';
  // Only worth saying when the prices have actually held for a long time. Every
  // meter predates the download by some margin, so reporting any gap at all
  // would put a second date on every slide that never means anything.
  //
  // Measured over all 403 non-empty pricing files, the share that clears each
  // bar is: 30d 85%, 90d 68%, 180d 60%, 365d 37%, 730d 23% (per service, of 48
  // distinct: 92 / 75 / 67 / 48 / 33%). A year sounds like the natural line but
  // at ~48% per service it still fires on nearly every deck — `virtual_network`
  // and `key_vault` are on almost every diagram and both clear it — so the line
  // would be furniture again. Two years is the bar that isolates the cases this
  // is for: Log Analytics, Application Insights, IoT Hub, Azure ML and Data
  // Factory, all repriced somewhere between five and eight years ago.
  const oldest = oldestMeterAsOf;
  if (oldest && ageInDays(oldest, PRICING_DATA_AS_OF) > STABLE_PRICE_DAYS) {
    breakdown.oldestMeterAsOf = oldest;
  }

  return breakdown;
}

function unavailableImportedPricing(value: unknown, region: string): NodePricingConfig {
  const raw = value && typeof value === 'object' && !Array.isArray(value)
    ? value as Partial<NodePricingConfig> : {};
  const quantity = typeof raw.quantity === 'number' && Number.isInteger(raw.quantity) &&
    raw.quantity >= 1 && raw.quantity <= MAX_PRICING_QUANTITY ? raw.quantity : 1;
  const tier = typeof raw.tier === 'string' ? raw.tier : '';
  const skuName = typeof raw.skuName === 'string' ? raw.skuName : tier;
  const unit = typeof raw.unit === 'string' && raw.unit ? raw.unit : 'USD/unit/month';
  return {
    ...raw, estimatedCost: null, quantity, tier, skuName, region, unit,
    isCustom: raw.isCustom === true,
    lastUpdated: new Date().toISOString(),
    customPrice: typeof raw.customPrice === 'number' && Number.isFinite(raw.customPrice) &&
      raw.customPrice >= 0 && raw.customPrice <= 1e12 ? raw.customPrice : undefined,
    reserved1yrCost: undefined, reservedIsSavingsPlan: false, meterAsOf: undefined,
    usage: isPricingUsage(raw.usage) ? { ...raw.usage } : undefined,
    provenance: {
      kind: 'unpriced', source: 'legacy', unit, assumptions: [],
      note: 'Imported pricing is incomplete or invalid. Confirm quantity and a monthly estimate before including it in the subtotal.',
    },
  };
}

/**
 * Refresh pricing for all nodes (when region changes)
 */
export async function refreshAllNodePricing(
  nodes: Node[],
  newRegion: string
): Promise<Node[]> {
  const updatedNodes: Node[] = [];

  for (const node of nodes) {
    if (node.type !== 'groupNode' && node.data.pricing) {
      const serviceType = node.data.serviceName || node.data.serviceType || node.data.label || 'Unknown';
      let updatedPricing: NodePricingConfig;
      try {
        updatedPricing = await updateNodePricing(
          serviceType,
          node.data.pricing,
          undefined,
          node.data.pricing.quantity,
          newRegion
        );
      } catch {
        // A malformed legacy record must not prevent other nodes from refreshing.
        // The structural quantity default never contributes a guessed cost.
        updatedPricing = unavailableImportedPricing(node.data.pricing, newRegion);
      }

      updatedNodes.push({
        ...node,
        data: {
          ...node.data,
          pricing: updatedPricing
        }
      });
    } else {
      updatedNodes.push(node);
    }
  }

  return updatedNodes;
}

/**
 * Get cost summary text for export
 */
export function getCostSummaryText(breakdown: CostBreakdown): string {
  const lines: string[] = [];
  
  lines.push('=== COST ESTIMATION SUMMARY ===');
  lines.push('');
  lines.push(`${breakdown.estimateCompleteness && breakdown.estimateCompleteness !== 'complete'
    ? 'Priced Monthly Subtotal' : 'Total Monthly Cost'}: $${breakdown.totalMonthlyCost.toFixed(2)}`);
  if (breakdown.estimateCompleteness === 'unpriced') lines.push('No usable total estimate is available; unknown prices are not free.');
  if (breakdown.unpricedServices?.length) {
    lines.push(`Partial subtotal: excludes ${breakdown.unpricedServices.length} unpriced service(s).`);
    breakdown.unpricedServices.forEach(service => lines.push(`  Unpriced: ${service.serviceName} — ${service.reason || 'No usable monthly estimate is available.'}`));
  }
  if (breakdown.capacityServices?.length) lines.push('Shared-capacity workloads require a separately priced capacity.');
  if (breakdown.missingCapacityEstimate) lines.push('Incomplete estimate: no shared capacity estimate is present.');
  lines.push(`Region: ${breakdown.region}`);
  lines.push(`Currency: ${breakdown.currency}`);
  if (breakdown.pricingTerm) lines.push(`Pricing term: ${breakdown.pricingTerm}`);
  if (breakdown.pricesAsOf) {
    const f = getPricingFreshness(breakdown.pricesAsOf);
    lines.push(`Prices as of: ${breakdown.pricesAsOf}${f.isStale ? ` (⚠️ ${f.ageLabel} — refresh with "npm run pricing:refresh")` : ''}`);
  }
  if (breakdown.oldestMeterAsOf) {
    lines.push(`Oldest unchanged price: ${breakdown.oldestMeterAsOf} `
      + '(these are current Azure prices; the date is when Azure last changed the longest-standing one)');
  }
  lines.push(`Last Updated: ${new Date(breakdown.lastCalculated).toLocaleString()}`);
  lines.push('');
  
  lines.push('BY SERVICE:');
  breakdown.byService.forEach(svc => {
    lines.push(`  ${svc.serviceName} (${svc.tier}): $${svc.cost.toFixed(2)}/mo (${svc.quantity} units total)`);
  });
  lines.push('');
  
  lines.push('BY GROUP:');
  breakdown.byGroup.forEach(grp => {
    lines.push(`  ${grp.groupLabel}: $${grp.cost.toFixed(2)}/mo (${grp.serviceCount} services)`);
  });
  lines.push('');
  
  lines.push('BY CATEGORY:');
  breakdown.byCategory.forEach(cat => {
    lines.push(`  ${cat.category}: $${cat.cost.toFixed(2)}/mo (${cat.percentage.toFixed(1)}%)`);
  });

  if (breakdown.unpricedServices?.length) {
    lines.push('');
    lines.push(`NOT INCLUDED IN THE SUBTOTAL (${breakdown.unpricedServices.length} service(s) without usable monthly estimates):`);
    breakdown.unpricedServices.forEach(svc => {
      lines.push(`  ${svc.serviceName}`);
    });
  }

  return lines.join('\n');
}

/**
 * Get cost summary as Markdown for export.
 *
 * Produces a well-formatted Markdown document with headings and tables that
 * render correctly in GitHub, VS Code preview, Confluence, Teams, etc.
 */
export function getCostSummaryMarkdown(breakdown: CostBreakdown): string {
  const lines: string[] = [];
  const annual = breakdown.totalMonthlyCost * 12;

  lines.push('# Azure Architecture — Cost Estimation Summary');
  lines.push('');
  lines.push(`> **${breakdown.estimateCompleteness && breakdown.estimateCompleteness !== 'complete' ? 'Priced subtotal' : 'Total'}: \`$${breakdown.totalMonthlyCost.toFixed(2)}/mo\`** · **\`$${annual.toFixed(2)}/yr\`** · Region: \`${breakdown.region}\` · ${breakdown.currency}${breakdown.pricingTerm ? ` · ${breakdown.pricingTerm}` : ''}`);
  if (breakdown.estimateCompleteness === 'unpriced') lines.push('> No usable total estimate is available; unknown prices are not free.');
  if (breakdown.unpricedServices?.length) {
    lines.push(`> **Partial subtotal:** ${breakdown.unpricedServices.length} unpriced service(s) excluded, not free.`);
    breakdown.unpricedServices.forEach(service => lines.push(`> Unpriced: ${escapeMd(service.serviceName)} — ${escapeMd(service.reason || 'No usable monthly estimate is available.')}`));
  }
  if (breakdown.capacityServices?.length) lines.push('> Shared-capacity workloads require a separately priced capacity.');
  if (breakdown.missingCapacityEstimate) lines.push('> Incomplete estimate: no shared capacity estimate is present.');
  lines.push('');
  lines.push('| Field | Value |');
  lines.push('| --- | --- |');
  lines.push(`| ${breakdown.estimateCompleteness && breakdown.estimateCompleteness !== 'complete'
    ? 'Priced monthly subtotal' : 'Total monthly cost'} | **$${breakdown.totalMonthlyCost.toFixed(2)}** |`);
  lines.push(`| Annual projection | $${annual.toFixed(2)} |`);
  lines.push(`| Region | ${breakdown.region} |`);
  lines.push(`| Currency | ${breakdown.currency} |`);
  if (breakdown.pricingTerm) lines.push(`| Pricing term | ${breakdown.pricingTerm} |`);
  if (breakdown.pricesAsOf) lines.push(`| Prices as of | ${breakdown.pricesAsOf} |`);
  if (breakdown.oldestMeterAsOf) lines.push(`| Oldest unchanged price | ${breakdown.oldestMeterAsOf} |`);
  lines.push(`| Last updated | ${new Date(breakdown.lastCalculated).toLocaleString()} |`);
  lines.push('');

  lines.push('## By service');
  lines.push('');
  lines.push('| Service | Tier | Qty | Monthly cost |');
  lines.push('| --- | --- | ---: | ---: |');
  breakdown.byService.forEach(svc => {
    lines.push(`| ${escapeMd(svc.serviceName)} | ${escapeMd(svc.tier)} | ${svc.quantity} | $${svc.cost.toFixed(2)} |`);
  });
  lines.push(`| **${breakdown.estimateCompleteness && breakdown.estimateCompleteness !== 'complete' ? 'Subtotal' : 'Total'}** | | | **$${breakdown.totalMonthlyCost.toFixed(2)}** |`);
  lines.push('');

  if (breakdown.byGroup.length > 0) {
    lines.push('## By group');
    lines.push('');
    lines.push('| Group | Services | Monthly cost |');
    lines.push('| --- | ---: | ---: |');
    breakdown.byGroup.forEach(grp => {
      lines.push(`| ${escapeMd(grp.groupLabel)} | ${grp.serviceCount} | $${grp.cost.toFixed(2)} |`);
    });
    lines.push('');
  }

  lines.push('---');
  lines.push('');
  if (breakdown.unpricedServices?.length) {
    lines.push('## Not included in the total');
    lines.push('');
    lines.push(`${breakdown.unpricedServices.length} service(s) have no usable monthly estimate and are **excluded** from the subtotal above:`);
    lines.push('');
    breakdown.unpricedServices.forEach(svc => {
      lines.push(`- ${escapeMd(svc.serviceName)}`);
    });
    lines.push('');
  }
  lines.push('_Estimates are indicative. Usage-based services (e.g. Functions, OpenAI) may vary with actual consumption. Generated by Microsoft Product Architecture Diagram Builder._');

  return lines.join('\n');
}

/**
 * Escape characters that would break Markdown table cells.
 */
function escapeMd(value: string): string {
  return value.replace(/\|/g, '\\|');
}

/**
 * Export cost breakdown as CSV
 */
export function exportCostBreakdownCSV(breakdown: CostBreakdown, nodes?: Node[]): string {
  const lines: string[] = [];
  
  // Header
  lines.push('Azure Architecture Cost Breakdown');
  lines.push(`${breakdown.estimateCompleteness && breakdown.estimateCompleteness !== 'complete'
    ? 'Priced Monthly Subtotal' : 'Total Monthly Cost'},$${breakdown.totalMonthlyCost.toFixed(2)}`);
  if (breakdown.estimateCompleteness) lines.push(`Estimate Completeness,${breakdown.estimateCompleteness}`);
  if (breakdown.estimateCompleteness === 'unpriced') lines.push('Warning,No usable total estimate; unknown prices are not free');
  if (breakdown.missingCapacityEstimate) lines.push('Warning,Shared capacity estimate is missing');
  if (breakdown.unpricedServices?.length) {
    lines.push(`Unpriced Services Excluded,${breakdown.unpricedServices.length}`);
    lines.push('Warning,Partial subtotal (unpriced does not mean free)');
  }
  lines.push(`Region,${csvTextCell(breakdown.region)}`);
  if (breakdown.pricingTerm) lines.push(`Pricing Term,${csvTextCell(breakdown.pricingTerm)}`);
  if (breakdown.pricesAsOf) lines.push(`Prices As Of,${csvTextCell(breakdown.pricesAsOf)}`);
  if (breakdown.oldestMeterAsOf) lines.push(`Oldest Unchanged Price,${csvTextCell(breakdown.oldestMeterAsOf)}`);
  lines.push(`Date,${csvTextCell(new Date(breakdown.lastCalculated).toLocaleDateString())}`);
  lines.push('');
  
  // By Service
  lines.push('Service Name,Service Type,Tier,Quantity,Monthly Cost,Pricing Type');
  breakdown.byService.forEach(svc => {
    // Check if this service is usage-based
    const node = nodes?.find(n => n.id === svc.nodeId);
    const pricing = node?.data?.pricing as NodePricingConfig | undefined;
    const pricingType = pricing?.isUsageBased ? 'Usage-based (estimate)' : 'Fixed';
    
    lines.push(`${csvTextCell(svc.serviceName, true)},${csvTextCell(svc.serviceType)},${csvTextCell(svc.tier)},${svc.quantity},$${svc.cost.toFixed(2)},${csvTextCell(pricingType)}`);
  });
  lines.push('');
  
  // By Group
  lines.push('Group Name,Service Count,Monthly Cost');
  breakdown.byGroup.forEach(grp => {
    lines.push(`${csvTextCell(grp.groupLabel, true)},${grp.serviceCount},$${grp.cost.toFixed(2)}`);
  });
  lines.push('');
  
  // By Category
  lines.push('Category,Monthly Cost,Percentage');
  breakdown.byCategory.forEach(cat => {
    lines.push(`${csvTextCell(cat.category)},$${cat.cost.toFixed(2)},${cat.percentage.toFixed(1)}%`);
  });
  
  return lines.join('\n');
}

/**
 * Export cost breakdown as JSON
 */
export function exportCostBreakdownJSON(breakdown: CostBreakdown): string {
  return JSON.stringify(breakdown, null, 2);
}
