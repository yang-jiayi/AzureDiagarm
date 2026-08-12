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
  PricingTier,
  ServicePricing,
} from '../types/pricing';
import { 
  getServicePricing
} from './azurePricingService';
import { 
  getActiveRegion,
  oldestMeterVintage
} from './regionalPricingService';
import { 
  getAzureServiceName, 
  getDefaultTier, 
  getFallbackPricing,
  getFallbackDefaultLevel,
  getFallbackDefaultSku,
  getReserved1yrDiscount,
  PRICING_DATA_AS_OF,
  hasPricingData,
  USAGE_BASED_SERVICES 
} from '../data/azurePricing';
import { 
  applyRegionalPricing,
  getPricingFreshness
} from '../utils/pricingHelpers';
import { csvTextCell } from '../utils/csv';

/**
 * Initialize pricing for a new node
 */
export async function initializeNodePricing(
  serviceType: string,
  region?: string
): Promise<NodePricingConfig | null> {
  const targetRegion = region || getActiveRegion();
  console.log('🔍 Initializing pricing for:', serviceType, 'in region:', targetRegion);
  
  // Check if this service has pricing data
  if (!hasPricingData(serviceType)) {
    console.warn(`⚠️ No pricing data available for ${serviceType}`);
    return null;
  }

  try {
    // Get Azure service name
    const serviceName = getAzureServiceName(serviceType);
    const defaultTier = getDefaultTier(serviceType);
    console.log('  → Mapped to Azure service:', serviceName, 'Default tier:', defaultTier);
    
    // Check if this is a usage-based service (need this for all code paths)
    const isUsageBased = USAGE_BASED_SERVICES.includes(serviceType);
    
    // Try to fetch from API
    const pricing = await getServicePricing(serviceType, serviceName, targetRegion);
    
    if (pricing && pricing.tiers.length > 0) {
      // Use API data
      const tier = pricing.tiers.find(t =>
        t.id === defaultTier || t.skuName === defaultTier || t.name === defaultTier
      ) || pricing.tiers[0];
      console.log('  ✅ Found tier:', tier.name, 'Price:', tier.monthlyPrice, '/mo (hourly:', tier.hourlyPrice, ')');
      
      // If pricing is $0 (usage-based services like Storage), use fallback pricing
      if (tier.monthlyPrice === 0 || tier.monthlyPrice === null || tier.monthlyPrice === undefined) {
        console.log('  💡 Usage-based pricing ($0 base), using fallback estimate');
        const fallbackPrice = getFallbackPricing(serviceType, 'standard');
        if (fallbackPrice !== null && fallbackPrice > 0) {
          const basePrice = applyRegionalPricing(fallbackPrice, targetRegion);
          
          return {
            estimatedCost: basePrice,
            tier: tier.name,
            tierId: tier.id,
            skuName: tier.skuName,
            quantity: 1,
            region: targetRegion,
            unit: tier.unit,
            lastUpdated: new Date().toISOString(),
            isCustom: false,
            isUsageBased: true
          };
        }
      }
      
      return {
        estimatedCost: tier.monthlyPrice,
        tier: tier.name,
        tierId: tier.id,
        skuName: tier.skuName,
        quantity: 1,
        region: targetRegion,
        unit: tier.unit,
        lastUpdated: new Date().toISOString(),
        isCustom: false,
        isUsageBased: isUsageBased,
        reserved1yrCost: tier.reserved1yrMonthly,
        reservedIsSavingsPlan: tier.reserved1yrMonthly != null
      };
    } else {
      // Fallback to static data — use the service's default SKU/level
      const fallbackPrice = getFallbackPricing(serviceType, getFallbackDefaultLevel(serviceType));
      if (fallbackPrice === null) {
        // No API price and no static price: report "unknown" by returning null
        // rather than 0, which the badge would render as a green "Free".
        console.warn(`⚠️ No resolvable price for ${serviceType} — reporting as unavailable`);
        return null;
      }
      const basePrice = applyRegionalPricing(fallbackPrice, targetRegion);
      const skuLabel = getFallbackDefaultSku(serviceType);
      console.log('  💾 Using fallback pricing:', basePrice, '/mo');
      
      return {
        estimatedCost: basePrice,
        tier: skuLabel,
        skuName: skuLabel,
        quantity: 1,
        region: targetRegion,
        unit: 'per instance/month',
        lastUpdated: new Date().toISOString(),
        isCustom: false,
        isUsageBased: isUsageBased
      };
    }
  } catch (error) {
    console.error(`Error initializing pricing for ${serviceType}:`, error);
    
    // Final fallback
    const fallbackPrice = getFallbackPricing(serviceType, getFallbackDefaultLevel(serviceType));
    if (fallbackPrice === null) return null;
    const basePrice = applyRegionalPricing(fallbackPrice, targetRegion);
    const skuLabel = getFallbackDefaultSku(serviceType);
    const isUsageBased = USAGE_BASED_SERVICES.includes(serviceType);
    
    return {
      estimatedCost: basePrice,
      tier: skuLabel,
      skuName: skuLabel,
      quantity: 1,
      region: targetRegion,
      unit: 'per instance/month',
      lastUpdated: new Date().toISOString(),
      isCustom: false,
      isUsageBased: isUsageBased
    };
  }
}

/**
 * Update pricing when tier or quantity changes
 */
function pricesMatch(left: number, right: number): boolean {
  return Math.abs(left - right) <= Math.max(0.01, Math.abs(right) * 0.001);
}

function findTierCandidates(pricing: ServicePricing, selector: string): PricingTier[] {
  const exactIdMatch = pricing.tiers.find(candidate => candidate.id === selector);
  if (exactIdMatch) return [exactIdMatch];

  const exactNameMatches = pricing.tiers.filter(candidate => candidate.name === selector);
  if (exactNameMatches.length > 0) return exactNameMatches;

  return pricing.tiers.filter(candidate => candidate.skuName === selector);
}

async function resolvePricingTier(
  serviceType: string,
  serviceName: string,
  currentConfig: NodePricingConfig,
  selector: string,
  region: string,
  pricing: ServicePricing,
): Promise<PricingTier> {
  const matchingTiers = findTierCandidates(pricing, selector);
  if (matchingTiers.length === 0) {
    throw new Error(`Tier ${selector} is not available for ${serviceType} in ${region}`);
  }
  if (matchingTiers.length === 1) return matchingTiers[0];

  // New App Service tiers have platform-specific ids, but older saved diagrams
  // only stored a shared SKU such as "S1". Recover the original platform from
  // its source-region unit price, then select that same product in the target
  // region. If the old estimate cannot identify one platform uniquely, keep
  // treating it as ambiguous instead of silently changing the workload.
  const namedTier = matchingTiers.filter(candidate => candidate.name === currentConfig.tier);
  if (namedTier.length === 1) return namedTier[0];

  const sourcePricing = currentConfig.region === region
    ? pricing
    : await getServicePricing(serviceType, serviceName, currentConfig.region);
  if (sourcePricing) {
    const sourceCandidates = findTierCandidates(sourcePricing, selector);
    const sourcePriceMatches = sourceCandidates.filter(candidate =>
      pricesMatch(candidate.monthlyPrice, currentConfig.estimatedCost)
    );
    if (sourcePriceMatches.length === 1 && sourcePriceMatches[0].id) {
      const targetTier = pricing.tiers.find(candidate => candidate.id === sourcePriceMatches[0].id);
      if (targetTier) return targetTier;
    }
  }

  throw new Error(`Tier ${selector} is ambiguous for ${serviceType}; select a platform-specific SKU`);
}

export async function updateNodePricing(
  serviceType: string,
  currentConfig: NodePricingConfig,
  newTier?: string,
  newQuantity?: number,
  newRegion?: string
): Promise<NodePricingConfig> {
  const tier = newTier || currentConfig.tierId || currentConfig.skuName || currentConfig.tier;
  const quantity = newQuantity ?? currentConfig.quantity;
  const region = newRegion || currentConfig.region;
  
  try {
    const serviceName = getAzureServiceName(serviceType);
    const pricing = await getServicePricing(serviceType, serviceName, region);
    
    if (pricing) {
      const selectedTier = await resolvePricingTier(
        serviceType,
        serviceName,
        currentConfig,
        tier,
        region,
        pricing,
      );
      const isUsageBased = currentConfig.isUsageBased === true
        || USAGE_BASED_SERVICES.includes(serviceType);
      let estimatedCost = selectedTier.monthlyPrice;
      if (isUsageBased && estimatedCost <= 0) {
        const fallbackPrice = getFallbackPricing(serviceType, 'standard');
        estimatedCost = fallbackPrice !== null && fallbackPrice > 0
          ? applyRegionalPricing(fallbackPrice, region)
          : currentConfig.estimatedCost;
      }

      // `estimatedCost` is PER UNIT everywhere else in the system —
      // calculateCostBreakdown and AzureNode both multiply it by quantity
      // themselves. Storing a quantity-multiplied total here would be counted
      // twice (quantity squared). Price one unit and let the callers scale it.
      return {
        ...currentConfig,
        estimatedCost,
        tier: selectedTier.name,
        tierId: selectedTier.id,
        skuName: selectedTier.skuName,
        quantity,
        region,
        lastUpdated: new Date().toISOString(),
        isCustom: false,
        customPrice: undefined,
        isUsageBased,
        reserved1yrCost: isUsageBased ? undefined : selectedTier.reserved1yrMonthly,
        reservedIsSavingsPlan: !isUsageBased && selectedTier.reserved1yrMonthly != null
      };
    } else {
      throw new Error(`Pricing data is unavailable for ${serviceType} in ${region}`);
    }
  } catch (error) {
    console.error(`Error updating pricing for ${serviceType}:`, error);
    throw error;
  }
}

/**
 * Set custom pricing for a node
 */
export function setCustomPricing(
  currentConfig: NodePricingConfig,
  customPrice: number
): NodePricingConfig {
  if (!Number.isFinite(customPrice) || customPrice < 0) {
    throw new Error('Custom pricing must be a finite number of zero or more');
  }
  return {
    ...currentConfig,
    estimatedCost: customPrice,
    customPrice: customPrice,
    isCustom: true,
    reserved1yrCost: undefined,
    reservedIsSavingsPlan: false,
    lastUpdated: new Date().toISOString()
  };
}

/**
 * Tiers/SKUs a service can be switched to, for the per-node cost editor.
 *
 * Estimates otherwise sit on two fixed assumptions — the catalog default tier
 * and quantity 1 — which users pushed back on. Returns [] when the service has
 * no catalog pricing (usage-based services, or anything with hasPricingData
 * false), in which case only a custom override is meaningful.
 */
export async function getAvailableTiers(
  serviceType: string,
  region?: string
): Promise<PricingTier[]> {
  const targetRegion = region || getActiveRegion();
  if (!hasPricingData(serviceType)) return [];
  try {
    const serviceName = getAzureServiceName(serviceType);
    const pricing = await getServicePricing(serviceType, serviceName, targetRegion);
    return pricing?.tiers ?? [];
  } catch (error) {
    console.error(`Error loading tiers for ${serviceType}:`, error);
    return [];
  }
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
 * How long a price must have held before its age is worth reporting. Azure
 * reprices most meters well inside a year, so anything that survives this has
 * been genuinely stable rather than merely not-refreshed-this-week.
 */
const STABLE_PRICE_DAYS = 365;

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
  };

  // Track costs by group and category
  const groupCosts = new Map<string, { label: string; cost: number; count: number }>();
  const categoryCosts = new Map<string, number>();
  const pricingRegions = new Set<string>();
  // Which catalogue services the estimate actually priced, so the vintage
  // reported is the vintage of the numbers on this page rather than of the
  // whole dataset. Node labels are user-editable, so the lookup key is the
  // service name the pricing loader was given.
  const pricedServiceNames = new Set<string>();

  // Calculate per-service costs
  const unpricedServices: { nodeId: string; serviceName: string }[] = [];
  nodes.forEach(node => {
    const pricing = node.data.pricing as NodePricingConfig | undefined;
    
    if (!pricing) {
      const serviceName = String(node.data.serviceName || node.data.label || '').trim();
      if (serviceName) unpricedServices.push({ nodeId: node.id, serviceName });
      return;
    }
    const pricingRegion = typeof pricing.region === 'string' ? pricing.region.trim() : '';
    pricingRegions.add(pricingRegion || 'Unknown');
    // A custom price is a number the user typed, so no Azure meter stands
    // behind it and it must not drag the reported vintage backwards.
    if (!pricing.isCustom) {
      const priced = String(node.data.serviceName || node.data.label || '').trim();
      if (priced) pricedServiceNames.add(priced);
    }

    let cost = pricing.estimatedCost * pricing.quantity;
    // Apply the 1-year commitment to reservation-eligible, non-usage-based
    // catalog prices. Custom prices are user-provided final monthly amounts,
    // so applying another generic discount would understate the estimate.
    if (pricingMode === 'reserved1yr' && !pricing.isUsageBased && !pricing.isCustom) {
      if (pricing.reserved1yrCost != null && pricing.reserved1yrCost > 0) {
        cost = pricing.reserved1yrCost * pricing.quantity;
      } else {
        const serviceType = String(node.data.serviceName || node.data.label || '');
        const discount = getReserved1yrDiscount(serviceType);
        if (discount > 0) cost = cost * (1 - discount);
      }
    }
    breakdown.totalMonthlyCost += cost;

    // Add to service breakdown
    breakdown.byService.push({
      serviceName: node.data.label || 'Unnamed Service',
      serviceType: node.data.category || 'Other',
      nodeId: node.id,
      cost: cost,
      quantity: pricing.quantity,
      tier: pricing.tier
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

  if (unpricedServices.length > 0) breakdown.unpricedServices = unpricedServices;

  // Only worth saying when the prices have actually held for a while. Every
  // meter predates the download by some margin, so reporting any gap at all
  // would put a second date on every slide that never means anything; a year
  // is well past Azure's normal repricing cadence, so what survives it is the
  // genuinely long-stable pricing worth mentioning out loud.
  const oldest = oldestMeterVintage([...pricedServiceNames]);
  if (oldest && ageInDays(oldest, PRICING_DATA_AS_OF) > STABLE_PRICE_DAYS) {
    breakdown.oldestMeterAsOf = oldest;
  }

  return breakdown;
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
    if (node.data.pricing) {
      const serviceType = String(node.data.serviceName || node.data.label || 'Unknown');
      const currentPricing = node.data.pricing as NodePricingConfig;
      const updatedPricing = currentPricing.isCustom
        ? {
            ...currentPricing,
            region: newRegion,
            lastUpdated: new Date().toISOString(),
          }
        : await updateNodePricing(
            serviceType,
            currentPricing,
            currentPricing.tierId || currentPricing.skuName || currentPricing.tier,
            currentPricing.quantity,
            newRegion
          );

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
  lines.push(`Total Monthly Cost: $${breakdown.totalMonthlyCost.toFixed(2)}`);
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
    lines.push(`  ${svc.serviceName} (${svc.tier}): $${svc.cost.toFixed(2)}/mo x${svc.quantity}`);
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
    lines.push(`NOT INCLUDED IN THE TOTAL (${breakdown.unpricedServices.length} service(s) without published pricing):`);
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
  lines.push(`> **Total: \`$${breakdown.totalMonthlyCost.toFixed(2)}/mo\`** · **\`$${annual.toFixed(2)}/yr\`** · Region: \`${breakdown.region}\` · ${breakdown.currency}${breakdown.pricingTerm ? ` · ${breakdown.pricingTerm}` : ''}`);
  lines.push('');
  lines.push('| Field | Value |');
  lines.push('| --- | --- |');
  lines.push(`| Total monthly cost | **$${breakdown.totalMonthlyCost.toFixed(2)}** |`);
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
  lines.push(`| **Total** | | | **$${breakdown.totalMonthlyCost.toFixed(2)}** |`);
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
    lines.push(`${breakdown.unpricedServices.length} service(s) have no published pricing in this build and are **excluded** from the total above:`);
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
  lines.push(`Total Monthly Cost,$${breakdown.totalMonthlyCost.toFixed(2)}`);
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
