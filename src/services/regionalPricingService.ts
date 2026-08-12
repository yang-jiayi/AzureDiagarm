// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

/**
 * Regional Pricing Service
 * Manages loading and querying pricing data for different Azure regions
 */

import { AzureRetailPrice, ServicePricing, PricingTier } from '../types/pricing';

export type AzureRegion = 'eastus2' | 'swedencentral' | 'westeurope' | 'canadacentral' | 'brazilsouth' | 'australiaeast' | 'southeastasia' | 'mexicocentral' | 'japaneast';

export type RegionType = 'HERO' | 'HUB' | 'SATELLITE' | 'MICRO';

export interface RegionInfo {
  id: AzureRegion;
  displayName: string;
  location: string;
  flag: string;
  regionType: RegionType;
  geography: string;
}

export const AVAILABLE_REGIONS: RegionInfo[] = [
  { id: 'eastus2',       displayName: 'East US 2',        location: 'Virginia',    flag: '🇺🇸', regionType: 'HERO', geography: 'United States' },
  { id: 'australiaeast', displayName: 'Australia East',   location: 'Sydney',      flag: '🇦🇺', regionType: 'HERO', geography: 'Australia' },
  { id: 'canadacentral', displayName: 'Canada Central',   location: 'Toronto',     flag: '🇨🇦', regionType: 'HUB',  geography: 'Canada' },
  { id: 'brazilsouth',   displayName: 'Brazil South',     location: 'São Paulo',   flag: '🇧🇷', regionType: 'HUB',  geography: 'Brazil' },
  { id: 'mexicocentral', displayName: 'Mexico Central',   location: 'Querétaro',   flag: '🇲🇽', regionType: 'HUB',  geography: 'Mexico' },
  { id: 'westeurope',    displayName: 'West Europe',      location: 'Netherlands', flag: '🇳🇱', regionType: 'HUB',  geography: 'Europe' },
  { id: 'swedencentral', displayName: 'Sweden Central',   location: 'Gävle',       flag: '🇸🇪', regionType: 'HUB',  geography: 'Europe' },
  { id: 'southeastasia', displayName: 'Southeast Asia',   location: 'Singapore',   flag: '🇸🇬', regionType: 'HUB',  geography: 'Asia Pacific' },
  { id: 'japaneast',     displayName: 'Japan East',       location: 'Tokyo',       flag: '🇯🇵', regionType: 'HERO', geography: 'Japan' },
];

interface RegionalPricingData {
  BillingCurrency: string;
  Items: AzureRetailPrice[];
  /**
   * The newest meter effective date in the file this data came from, which is
   * how old the prices actually are. Not the date the file was downloaded: the
   * corpus spans 2018-02 to 2026-07, so a refresh run in 2026 still ships 2018
   * meters for the services Azure has not repriced since.
   */
  pricesAsOf?: string;
}

/** Compacted on-disk shape produced by scripts/prep-pricing-data.mjs. */
interface CompactPricingData {
  BillingCurrency?: string;
  ServiceName?: string;
  PricesAsOf?: string;
  Items?: Array<Partial<AzureRetailPrice> & Record<string, unknown>>;
}

// Small LRU so the pricing caches never retain every parsed dataset forever.
class LruCache<K, V> {
  private readonly map = new Map<K, V>();
  constructor(private readonly max: number) {}
  get(key: K): V | undefined {
    const value = this.map.get(key);
    if (value !== undefined) {
      // Refresh recency.
      this.map.delete(key);
      this.map.set(key, value);
    }
    return value;
  }
  has(key: K): boolean {
    return this.map.has(key);
  }
  set(key: K, value: V): void {
    if (this.map.has(key)) this.map.delete(key);
    this.map.set(key, value);
    while (this.map.size > this.max) {
      const oldest = this.map.keys().next().value as K | undefined;
      if (oldest === undefined) break;
      this.map.delete(oldest);
    }
  }
  clear(): void {
    this.map.clear();
  }
  entries(): IterableIterator<[K, V]> {
    return this.map.entries();
  }
  get size(): number {
    return this.map.size;
  }
}

// Cache of expanded raw file data, keyed by `${region}/${fileKey}` so a file
// shared by several services (e.g. Foundry tools) is fetched once.
const rawFileCache = new LruCache<string, RegionalPricingData>(24);

// Cache for parsed service pricing (the hot path reused by callers).
const parsedPricingCache = new LruCache<string, ServicePricing>(128);

// Current active region
let currentRegion: AzureRegion = 'japaneast';

/**
 * Base URL for the static pricing assets under `public/pricing/regions/**`.
 * The Vite BASE_URL define is a literal replaced at build time with the
 * configured `base` (defaults to '/'), so pricing loads correctly whether the
 * app is served from root or a sub-path. The try/catch keeps the module safe to
 * import in plain Node (unit tests), where that define is not injected.
 */
function pricingBaseUrl(): string {
  let base = '/';
  try {
    base = import.meta.env.BASE_URL || '/';
  } catch {
    base = '/';
  }
  if (!base.endsWith('/')) base += '/';
  return `${base}pricing/regions`;
}

/**
 * Restore a compacted pricing file to the shape the parser expects. Mirrors
 * scripts/prep-pricing-data.mjs `expandPricingData` — only fields the runtime
 * reads are reconstructed. Exported for the round-trip test.
 */export function expandPricingData(compact: CompactPricingData): RegionalPricingData {
  const serviceName = compact?.ServiceName;
  const items = Array.isArray(compact?.Items) ? compact.Items : [];
  const expanded = items.map((item) => {
    const next: Record<string, unknown> = { ...item };
    if (next.type === undefined) next.type = 'Consumption';
    if (next.serviceName === undefined && serviceName !== undefined) {
      next.serviceName = serviceName;
    }
    if (next.unitPrice === undefined) next.unitPrice = next.retailPrice;
    if (next.armSkuName === undefined) next.armSkuName = '';
    if (next.productName === undefined) next.productName = '';
    if (Array.isArray(next.savingsPlan)) {
      next.savingsPlan = (next.savingsPlan as Array<Record<string, unknown>>).map((plan) => (
        plan.unitPrice === undefined ? { ...plan, unitPrice: plan.retailPrice } : plan
      ));
    }
    return next as unknown as AzureRetailPrice;
  });
  return {
    BillingCurrency: compact?.BillingCurrency ?? 'USD',
    Items: expanded,
    pricesAsOf: typeof compact?.PricesAsOf === 'string' && ISO_DATE.test(compact.PricesAsOf)
      ? compact.PricesAsOf
      : undefined,
  };
}

/**
 * Fetch a single regional pricing file, expanding it to the parser shape.
 * Degrades gracefully: any network error or non-OK status resolves to null so
 * pricing simply reads as "unavailable" instead of crashing the app.
 */
async function fetchRegionalFile(region: AzureRegion, fileKey: string): Promise<RegionalPricingData | null> {
  const url = `${pricingBaseUrl()}/${region}/${fileKey}.json`;
  try {
    const response = await fetch(url);
    if (!response.ok) {
      console.warn(`⚠️ Pricing data unavailable (${response.status}) at ${url}`);
      return null;
    }
    const compact = (await response.json()) as CompactPricingData;
    return expandPricingData(compact);
  } catch (error) {
    console.warn(`⚠️ Failed to fetch pricing data at ${url}:`, error);
    return null;
  }
}

/** Fetch a raw file with LRU caching, deduped across services by file key. */
async function getRawFile(region: AzureRegion, fileKey: string): Promise<RegionalPricingData | null> {
  const cacheKey = `${region}/${fileKey}`;
  const cached = rawFileCache.get(cacheKey);
  if (cached) return cached;
  const data = await fetchRegionalFile(region, fileKey);
  if (data) rawFileCache.set(cacheKey, data);
  return data;
}

/**
 * Map AI service display names to Foundry productNames
 */
const AI_SERVICE_PRODUCT_MAP: Record<string, { file: string; productName: string; defaultSku?: string }> = {
  'Azure OpenAI': { file: 'foundry_models', productName: 'Azure OpenAI', defaultSku: 'gpt4omini' },
  'OpenAI': { file: 'foundry_models', productName: 'Azure OpenAI', defaultSku: 'gpt4omini' },
  'Azure AI Document Intelligence': { file: 'foundry_tools', productName: 'Azure Document Intelligence', defaultSku: 'Standard' },
  'Azure Document Intelligence': { file: 'foundry_tools', productName: 'Azure Document Intelligence', defaultSku: 'Standard' },
  'Document Intelligence': { file: 'foundry_tools', productName: 'Azure Document Intelligence', defaultSku: 'Standard' },
  'Form Recognizer': { file: 'foundry_tools', productName: 'Form Recognizer', defaultSku: 'Standard' },
  'Language': { file: 'foundry_tools', productName: 'Azure Language', defaultSku: 'Standard' },
  'Text Analytics': { file: 'foundry_tools', productName: 'Azure Language', defaultSku: 'Standard' },
  'Speech': { file: 'foundry_tools', productName: 'Azure Speech', defaultSku: 'Standard' },
  'Speech Services': { file: 'foundry_tools', productName: 'Azure Speech', defaultSku: 'Standard' },
  'Vision': { file: 'foundry_tools', productName: 'Azure Vision', defaultSku: 'Standard' },
  'Computer Vision': { file: 'foundry_tools', productName: 'Azure Vision', defaultSku: 'Standard' },
  'Face': { file: 'foundry_tools', productName: 'Azure Vision - Face', defaultSku: 'Standard' },
  'Translator': { file: 'foundry_tools', productName: 'Azure Translator', defaultSku: 'Standard' },
  'Custom Vision': { file: 'foundry_tools', productName: 'Azure Custom Vision', defaultSku: 'Standard' },
  'Content Safety': { file: 'foundry_tools', productName: 'Content Safety', defaultSku: 'Standard' },
};

/**
 * Check if a service is an AI service that needs Foundry data
 */
function isAIService(serviceName: string): boolean {
  return AI_SERVICE_PRODUCT_MAP.hasOwnProperty(serviceName);
}

// ── Microsoft Fabric (region-aware) ─────────────────────────────────────────
// Fabric is licensed by Capacity (F-SKUs) and OneLake storage is billed per GB.
// Both vary slightly by region, so we read the true per-region rates from the
// fetched microsoft_fabric.json instead of the static fallback ladder.

function isFabricCapacityService(name: string): boolean {
  return name === 'Microsoft Fabric Capacity';
}

function isOneLakeService(name: string): boolean {
  return name === 'OneLake' || name === 'OneLake Storage';
}

/** Most common value in a list (mode), or a default when empty. */
function modeOrDefault(values: number[], fallback: number): number {
  if (values.length === 0) return fallback;
  const counts = new Map<number, number>();
  for (const v of values) counts.set(v, (counts.get(v) || 0) + 1);
  let best = fallback;
  let bestCount = -1;
  for (const [v, c] of counts) {
    if (c > bestCount) { bestCount = c; best = v; }
  }
  return best;
}

/**
 * Build per-region Fabric pricing from the fetched microsoft_fabric.json.
 * - Capacity: F-SKU monthly = (per-CU-hour rate) × CUs × 730 hours.
 *   The per-CU-hour rate is the mode of the "Capacity Usage CU" consumption
 *   meters (≈ $0.18, with small regional variance).
 * - OneLake: uses the "OneLake Storage Hot Data Stored" per-GB meter.
 */
async function getFabricRegionalPricing(
  serviceName: string,
  region: AzureRegion
): Promise<ServicePricing | null> {
  const data = await getRawFile(region, 'microsoft_fabric');
  if (!data) {
    console.warn(`⚠️ No Fabric pricing data available for ${region}`);
    return null;
  }

  if (isFabricCapacityService(serviceName)) {
    const rates = data.Items
      .filter(i => i.type === 'Consumption'
        && (i as any).unitOfMeasure === '1 Hour'
        && /Capacity Usage CU/i.test(i.meterName))
      .map(i => i.retailPrice || i.unitPrice)
      .filter(r => r > 0);
    const rate = modeOrDefault(rates, 0.18);
    const skus: Array<[string, number]> = [['F2', 2], ['F8', 8], ['F64', 64]];
    const tiers: PricingTier[] = skus.map(([name, cu]) => ({
      name,
      skuName: name,
      monthlyPrice: parseFloat((rate * cu * 730).toFixed(2)),
      hourlyPrice: parseFloat((rate * cu).toFixed(4)),
      unit: 'per capacity/month',
      description: `${name} — ${cu} CU @ $${rate}/CU-hour (${region})`
    }));
    return {
      serviceType: serviceName,
      serviceName,
      defaultTier: 'F2',
      tiers,
      calculationType: 'hourly',
      lastUpdated: new Date().toISOString(),
    };
  }

  if (isOneLakeService(serviceName)) {
    const hot = data.Items.find(i =>
      i.type === 'Consumption' && /OneLake Storage Hot Data Stored/i.test(i.meterName));
    const perGB = (hot?.retailPrice ?? hot?.unitPrice) || 0.023;
    const sizes: Array<[string, number]> = [['~200 GB', 200], ['~1 TB', 1000], ['~10 TB', 10000]];
    const tiers: PricingTier[] = sizes.map(([name, gb]) => ({
      name,
      skuName: name,
      monthlyPrice: parseFloat((perGB * gb).toFixed(2)),
      hourlyPrice: perGB,
      unit: 'per month (storage)',
      description: `${gb} GB Hot @ $${perGB}/GB (${region})`
    }));
    return {
      serviceType: serviceName,
      serviceName,
      defaultTier: '~1 TB',
      tiers,
      calculationType: 'usage',
      lastUpdated: new Date().toISOString(),
    };
  }

  return null;
}

/**
 * Set the active region for pricing queries
 */
export function setActiveRegion(region: AzureRegion): void {
  console.log(`🌍 Switching pricing region to: ${region}`);
  currentRegion = region;
  // Clear parsed pricing cache when region changes
  parsedPricingCache.clear();
}

/**
 * Get the current active region
 */
export function getActiveRegion(): AzureRegion {
  return currentRegion;
}

/**
 * Get region display info
 */
export function getRegionInfo(region: AzureRegion): RegionInfo | undefined {
  return AVAILABLE_REGIONS.find(r => r.id === region);
}

/**
 * ISO calendar date, which is the only form `PricesAsOf` is written in.
 *
 * The vintage used to live in a module-level registry keyed by service name,
 * written as files were parsed and read back when the breakdown was assembled.
 * That was wrong in four separate ways, all of them the same mistake — the date
 * was not attached to the number it described:
 *
 *  - it was recorded under the *mapped* Azure name and read back under the raw
 *    canvas name, so it silently answered "nothing" for the ~half of the
 *    catalogue where those differ;
 *  - it was oldest-wins *across regions*, so switching region (or using the
 *    built-in nine-region comparison) permanently attributed one region's
 *    vintage to another's prices — the one direction that over-claims;
 *  - it was recorded when a file *loaded*, not when a meter was *used*, so a
 *    node priced from the static fallback table still got an attestation that
 *    Azure had set that price on some date;
 *  - it lived only in session state, so restoring a saved diagram and exporting
 *    gave a different answer from exporting after any later pricing fetch.
 *
 * The date now rides on the pricing object and is stamped onto the node, so it
 * is bound to its own number, its own region, and its own provenance.
 */
const ISO_DATE = /^\d{4}-\d{2}-\d{2}$/;

/**
 * Load pricing data for a specific service in a region
 */
async function loadServiceData(region: AzureRegion, serviceName: string): Promise<RegionalPricingData | null> {
  // AI services share Foundry files and are filtered by productName.
  if (isAIService(serviceName)) {
    const aiMapping = AI_SERVICE_PRODUCT_MAP[serviceName];
    const fullData = await getRawFile(region, aiMapping.file);
    if (!fullData) {
      console.warn(`⚠️ No pricing data available for ${serviceName} (${aiMapping.file}) in ${region}`);
      return null;
    }
    const filteredItems = fullData.Items.filter(item =>
      (item as any).productName === aiMapping.productName
    );
    return {
      BillingCurrency: fullData.BillingCurrency,
      Items: filteredItems,
      pricesAsOf: fullData.pricesAsOf,
    };
  }

  // Regular service - load by filename derived from the service name.
  const filename = serviceName.toLowerCase().replace(/\s+/g, '_');
  const data = await getRawFile(region, filename);
  if (!data) {
    console.warn(`⚠️ No pricing data available for ${serviceName} in ${region}`);
    return null;
  }
  return data;
}

/**
 * Get available services for a region by checking which files exist
 */
export function getAvailableServices(_region: AzureRegion): string[] {
  // These are the services we have data for
  return [
    'Azure App Service',
    'Virtual Machines',
    'Azure Cosmos DB',
    'Storage',
    'SQL Database',
    'Azure Kubernetes Service',
    'Container Instances',
    'Application Gateway',
    'Azure Machine Learning',
    'Azure AI Search',
  ];
}

/**
 * Filter pricing items by service and region
 */
export function filterPricingItems(
  items: AzureRetailPrice[],
  serviceName: string,
  consumptionOnly: boolean = true
): AzureRetailPrice[] {
  const filtered = items.filter(item => {
    // Match service name (case insensitive)
    const matches = item.serviceName.toLowerCase() === serviceName.toLowerCase();
    if (!matches) return false;
    
    // Only consumption pricing (not reservations or spot)
    if (consumptionOnly && item.type !== 'Consumption') return false;

    // The Retail Prices API groups Static Web Apps, domains, and SSL add-ons
    // under the Azure App Service service name. They are separate products,
    // not hosting-plan SKUs, and must not appear in the App Service tier picker.
    if (
      serviceName.toLowerCase() === 'azure app service'
      && !/^Azure App Service (Free|Shared|Basic|Standard|Premium|Isolated)\b/i.test(item.productName)
    ) {
      return false;
    }
    
    return true;
  });
  
  console.log(`🔍 Filtered ${filtered.length} items for ${serviceName} from ${items.length} total`);
  return filtered;
}

/**
 * Parse pricing items into tiers
 */
export function parsePricingTiers(items: AzureRetailPrice[], serviceName: string): PricingTier[] {
  const tierMap = new Map<string, PricingTier>();
  const isAppService = serviceName.toLowerCase() === 'azure app service';

  // Convert a per-unit rate into a monthly cost given the meter's unit-of-measure.
  const toMonthly = (rate: number, unitOfMeasure: string): number => {
    const normalizedUnit = unitOfMeasure.toLowerCase();
    if (normalizedUnit.includes('/month')) return rate;
    if (normalizedUnit.includes('/year')) return rate / 12;
    if (normalizedUnit.includes('/day')) return rate * 30;
    if (normalizedUnit === '1k' || normalizedUnit.includes('1000')) return rate * 100;
    return rate * 730; // default: hourly × 730 hours/month
  };

  items.forEach(item => {
    const skuName = item.skuName || item.armSkuName;
    if (!skuName) return;
    const tierId = isAppService
      ? `${item.productName}::${skuName}`
      : skuName;
    const displayName = isAppService
      ? `${skuName} (${/- Linux$/i.test(item.productName) ? 'Linux' : 'Windows'})`
      : skuName;
    
    // Handle different billing units for AI services
    const unitOfMeasure = (item as any).unitOfMeasure || '1 Hour';
    const hourlyPrice = item.retailPrice || item.unitPrice;
    const monthlyPrice = toMonthly(hourlyPrice, unitOfMeasure);

    // Real 1-year Savings Plan monthly, when the meter carries a savings-plan rate.
    let reserved1yrMonthly: number | undefined;
    const oneYear = Array.isArray(item.savingsPlan)
      ? item.savingsPlan.find(p => /1\s*year/i.test(p.term || ''))
      : undefined;
    if (oneYear) {
      const spRate = oneYear.retailPrice || oneYear.unitPrice;
      if (spRate > 0) reserved1yrMonthly = toMonthly(spRate, unitOfMeasure);
    }

    // Only add if we don't have this SKU yet, or if this is cheaper
    if (!tierMap.has(tierId) || tierMap.get(tierId)!.monthlyPrice > monthlyPrice) {
      tierMap.set(tierId, {
        id: tierId,
        name: displayName,
        skuName: skuName,
        monthlyPrice: monthlyPrice,
        hourlyPrice: hourlyPrice,
        unit: item.unitOfMeasure,
        description: item.meterName,
        reserved1yrMonthly
      });
    }
  });
  
  const tiers = Array.from(tierMap.values()).sort((a, b) => a.monthlyPrice - b.monthlyPrice);
  console.log(`📊 Parsed ${tiers.length} pricing tiers. First few:`, tiers.slice(0, 3).map(t => ({ name: t.name, monthly: t.monthlyPrice })));
  return tiers;
}

/**
 * Get pricing for a service in the current active region
 */
export async function getRegionalServicePricing(
  serviceName: string,
  region?: AzureRegion
): Promise<ServicePricing | null> {
  const targetRegion = region || currentRegion;
  const cacheKey = `${serviceName}-${targetRegion}`;
  
  // Check cache
  if (parsedPricingCache.has(cacheKey)) {
    return parsedPricingCache.get(cacheKey)!;
  }
  
  console.log(`📊 Getting pricing from regional data for ${serviceName} in ${targetRegion}...`);
  
  // Microsoft Fabric is region-aware but parsed specially from microsoft_fabric.json
  if (isFabricCapacityService(serviceName) || isOneLakeService(serviceName)) {
    const fabricPricing = await getFabricRegionalPricing(serviceName, targetRegion);
    if (fabricPricing) {
      parsedPricingCache.set(cacheKey, fabricPricing);
      console.log(`✅ Loaded region-aware Fabric pricing for ${serviceName} in ${targetRegion}`);
      return fabricPricing;
    }
    // fall through to static fallback if the regional file is missing
    return null;
  }
  
  // Load service data for the region
  const data = await loadServiceData(targetRegion, serviceName);
  
  if (!data || data.Items.length === 0) {
    console.warn(`⚠️ No regional pricing data found for ${serviceName} in ${targetRegion}`);
    return null;
  }
  
  // Filter and parse the items
  const filteredItems = filterPricingItems(data.Items, serviceName);
  
  if (filteredItems.length === 0) {
    console.warn(`⚠️ No consumption pricing items for ${serviceName} in ${targetRegion}`);
    return null;
  }
  
  const tiers = parsePricingTiers(filteredItems, serviceName);
  
  if (tiers.length === 0) {
    console.warn(`⚠️ No pricing tiers parsed for ${serviceName} in ${targetRegion}`);
    return null;
  }
  
  console.log(`✅ Found ${tiers.length} tiers for ${serviceName} in ${targetRegion}`);
  
  const pricing: ServicePricing = {
    serviceType: serviceName,
    serviceName,
    defaultTier: tiers[0]?.name || 'Standard',
    tiers,
    calculationType: 'hourly',
    lastUpdated: new Date().toISOString(),
    // Travels with the tiers it describes, so whoever ends up printing one of
    // these numbers can say how long Azure has been charging it without
    // knowing which file, which region or which name it was found under.
    meterAsOf: data.pricesAsOf && ISO_DATE.test(data.pricesAsOf) ? data.pricesAsOf : undefined,
  };
  
  // Cache the result
  parsedPricingCache.set(cacheKey, pricing);
  
  return pricing;
}

/**
 * Get pricing summary for the current region
 */
export function getRegionalPricingSummary(region?: AzureRegion): {
  region: AzureRegion;
  servicesLoaded: number;
  totalItems: number;
  cacheSize: number;
} {
  const targetRegion = region || currentRegion;

  let servicesLoaded = 0;
  let totalItems = 0;
  const prefix = `${targetRegion}/`;
  for (const [key, data] of rawFileCache.entries()) {
    if (key.startsWith(prefix)) {
      servicesLoaded += 1;
      totalItems += data.Items.length;
    }
  }

  return {
    region: targetRegion,
    servicesLoaded,
    totalItems,
    cacheSize: parsedPricingCache.size,
  };
}

/**
 * Preload common services for faster initial pricing
 */
export async function preloadCommonServices(region?: AzureRegion): Promise<void> {
  const targetRegion = region || currentRegion;
  const commonServices = [
    'Azure App Service',
    'Virtual Machines',
    'Storage',
    'SQL Database',
    'Azure Cosmos DB',
  ];
  
  console.log(`⏳ Preloading ${commonServices.length} common services for ${targetRegion}...`);
  
  const promises = commonServices.map(service => loadServiceData(targetRegion, service));
  await Promise.all(promises);
  
  const summary = getRegionalPricingSummary(targetRegion);
  console.log(`✅ Preloaded ${summary.servicesLoaded} services (${summary.totalItems} items) for ${targetRegion}`);
}
