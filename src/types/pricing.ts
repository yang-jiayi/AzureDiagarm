// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

/**
 * Azure Pricing Types
 * Based on Azure Retail Prices API
 * https://learn.microsoft.com/en-us/rest/api/cost-management/retail-prices/azure-retail-prices
 */

/**
 * Response from Azure Retail Prices API
 */
export interface AzureRetailPricesResponse {
  BillingCurrency: string;
  CustomerEntityId: string;
  CustomerEntityType: string;
  Items: AzureRetailPrice[];
  NextPageLink: string | null;
  Count: number;
}

/**
 * Individual price item from Azure Retail Prices API
 */
/**
 * A commitment-based rate embedded on a Consumption meter by the Azure Retail
 * Prices API. `term` is "1 Year" or "3 Years"; the price is the discounted
 * hourly/usage rate for that Savings Plan for Compute commitment.
 */
export interface SavingsPlanRate {
  term: string;         // e.g. "1 Year", "3 Years"
  retailPrice: number;
  unitPrice: number;
}

export interface AzureRetailPrice {
  currencyCode: string;
  tierMinimumUnits: number;
  retailPrice: number;
  unitPrice: number;
  armRegionName: string;
  location: string;
  effectiveStartDate: string;
  meterId: string;
  meterName: string;
  productId: string;
  skuId: string;
  productName: string;
  skuName: string;
  serviceName: string;
  serviceId: string;
  serviceFamily: string;
  unitOfMeasure: string;
  type: string;
  isPrimaryMeterRegion: boolean;
  armSkuName: string;
  /** 1-year / 3-year Savings Plan for Compute rates, when the meter carries them. */
  savingsPlan?: SavingsPlanRate[];
}

/**
 * Simplified pricing tier for our application
 */
export interface PricingTier {
  /** Stable selector when the same Azure SKU name exists in multiple products or platforms. */
  id?: string;
  name: string;           // e.g., "B1 (Basic)", "S1 (Standard)", "P1V2 (Premium)"
  skuName: string;        // e.g., "B1", "S1", "P1V2"
  monthlyPrice: number;   // Estimated monthly cost in USD
  hourlyPrice?: number;   // Hourly rate if available
  unit: string;           // e.g., "per instance", "per GB", "per hour"
  description?: string;   // Brief description of tier
  /**
   * Real 1-year Savings Plan monthly cost for this SKU, derived from the
   * meter's embedded savingsPlan[] "1 Year" rate. Undefined when the meter
   * carries no savings-plan rate (then the cost engine falls back to a
   * representative discount percentage).
   */
  reserved1yrMonthly?: number;
}

/**
 * Service pricing information
 */
export interface ServicePricing {
  serviceType: string;                    // e.g., "App Service", "Cosmos DB"
  serviceName: string;                    // Azure service name from API
  defaultTier: string;                    // Default tier to use (e.g., "Standard")
  tiers: PricingTier[];                   // Available tiers
  calculationType: 'hourly' | 'monthly' | 'usage'; // How to calculate cost
  sourceUrl?: string;                     // Link to official pricing page
  lastUpdated: string;                    // ISO timestamp
}

/**
 * Pricing configuration stored in node data
 */
export interface NodePricingConfig {
  estimatedCost: number;      // Monthly cost in USD
  tier: string;               // Selected tier name
  tierId?: string;            // Stable catalog selector for otherwise ambiguous SKU names
  skuName: string;            // SKU identifier
  quantity: number;           // Number of instances/units
  region: string;             // Azure region
  unit: string;               // Unit of measurement
  lastUpdated: string;        // ISO timestamp
  isCustom: boolean;          // Whether user manually set price
  customPrice?: number;       // Custom monthly price if set
  isUsageBased?: boolean;     // Whether pricing is usage-based (consumption)
  /**
   * Real 1-year Savings Plan monthly cost (per unit) for this SKU, when the
   * meter carried a savings-plan rate. Used for the reserved-term estimate in
   * preference to the representative discount table.
   */
  reserved1yrCost?: number;
  /** True when reserved1yrCost came from a real savings-plan meter (not the % fallback). */
  reservedIsSavingsPlan?: boolean;
  usageEstimate?: {           // For usage-based services
    type: 'light' | 'medium' | 'heavy';
    description: string;
  };
}

/**
 * Regional pricing multiplier
 */
export interface RegionPricing {
  region: string;
  displayName: string;
  armRegionName: string;
  multiplier: number;
}

/**
 * Cost breakdown for estimation panel
 */
export interface CostBreakdown {
  totalMonthlyCost: number;
  byService: {
    serviceName: string;
    serviceType: string;
    nodeId: string;
    cost: number;
    quantity: number;
    tier: string;
  }[];
  byGroup: {
    groupId: string;
    groupLabel: string;
    cost: number;
    serviceCount: number;
  }[];
  byCategory: {
    category: string;
    cost: number;
    percentage: number;
  }[];
  region: string;
  currency: string;
  lastCalculated: string;
  /**
   * Azure services on the canvas that carry no resolvable price, so the total
   * above excludes them. Surfaced in the UI and the summary export so a
   * partial estimate is never presented as a complete one.
   */
  unpricedServices?: { nodeId: string; serviceName: string }[];
  /** Date the underlying pricing data was last refreshed (YYYY-MM-DD). */
  pricesAsOf?: string;
  /**
   * The oldest meter date behind the services in this estimate (YYYY-MM-DD),
   * when the prices have been unchanged for long enough to be worth saying.
   *
   * Not a staleness warning. Azure's retail API returns current prices, so a
   * 2018 meter is genuinely today's price — it is simply one Azure has not
   * repriced since. `pricesAsOf` says when the data was fetched, which answers
   * "is this current?"; this answers the different question a customer asks in
   * the room, "how firm is this number?" The shipped corpus spans 2018-02 to
   * 2026-07, and a price that has held for eight years is a materially
   * different planning input from one that moved last month.
   */
  oldestMeterAsOf?: string;
  /** Billing term the costs reflect (e.g. "Pay-as-you-go", "Savings Plan (1-year)"). */
  pricingTerm?: string;
}

export type PricingCurrency = 'USD' | 'JPY' | 'EUR' | 'GBP';
export type PricingScenarioKind = 'development' | 'production' | 'custom';

export interface PricingScenario {
  id: string;
  name: string;
  kind: PricingScenarioKind;
  pricingMode: 'payg' | 'reserved1yr';
  /** Multiplier applied to fixed/capacity-based services. */
  capacityMultiplier: number;
  /** Multiplier applied to consumption/usage-based services. */
  usageMultiplier: number;
  /** Additional negotiated discount applied after the Azure pricing term. */
  discountPercent: number;
  /** Optional support/operations allowance added after discounts. */
  supportPercent: number;
  currency: PricingCurrency;
  /** Planning conversion rate expressed as units of currency per USD. */
  exchangeRate: number;
}

export interface PricingScenarioResult {
  scenario: PricingScenario;
  baseMonthlyUsd: number;
  discountedMonthlyUsd: number;
  supportMonthlyUsd: number;
  totalMonthlyUsd: number;
  totalAnnualUsd: number;
  totalMonthly: number;
  totalAnnual: number;
  serviceCount: number;
  usageBasedCount: number;
  pricesAsOf: string;
}

/**
 * Cached pricing data
 */
export interface CachedPricing {
  data: ServicePricing;
  timestamp: number;
  expiresAt: number;
}

/**
 * API query parameters
 */
export interface PricingQueryParams {
  serviceName?: string;
  armRegionName?: string;
  currencyCode?: string;
  filter?: string;
}
