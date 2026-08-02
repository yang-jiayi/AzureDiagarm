// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

import type { Node } from 'reactflow';
import type {
  NodePricingConfig,
  PricingCurrency,
  PricingScenario,
  PricingScenarioResult,
} from '../types/pricing';
import { PRICING_DATA_AS_OF } from '../data/azurePricing';
import { calculateCostBreakdown } from './costEstimationService';
import { readLocalStorage, writeLocalStorage } from '../utils/safeStorage';
import { csvTextCell } from '../utils/csv';

const STORAGE_KEY = 'azurediagarm.pricing-scenarios.v1';
const MAX_SCENARIOS = 12;

const DEFAULT_EXCHANGE_RATES: Record<PricingCurrency, number> = {
  USD: 1,
  JPY: 150,
  EUR: 0.92,
  GBP: 0.79,
};

export const DEFAULT_PRICING_SCENARIOS: PricingScenario[] = [
  {
    id: 'development',
    name: 'Development / Test',
    kind: 'development',
    pricingMode: 'payg',
    capacityMultiplier: 0.35,
    usageMultiplier: 0.25,
    discountPercent: 0,
    supportPercent: 0,
    currency: 'USD',
    exchangeRate: 1,
  },
  {
    id: 'production',
    name: 'Production',
    kind: 'production',
    pricingMode: 'reserved1yr',
    capacityMultiplier: 1,
    usageMultiplier: 1,
    discountPercent: 0,
    supportPercent: 0,
    currency: 'USD',
    exchangeRate: 1,
  },
];

function boundedNumber(value: unknown, fallback: number, min: number, max: number): number {
  const parsed = Number(value);
  if (!Number.isFinite(parsed)) return fallback;
  return Math.min(max, Math.max(min, parsed));
}

function normalizeScenario(value: unknown, index: number): PricingScenario | null {
  if (!value || typeof value !== 'object') return null;
  const raw = value as Partial<PricingScenario>;
  const currency: PricingCurrency = raw.currency === 'JPY'
    || raw.currency === 'EUR'
    || raw.currency === 'GBP'
    ? raw.currency
    : 'USD';
  const kind = raw.kind === 'development' || raw.kind === 'production'
    ? raw.kind
    : 'custom';
  const name = typeof raw.name === 'string' ? raw.name.trim().slice(0, 80) : '';
  if (!name) return null;
  return {
    id: typeof raw.id === 'string' && raw.id.trim()
      ? raw.id.trim().slice(0, 100)
      : `scenario-${index + 1}`,
    name,
    kind,
    pricingMode: raw.pricingMode === 'reserved1yr' ? 'reserved1yr' : 'payg',
    capacityMultiplier: boundedNumber(raw.capacityMultiplier, 1, 0, 20),
    usageMultiplier: boundedNumber(raw.usageMultiplier, 1, 0, 20),
    discountPercent: boundedNumber(raw.discountPercent, 0, 0, 90),
    supportPercent: boundedNumber(raw.supportPercent, 0, 0, 50),
    currency,
    exchangeRate: boundedNumber(
      raw.exchangeRate,
      DEFAULT_EXCHANGE_RATES[currency],
      0.0001,
      100_000,
    ),
  };
}

export function normalizePricingScenarios(value: unknown): PricingScenario[] {
  if (!Array.isArray(value)) {
    return DEFAULT_PRICING_SCENARIOS.map((scenario) => ({ ...scenario }));
  }
  const scenarios = value
    .slice(0, MAX_SCENARIOS)
    .map(normalizeScenario)
    .filter((scenario): scenario is PricingScenario => Boolean(scenario));
  return scenarios.length > 0
    ? scenarios
    : DEFAULT_PRICING_SCENARIOS.map((scenario) => ({ ...scenario }));
}

export function loadPricingScenarios(): PricingScenario[] {
  const raw = readLocalStorage(STORAGE_KEY);
  if (!raw) return DEFAULT_PRICING_SCENARIOS.map((scenario) => ({ ...scenario }));
  try {
    return normalizePricingScenarios(JSON.parse(raw));
  } catch {
    return DEFAULT_PRICING_SCENARIOS.map((scenario) => ({ ...scenario }));
  }
}

export function savePricingScenarios(scenarios: PricingScenario[]): void {
  const normalized = scenarios
    .slice(0, MAX_SCENARIOS)
    .map(normalizeScenario)
    .filter((scenario): scenario is PricingScenario => Boolean(scenario));
  writeLocalStorage(STORAGE_KEY, JSON.stringify(normalized));
}

export function createPricingScenario(
  source?: PricingScenario,
): PricingScenario {
  const id = typeof crypto !== 'undefined' && typeof crypto.randomUUID === 'function'
    ? crypto.randomUUID()
    : `scenario-${Date.now()}-${Math.random().toString(36).slice(2, 9)}`;
  return {
    ...(source || {
      name: 'Custom scenario',
      kind: 'custom' as const,
      pricingMode: 'payg' as const,
      capacityMultiplier: 1,
      usageMultiplier: 1,
      discountPercent: 0,
      supportPercent: 0,
      currency: 'USD' as const,
      exchangeRate: 1,
    }),
    id,
    kind: 'custom',
    name: source ? `${source.name} copy`.slice(0, 80) : 'Custom scenario',
  };
}

export function defaultExchangeRate(currency: PricingCurrency): number {
  return DEFAULT_EXCHANGE_RATES[currency];
}

export function calculatePricingScenario(
  nodes: Node[],
  scenario: PricingScenario,
): PricingScenarioResult {
  const breakdown = calculateCostBreakdown(nodes, undefined, scenario.pricingMode);
  const nodesById = new Map(nodes.map((node) => [node.id, node]));
  let scaledMonthlyUsd = 0;
  let usageBasedCount = 0;

  for (const service of breakdown.byService) {
    const pricing = nodesById.get(service.nodeId)?.data?.pricing as NodePricingConfig | undefined;
    const usageBased = pricing?.isUsageBased === true;
    if (usageBased) usageBasedCount += 1;
    scaledMonthlyUsd += service.cost * (
      usageBased ? scenario.usageMultiplier : scenario.capacityMultiplier
    );
  }

  const discountedMonthlyUsd = scaledMonthlyUsd
    * (1 - scenario.discountPercent / 100);
  const supportMonthlyUsd = discountedMonthlyUsd * (scenario.supportPercent / 100);
  const totalMonthlyUsd = discountedMonthlyUsd + supportMonthlyUsd;
  const rate = scenario.currency === 'USD' ? 1 : scenario.exchangeRate;

  return {
    scenario,
    baseMonthlyUsd: scaledMonthlyUsd,
    discountedMonthlyUsd,
    supportMonthlyUsd,
    totalMonthlyUsd,
    totalAnnualUsd: totalMonthlyUsd * 12,
    totalMonthly: totalMonthlyUsd * rate,
    totalAnnual: totalMonthlyUsd * rate * 12,
    serviceCount: breakdown.byService.length,
    usageBasedCount,
    pricesAsOf: breakdown.pricesAsOf || PRICING_DATA_AS_OF,
  };
}

export function formatScenarioCurrency(
  value: number,
  currency: PricingCurrency,
  locale: string,
): string {
  return new Intl.NumberFormat(locale, {
    style: 'currency',
    currency,
    maximumFractionDigits: currency === 'JPY' ? 0 : 2,
  }).format(value);
}

export function exportPricingScenariosCsv(results: PricingScenarioResult[]): string {
  const lines = [
    'Scenario,Kind,Pricing Term,Capacity Multiplier,Usage Multiplier,Discount %,Support %,Currency,Planning FX per USD,Monthly,Annual,Monthly USD,Annual USD,Prices As Of',
  ];
  for (const result of results) {
    const scenario = result.scenario;
    lines.push([
      csvTextCell(scenario.name, true),
      csvTextCell(scenario.kind),
      csvTextCell(scenario.pricingMode),
      scenario.capacityMultiplier.toFixed(2),
      scenario.usageMultiplier.toFixed(2),
      scenario.discountPercent.toFixed(2),
      scenario.supportPercent.toFixed(2),
      scenario.currency,
      scenario.exchangeRate.toFixed(4),
      result.totalMonthly.toFixed(2),
      result.totalAnnual.toFixed(2),
      result.totalMonthlyUsd.toFixed(2),
      result.totalAnnualUsd.toFixed(2),
      csvTextCell(result.pricesAsOf),
    ].join(','));
  }
  return lines.join('\n');
}
