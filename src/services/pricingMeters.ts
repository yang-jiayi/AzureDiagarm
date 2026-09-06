import type { AzureRetailPrice, PricingProvenance, PricingTier } from '../types/pricing';

export function meterProvenance(item: AzureRetailPrice): PricingProvenance {
  return {
    kind: 'official-meter', source: 'azure-retail-prices',
    asOf: item.effectiveStartDate || undefined,
    meterId: item.meterId || undefined,
    meterName: item.meterName,
    unit: item.unitOfMeasure,
    assumptions: [],
    note: 'Selected meter only; other meters, taxes and discounts are excluded.',
  };
}

/** Only time-based instance meters have an implicit monthly conversion. */
export function parsePricingTiers(items: AzureRetailPrice[], serviceName = ''): PricingTier[] {
  const tierMap = new Map<string, PricingTier>();
  const isAppService = serviceName.toLowerCase() === 'azure app service';
  const meterKey = (item: AzureRetailPrice) => JSON.stringify([
    item.productName || '', item.skuName || item.armSkuName, item.meterName, item.unitOfMeasure,
  ]);
  const tieredMeters = new Set(items.filter(item => item.tierMinimumUnits > 0).map(meterKey));
  const rates = new Map<string, number>();
  for (const item of items) {
    const key = meterKey(item);
    const rate = item.retailPrice ?? item.unitPrice;
    // Compacted files omit band thresholds but retain the different band rates.
    if (rates.has(key) && rates.get(key) !== rate) tieredMeters.add(key);
    rates.set(key, rate);
  }
  for (const item of items) {
    const skuName = item.skuName || item.armSkuName;
    const rate = item.retailPrice ?? item.unitPrice;
    if (!skuName || !Number.isFinite(rate) || rate < 0) continue;
    const unit = item.unitOfMeasure || '';
    const multiplier = /^1\s*hour$/i.test(unit) ? 730 :
      /^1\/?month$/i.test(unit.replace(/\s/g, '')) ? 1 :
      /^1\/?day$/i.test(unit.replace(/\s/g, '')) ? 30 :
      /^1\/?year$/i.test(unit.replace(/\s/g, '')) ? 1 / 12 : null;
    const provenance = meterProvenance(item);
    const tiered = tieredMeters.has(meterKey(item));
    const usage = multiplier === null || tiered;
    if (usage) {
      provenance.kind = 'unpriced';
      provenance.note = tiered
        ? 'Tiered meter needs a complete pricing calculation; supply a custom estimate.'
        : 'Usage is unspecified; the selected meter is not a complete service estimate.';
    } else {
      provenance.assumptions.push({
        label: 'Monthly operation', value: multiplier!,
        unit: multiplier === 730 ? 'hours/month' : multiplier === 30 ? 'days/month' :
          multiplier === 1 / 12 ? 'years/month' : 'months',
      });
    }
    const oneYear = Array.isArray(item.savingsPlan)
      ? item.savingsPlan.find(plan => /^1\s*year$/i.test(plan.term)) : undefined;
    const reservedRate = oneYear?.retailPrice ?? oneYear?.unitPrice;
    const tierId = isAppService ? `${item.productName}::${skuName}` : skuName;
    const tier: PricingTier = {
      id: tierId,
      name: isAppService ? `${skuName} (${/- Linux$/i.test(item.productName) ? 'Linux' : 'Windows'})` : skuName,
      skuName, monthlyPrice: usage ? null : rate * multiplier!,
      ...(multiplier === 730 && !usage ? { hourlyPrice: rate } : {}),
      unit, description: item.meterName, provenance,
      isUsageBased: usage,
      ...(usage && !!unit.trim() && !tiered
        ? { usage: { amount: null, unit, unitPrice: rate } } : {}),
      ...(!usage && multiplier !== null && reservedRate !== undefined && Number.isFinite(reservedRate) && reservedRate >= 0
        ? { reserved1yrMonthly: reservedRate * multiplier } : {}),
    };
    const existing = tierMap.get(tierId);
    if (!existing || (tier.monthlyPrice !== null &&
        (existing.monthlyPrice === null || tier.monthlyPrice < existing.monthlyPrice))) {
      tierMap.set(tierId, tier);
    }
  }
  return [...tierMap.values()].sort((a, b) =>
    (a.monthlyPrice ?? Infinity) - (b.monthlyPrice ?? Infinity));
}
