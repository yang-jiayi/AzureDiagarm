// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

/**
 * Per-node cost editor
 *
 * Lets a user change the tier/SKU, the instance count, or override the price
 * outright for a single service. Before this, every estimate used the catalog
 * default tier at quantity 1 with no way to change either, which users pushed
 * back on ("cost as a fixed value is not acceptable, i would rather hide it or
 * make it configurable").
 *
 * `estimatedCost` is stored PER UNIT — calculateCostBreakdown and AzureNode
 * both multiply by quantity themselves.
 */

import { useEffect, useRef, useState } from 'react';
import { DollarSign, X } from 'lucide-react';
import type { NodePricingConfig, PricingTier } from '../types/pricing';
import {
  getAvailableTiers,
  initializeNodePricing,
  updateNodePricing,
  setCustomPricing,
} from '../services/costEstimationService';
import { formatMonthlyCost } from '../utils/pricingHelpers';
import { useLanguage } from '../i18n/LanguageContext';
import { localize } from '../i18n/localization';
import './NodePricingEditor.css';

const MAX_QUANTITY = 100_000;
const CURRENT_ESTIMATE_TIER = '__azurediagarm_current_estimate__';

function normalizeQuantity(value: number): number {
  if (!Number.isFinite(value)) return 1;
  return Math.min(MAX_QUANTITY, Math.max(1, Math.trunc(value)));
}

function pricesMatch(left: number, right: number): boolean {
  return Math.abs(left - right) <= Math.max(0.01, Math.abs(right) * 0.001);
}

function tierValue(tier: PricingTier): string {
  return tier.id || tier.skuName || tier.name;
}

function tierIdentityMatches(tier: PricingTier, pricing: NodePricingConfig): boolean {
  if (pricing.tierId) return tier.id === pricing.tierId;
  return tier.skuName === pricing.skuName
    || tier.skuName === pricing.tier
    || tier.name === pricing.tier;
}

function findCurrentTier(
  tiers: PricingTier[],
  pricing: NodePricingConfig,
): PricingTier | undefined {
  const identityMatches = tiers.filter(tier => tierIdentityMatches(tier, pricing));
  if (pricing.tierId && pricing.isCustom) return identityMatches[0];
  if (!pricing.tierId && identityMatches.length !== 1) return undefined;
  return identityMatches.find(tier =>
    pricing.isCustom || pricesMatch(tier.monthlyPrice, pricing.estimatedCost ?? 0)
  );
}

interface NodePricingEditorProps {
  /** Service name, used as the pricing lookup key (node.data.label). */
  serviceType: string;
  pricing: NodePricingConfig;
  onApply: (updated: NodePricingConfig) => void;
  onClose: () => void;
}

export default function NodePricingEditor({
  serviceType,
  pricing,
  onApply,
  onClose,
}: NodePricingEditorProps) {
  const { language } = useLanguage();
  const [tiers, setTiers] = useState<PricingTier[]>([]);
  const [loadingTiers, setLoadingTiers] = useState(true);
  const [tier, setTier] = useState<string>(CURRENT_ESTIMATE_TIER);
  const [quantity, setQuantity] = useState<number>(() => normalizeQuantity(pricing.quantity));
  const [useCustom, setUseCustom] = useState<boolean>(!!pricing.isCustom);
  const [customPrice, setCustomPrice] = useState<string>(
    pricing.customPrice != null ? String(pricing.customPrice) : String(pricing.estimatedCost ?? 0),
  );
  const [saving, setSaving] = useState(false);
  const [saveError, setSaveError] = useState('');
  const activeRef = useRef(true);

  useEffect(() => {
    activeRef.current = true;
    return () => {
      activeRef.current = false;
    };
  }, []);

  useEffect(() => {
    let cancelled = false;
    getAvailableTiers(serviceType, pricing.region)
      .then(result => {
        if (!cancelled) {
          setTiers(result);
          const matchingTier = findCurrentTier(result, pricing);
          setTier(matchingTier
            ? tierValue(matchingTier)
            : CURRENT_ESTIMATE_TIER);
          setLoadingTiers(false);
        }
      })
      .catch(() => {
        if (!cancelled) setLoadingTiers(false);
      });
    return () => {
      cancelled = true;
    };
  }, [pricing, serviceType]);

  useEffect(() => {
    const handleKeyDown = (event: KeyboardEvent) => {
      if (event.key === 'Escape' && !saving) onClose();
    };
    document.addEventListener('keydown', handleKeyDown);
    return () => document.removeEventListener('keydown', handleKeyDown);
  }, [onClose, saving]);

  // Preview mirrors calculateMonthlyCost: unit price x quantity.
  //
  // Usage-based services are the exception. Their catalog "tiers" are
  // consumption meters (per 1K tokens, per hour) that carry a $0 monthly
  // price, while the badge shows a separate consumption fallback estimate.
  // Presenting those as selectable SKUs would let a user apply one and
  // silently zero a real estimate, so the tier picker is suppressed and the
  // existing estimate is kept.
  const tiersSelectable = !pricing.isUsageBased && tiers.length > 0;
  const matchingCurrentTier = findCurrentTier(tiers, pricing);
  const showCurrentEstimateOption = tiersSelectable && !matchingCurrentTier;
  const selectedTier = tiers.find(candidate =>
    candidate.id === tier || candidate.skuName === tier || candidate.name === tier
  );
  const parsedCustom = Number.parseFloat(customPrice);
  const customIsValid = Number.isFinite(parsedCustom) && parsedCustom >= 0;
  const unitCost = useCustom
    ? (customIsValid ? parsedCustom : 0)
    : tiersSelectable
      ? (selectedTier?.monthlyPrice ?? pricing.estimatedCost ?? 0)
      : (pricing.estimatedCost ?? 0);
  const previewTotal = unitCost * (quantity > 0 ? quantity : 1);

  const canApply = !saving
    && Number.isInteger(quantity)
    && quantity >= 1
    && quantity <= MAX_QUANTITY
    && Number.isFinite(previewTotal)
    && (!useCustom || customIsValid);

  const handleClose = () => {
    if (!saving) onClose();
  };

  const handleApply = async () => {
    if (!canApply) return;
    setSaveError('');
    setSaving(true);
    try {
      let updatedPricing: NodePricingConfig;
      if (useCustom) {
        // Custom price is per unit, so quantity still scales it.
        updatedPricing = { ...setCustomPricing(pricing, parsedCustom), quantity };
      } else if (pricing.isCustom && !tiersSelectable) {
        const automaticPricing = await initializeNodePricing(serviceType, pricing.region);
        if (!automaticPricing) {
          throw new Error(`Automatic pricing is not available for ${serviceType}`);
        }
        updatedPricing = { ...automaticPricing, quantity };
      } else if (!tiersSelectable) {
        // Keep the consumption estimate as-is; only the unit count changed.
        updatedPricing = { ...pricing, quantity, lastUpdated: new Date().toISOString() };
      } else if (tier === CURRENT_ESTIMATE_TIER) {
        // Imported and older diagrams may carry a generic tier label that does
        // not uniquely identify a catalog SKU. Preserve the displayed unit
        // price until the user deliberately selects a concrete catalog tier.
        updatedPricing = { ...pricing, quantity, lastUpdated: new Date().toISOString() };
      } else {
        updatedPricing = await updateNodePricing(serviceType, pricing, tier, quantity, pricing.region);
      }
      if (!activeRef.current) return;
      onApply(updatedPricing);
      onClose();
    } catch (error) {
      console.error(`Failed to update pricing for ${serviceType}:`, error);
      if (activeRef.current) {
        setSaveError(localize(language, {
          en: 'The pricing update failed. Check the selected tier and try again.',
          ja: '価格の更新に失敗しました。選択したTierを確認して、もう一度お試しください。',
        }));
      }
    } finally {
      if (activeRef.current) setSaving(false);
    }
  };

  return (
    <div className="npe-modal-overlay" onClick={handleClose}>
      <div
        className="npe-modal"
        onClick={e => e.stopPropagation()}
        role="dialog"
        aria-modal="true"
        aria-labelledby="node-pricing-editor-title"
      >
        <div className="npe-modal-header">
          <div className="npe-modal-title" id="node-pricing-editor-title">
            <DollarSign size={20} />
            <span>
              {localize(language, {
                en: `Cost settings — ${serviceType}`,
                ja: `コスト設定 — ${serviceType}`,
              })}
            </span>
          </div>
          <button
            type="button"
            className="npe-modal-close"
            onClick={handleClose}
            disabled={saving}
            autoFocus
            aria-label={localize(language, { en: 'Close', ja: '閉じる' })}
          >
            <X size={18} />
          </button>
        </div>

        <div className="npe-modal-body">
          <p className="npe-note">
            {localize(language, {
              en: `Estimates are indicative catalog prices for ${pricing.region}. Adjust the SKU and instance count to match your design, or override the figure with your own negotiated price.`,
              ja: `${pricing.region} の参考カタログ価格です。設計に合わせてSKUと数量を調整するか、契約価格で上書きしてください。`,
            })}
          </p>

          <label className="npe-field">
            <span className="npe-label">
              {localize(language, { en: 'Tier / SKU', ja: 'Tier / SKU' })}
            </span>
            {loadingTiers ? (
              <span className="npe-hint">
                {localize(language, { en: 'Loading available tiers…', ja: '利用可能なTierを読み込み中…' })}
              </span>
            ) : tiersSelectable ? (
              <select
                className="npe-input"
                value={tier}
                disabled={saving || useCustom}
                onChange={e => setTier(e.target.value)}
              >
                {showCurrentEstimateOption && (
                  <option value={CURRENT_ESTIMATE_TIER}>
                    {localize(language, {
                      en: 'Current estimate',
                      ja: '現在の見積もり',
                    })} — {formatMonthlyCost(pricing.estimatedCost ?? 0)}
                  </option>
                )}
                {tiers.map(t => (
                  <option key={tierValue(t)} value={tierValue(t)}>
                    {t.name} — {formatMonthlyCost(t.monthlyPrice)} {t.unit ? `(${t.unit})` : ''}
                  </option>
                ))}
              </select>
            ) : pricing.isUsageBased ? (
              <span className="npe-hint">
                {localize(language, {
                  en: `${serviceType} bills on consumption, so there is no monthly SKU to pick. The displayed figure estimates typical usage; override it below if you have a better value.`,
                  ja: `${serviceType} は従量課金のため、月額SKUはありません。表示額は一般的な使用量の参考値です。より正確な値がある場合は下で上書きしてください。`,
                })}
              </span>
            ) : (
              <span className="npe-hint">
                {localize(language, {
                  en: 'No catalog tiers are available for this service. Use a custom price below.',
                  ja: 'このサービスで利用可能なカタログTierはありません。下で独自価格を指定してください。',
                })}
              </span>
            )}
          </label>

          <label className="npe-field">
            <span className="npe-label">
              {localize(language, { en: 'Instances / units', ja: 'インスタンス / ユニット数' })}
            </span>
            <input
              className="npe-input"
              type="number"
              min={1}
              max={MAX_QUANTITY}
              step={1}
              value={quantity}
              disabled={saving}
              onChange={e => setQuantity(normalizeQuantity(Number(e.target.value)))}
            />
          </label>

          <label className="npe-checkbox">
            <input
              type="checkbox"
              checked={useCustom}
              disabled={saving}
              onChange={e => {
                const nextUseCustom = e.target.checked;
                setUseCustom(nextUseCustom);
                setSaveError('');
              }}
            />
            <span>
              {localize(language, {
                en: 'Override with a custom monthly price (per unit)',
                ja: '独自の月額価格（1ユニットあたり）で上書き',
              })}
            </span>
          </label>

          {useCustom && (
            <label className="npe-field">
              <span className="npe-label">
                {localize(language, {
                  en: 'Custom price (USD / month / unit)',
                  ja: '独自価格（USD / 月 / ユニット）',
                })}
              </span>
              <input
                className={`npe-input${customIsValid ? '' : ' npe-input--invalid'}`}
                type="number"
                min={0}
                step="0.01"
                value={customPrice}
                disabled={saving}
                onChange={e => setCustomPrice(e.target.value)}
              />
              {!customIsValid && (
                <span className="npe-error">
                  {localize(language, {
                    en: 'Enter a finite number of 0 or more.',
                    ja: '0以上の有限の数値を入力してください。',
                  })}
                </span>
              )}
            </label>
          )}

          <div className="npe-preview">
            <span className="npe-preview-label">
              {localize(language, { en: 'Estimated monthly cost', ja: '月額参考見積もり' })}
            </span>
            <span className="npe-preview-value">{formatMonthlyCost(previewTotal)}</span>
            {quantity > 1 && (
              <span className="npe-preview-detail">
                {formatMonthlyCost(unitCost)} × {quantity}
              </span>
            )}
          </div>
          {saveError && (
            <div className="npe-error" role="alert">
              {saveError}
            </div>
          )}
        </div>

        <div className="npe-modal-footer">
          <button type="button" className="npe-btn npe-btn--ghost" onClick={handleClose} disabled={saving}>
            {localize(language, { en: 'Cancel', ja: 'キャンセル' })}
          </button>
          <button type="button" className="npe-btn npe-btn--primary" onClick={handleApply} disabled={!canApply}>
            {saving
              ? localize(language, { en: 'Applying…', ja: '適用中…' })
              : localize(language, { en: 'Apply', ja: '適用' })}
          </button>
        </div>
      </div>
    </div>
  );
}
