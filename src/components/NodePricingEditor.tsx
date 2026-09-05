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
import { MAX_PRICING_AMOUNT, MAX_PRICING_QUANTITY } from '../services/pricingConfiguration';
import { useLanguage } from '../i18n/LanguageContext';
import { localize } from '../i18n/localization';
import ModalScaffold from './ModalScaffold';
import './NodePricingEditor.css';

const CURRENT_ESTIMATE_TIER = '__azurediagarm_current_estimate__';

function pricesMatch(left: number | null, right: number | null): boolean {
  if (left === null || right === null) return left === right;
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
    pricing.isCustom || pricesMatch(tier.monthlyPrice, pricing.estimatedCost)
  );
}

interface NodePricingEditorProps {
  /** Service name, used as the pricing lookup key (node.data.label). */
  serviceType: string;
  pricing: NodePricingConfig;
  onApply: (updated: NodePricingConfig) => void;
  onClose: () => void;
  returnFocusTarget?: HTMLElement | null;
}

export default function NodePricingEditor({
  serviceType,
  pricing,
  onApply,
  onClose,
  returnFocusTarget = null,
}: NodePricingEditorProps) {
  const { language } = useLanguage();
  const [tiers, setTiers] = useState<PricingTier[]>([]);
  const [loadingTiers, setLoadingTiers] = useState(true);
  const [tier, setTier] = useState<string>(CURRENT_ESTIMATE_TIER);
  const [quantityText, setQuantityText] = useState(() => String(pricing.quantity));
  const [useCustom, setUseCustom] = useState<boolean>(!!pricing.isCustom);
  const [customPrice, setCustomPrice] = useState<string>(
    pricing.customPrice != null ? String(pricing.customPrice) : String(pricing.estimatedCost ?? ''),
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
      .catch(error => {
        console.error(`Failed to load pricing options for ${serviceType}:`, error);
        if (!cancelled) {
          setLoadingTiers(false);
          setSaveError(localize(language, {
            en: 'Pricing options could not be loaded. Your existing estimate has not changed.',
            ja: '価格の選択肢を読み込めませんでした。既存の見積もりは変更されていません。',
          }));
        }
      });
    return () => {
      cancelled = true;
    };
  }, [pricing, serviceType, language]);

  // Preview mirrors calculateMonthlyCost: unit price x quantity.
  //
  // Usage-based services are the exception. Their catalog "tiers" are
  // consumption meters without a monthly total until usage is known. Keep
  // the current consumption estimate rather than replacing it with a raw meter.
  const tiersSelectable = !pricing.isUsageBased && tiers.length > 0;
  const matchingCurrentTier = findCurrentTier(tiers, pricing);
  const showCurrentEstimateOption = tiersSelectable && !matchingCurrentTier;
  const selectedTier = tiers.find(candidate =>
    candidate.id === tier || candidate.skuName === tier || candidate.name === tier
  );
  const parsedCustom = Number.parseFloat(customPrice);
  const customIsValid = Number.isFinite(parsedCustom) && parsedCustom >= 0
    && parsedCustom <= MAX_PRICING_AMOUNT;
  const quantity = Number(quantityText);
  const quantityIsValid = quantityText.trim() !== ''
    && Number.isInteger(quantity) && quantity >= 1 && quantity <= MAX_PRICING_QUANTITY;
  const unitCost = useCustom
    ? (customIsValid ? parsedCustom : null)
    : tiersSelectable
      ? (selectedTier ? selectedTier.monthlyPrice : pricing.estimatedCost)
      : pricing.estimatedCost;
  const previewTotal = unitCost === null || !quantityIsValid ? null : unitCost * quantity;
  const formatCost = (amount: number | null) => formatMonthlyCost(
    amount, localize(language, { en: 'Unpriced', ja: '価格未設定' }),
  );

  const canApply = !saving
    && quantityIsValid
    && (previewTotal === null || Number.isFinite(previewTotal))
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
    <ModalScaffold
      isOpen
      onClose={handleClose}
      className="npe-modal"
      overlayClassName="npe-modal-overlay"
      ariaLabelledBy="node-pricing-editor-title"
      closeOnBackdrop={!saving}
      closeOnEscape={!saving}
      returnFocusTarget={returnFocusTarget}
    >
        <div className="modal-header npe-modal-header">
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
            className="modal-close npe-modal-close"
            onClick={handleClose}
            disabled={saving}
            aria-label={localize(language, { en: 'Close', ja: '閉じる' })}
          >
            <X size={18} />
          </button>
        </div>

        <div className="modal-body npe-modal-body">
          <p className="npe-note">
            {localize(language, {
              en: `Estimates are indicative catalog prices for ${pricing.region}. Adjust the SKU and instance count to match your design, or override the figure with your own negotiated price.`,
              ja: `${pricing.region} の参考カタログ価格です。設計に合わせてSKUと数量を調整するか、契約価格で上書きしてください。`,
            })}
          </p>

          <label className="npe-field azd-field">
            <span className="npe-label">
              {localize(language, { en: 'Tier / SKU', ja: 'Tier / SKU' })}
            </span>
            {loadingTiers ? (
              <span className="npe-hint azd-field-hint">
                {localize(language, { en: 'Loading available tiers…', ja: '利用可能なTierを読み込み中…' })}
              </span>
            ) : tiersSelectable ? (
              <select
                className="npe-input azd-control"
                value={tier}
                disabled={saving || useCustom}
                onChange={e => setTier(e.target.value)}
              >
                {showCurrentEstimateOption && (
                  <option value={CURRENT_ESTIMATE_TIER}>
                    {localize(language, {
                      en: 'Current estimate',
                      ja: '現在の見積もり',
                    })} — {formatCost(pricing.estimatedCost)}
                  </option>
                )}
                {tiers.map(t => (
                  <option key={tierValue(t)} value={tierValue(t)}>
                    {t.name} — {formatCost(t.monthlyPrice)} {t.unit ? `(${t.unit})` : ''}
                  </option>
                ))}
              </select>
            ) : pricing.isUsageBased ? (
              <span className="npe-hint azd-field-hint">
                {localize(language, {
                  en: `${serviceType} bills on consumption, so there is no monthly SKU to pick. The displayed figure estimates typical usage; override it below if you have a better value.`,
                  ja: `${serviceType} は従量課金のため、月額SKUはありません。表示額は一般的な使用量の参考値です。より正確な値がある場合は下で上書きしてください。`,
                })}
              </span>
            ) : (
              <span className="npe-hint azd-field-hint">
                {localize(language, {
                  en: 'No catalog tiers are available for this service. Use a custom price below.',
                  ja: 'このサービスで利用可能なカタログTierはありません。下で独自価格を指定してください。',
                })}
              </span>
            )}
          </label>

          <label className="npe-field azd-field">
            <span className="npe-label">
              {localize(language, { en: 'Instances / units', ja: 'インスタンス / ユニット数' })}
            </span>
            <input
              className="npe-input azd-control"
              type="number"
              min={1}
              max={MAX_PRICING_QUANTITY}
              step={1}
              value={quantityText}
              disabled={saving}
              aria-invalid={!quantityIsValid}
              onChange={e => setQuantityText(e.target.value)}
            />
            {!quantityIsValid && (
              <span className="npe-error azd-field-error">
                {localize(language, {
                  en: `Enter an integer between 1 and ${MAX_PRICING_QUANTITY}.`,
                  ja: `1から${MAX_PRICING_QUANTITY}の整数を入力してください。`,
                })}
              </span>
            )}
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
            <label className="npe-field azd-field">
              <span className="npe-label">
                {localize(language, {
                  en: 'Custom price (USD / month / unit)',
                  ja: '独自価格（USD / 月 / ユニット）',
                })}
              </span>
              <input
                className={`npe-input azd-control${customIsValid ? '' : ' npe-input--invalid'}`}
                type="number"
                min={0}
                max={MAX_PRICING_AMOUNT}
                step="0.01"
                value={customPrice}
                disabled={saving}
                onChange={e => setCustomPrice(e.target.value)}
              />
              {!customIsValid && (
                <span className="npe-error azd-field-error">
                  {localize(language, {
                    en: `Enter a finite number between 0 and ${MAX_PRICING_AMOUNT}.`,
                    ja: `0から${MAX_PRICING_AMOUNT}の有限の数値を入力してください。`,
                  })}
                </span>
              )}
            </label>
          )}

          <div className="npe-preview azd-callout azd-callout--success">
            <span className="npe-preview-label">
              {localize(language, { en: 'Estimated monthly cost', ja: '月額参考見積もり' })}
            </span>
            <span className="npe-preview-value">{formatCost(previewTotal)}</span>
            {quantityIsValid && quantity > 1 && (
              <span className="npe-preview-detail">
                {formatCost(unitCost)} × {quantity}
              </span>
            )}
          </div>
          {saveError && (
            <div className="npe-error azd-callout azd-callout--danger" role="alert">
              {saveError}
            </div>
          )}
        </div>

        <div className="modal-actions npe-modal-footer">
          <button
            type="button"
            className="azd-button azd-button--secondary"
            onClick={handleClose}
            disabled={saving}
          >
            {localize(language, { en: 'Cancel', ja: 'キャンセル' })}
          </button>
          <button
            type="button"
            className="azd-button azd-button--primary"
            onClick={handleApply}
            disabled={!canApply}
          >
            {saving
              ? localize(language, { en: 'Applying…', ja: '適用中…' })
              : localize(language, { en: 'Apply', ja: '適用' })}
          </button>
        </div>
    </ModalScaffold>
  );
}
