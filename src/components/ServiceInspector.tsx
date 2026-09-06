import { useEffect, useMemo, useRef, useState } from 'react';
import type { Node } from 'reactflow';
import { X } from 'lucide-react';
import type { NodePricingConfig, PricingTier } from '../types/pricing';
import { getNodePricingTiers, setCustomPricing } from '../services/costEstimationService';
import {
  buildServiceInspectorData, getPricingProvenance, MAX_PRICING_AMOUNT, MAX_PRICING_QUANTITY,
  isPricingUsage, pricingFromTier, selectPricingTier, validatePricingAmount, validatePricingQuantity, withUsageAmount,
} from '../services/pricingConfiguration';
import { getActiveRegion } from '../services/regionalPricingService';
import { getDefaultTier, getFallbackDefaultSku } from '../data/azurePricing';
import { formatCurrency } from '../utils/pricingHelpers';
import { useModalFocus } from '../hooks/useModalFocus';
import { useLanguage } from '../i18n/LanguageContext';
import { localize } from '../i18n/localization';
import './ServiceInspector.css';

export interface ServiceInspectorProps {
  node: Node;
  onUpdateNode: (nodeId: string, data: Record<string, unknown>) => void;
  onClose: () => void;
}

function ServiceInspectorEditor({ node, onUpdateNode, onClose }: ServiceInspectorProps) {
  const { language } = useLanguage();
  const text = (en: string, ja: string) => localize(language, { en, ja });
  const scopeRef = useModalFocus(true, onClose);
  const initialData = useRef(node.data);
  const current = initialData.current.pricing as NodePricingConfig | undefined;
  const serviceType = String(initialData.current.serviceName || initialData.current.serviceType || initialData.current.label || '');
  const region = current?.region || getActiveRegion();
  const [label, setLabel] = useState(String(node.data.label || ''));
  const [description, setDescription] = useState(String(node.data.description || ''));
  const [tiers, setTiers] = useState<PricingTier[]>([]);
  const [loading, setLoading] = useState(true);
  const [sku, setSku] = useState(current?.tierId || current?.skuName || current?.tier || '');
  const [tierChanged, setTierChanged] = useState(false);
  const [quantity, setQuantity] = useState(String(current?.quantity ?? 1));
  const [custom, setCustom] = useState(!!current?.isCustom);
  const [customPrice, setCustomPrice] = useState(String(current?.customPrice ?? current?.estimatedCost ?? ''));
  const [usageAmount, setUsageAmount] = useState<string | undefined>(undefined);
  const [applyError, setApplyError] = useState('');

  useEffect(() => {
    let cancelled = false;
    setLoading(true);
    getNodePricingTiers(serviceType, region).then(options => {
      if (cancelled) return;
      setTiers(options);
      if (!current) {
        const defaultTier = selectPricingTier(options, getDefaultTier(serviceType)) ??
          options.find(tier => tier.skuName === getDefaultTier(serviceType)) ??
          selectPricingTier(options, getFallbackDefaultSku(serviceType)) ?? options[0];
        setSku(defaultTier?.id ?? defaultTier?.skuName ?? '');
      }
    }).catch(() => {
      if (!cancelled) setTiers([]);
    }).finally(() => {
      if (!cancelled) setLoading(false);
    });
    return () => { cancelled = true; };
  }, [serviceType, region, current]);

  const preview = useMemo((): { pricing?: NodePricingConfig; invalid: boolean } => {
    try {
      if (!quantity.trim()) throw new Error('Quantity is required');
      const count = Number(quantity);
      validatePricingQuantity(count);
      const selected = selectPricingTier(tiers, sku);
      let pricing: NodePricingConfig;
      if (current && !tierChanged && !(!custom && current.isCustom)) {
        pricing = { ...current, quantity: count, provenance: getPricingProvenance(current) };
      } else if (selected) {
        pricing = pricingFromTier(selected, count, region);
      } else if (tierChanged) {
        throw new Error('Invalid SKU');
      } else {
        pricing = {
          estimatedCost: null, tier: '', skuName: '', quantity: count, region,
          unit: 'USD/unit/month', isCustom: false, lastUpdated: new Date().toISOString(),
          provenance: getPricingProvenance(),
        };
      }
      if (custom) {
        if (!customPrice.trim()) throw new Error('Custom price is required');
        pricing = setCustomPricing(pricing, Number(customPrice));
      } else if (usageAmount !== undefined && pricing.usage) {
        pricing = withUsageAmount(pricing, usageAmount.trim() === '' ? null : Number(usageAmount));
      }
      if (pricing.estimatedCost !== null) validatePricingAmount(pricing.estimatedCost);
      return { pricing, invalid: false };
    } catch {
      return { invalid: true };
    }
  }, [quantity, tiers, sku, current, tierChanged, region, custom, customPrice, usageAmount]);

  const pricing = preview.pricing;
  const provenance = getPricingProvenance(pricing);
  const selected = selectPricingTier(tiers, sku);
  const fromTier = tierChanged || !current || (!custom && current.isCustom);
  const candidateUsage = custom ? undefined : (fromTier ? selected?.usage : current?.usage);
  const usageSource = fromTier ? selected?.provenance?.source : getPricingProvenance(current).source;
  const usage = isPricingUsage(candidateUsage) && usageSource === 'azure-retail-prices'
    ? candidateUsage : undefined;
  const outsideChange = initialData.current !== node.data;
  const invalid = preview.invalid || !label.trim() || label.trim().length > 200 ||
    description.length > 4000 || outsideChange;
  const money = (value: number) => formatCurrency(value, 'USD', language === 'ja' ? 'ja-JP' : 'en-US');
  const provenanceLabels = {
    'official-meter': text('Official retail meter', '公式小売メーター'),
    'usage-estimate': text('Usage estimate', '使用量に基づく見積もり'),
    'fallback-estimate': text('Fallback estimate', '代替の概算価格'),
    custom: text('Custom estimate', 'カスタム見積もり'),
    unpriced: text('Unpriced', '価格未設定'),
    capacity: text('Capacity billing', '容量課金'),
    unknown: text('Unknown source (imported)', '出典不明（インポート）'),
  };
  const assumptionText: Record<string, string> = {
    'Monthly operation': text('Monthly operation', '月間稼働時間'),
    'Monthly usage': text('Monthly usage', '月間使用量'),
    Capacity: text('Capacity', '容量'),
    'Bundled estimate level': text('Bundled estimate level', '組み込み見積もりレベル'),
    'Regional estimate multiplier': text('Regional estimate multiplier', '地域別概算係数'),
    'hours/month': text('hours/month', '時間/月'),
    'days/month': text('days/month', '日/月'),
    'years/month': text('years/month', '年/月'),
    months: text('months', 'か月'),
    'level (not a verified regional SKU)': text('level (not a verified regional SKU)', 'レベル（地域別SKU未確認）'),
    '× baseline': text('× baseline', '× 基準価格'),
  };

  function apply(event: React.FormEvent) {
    event.preventDefault();
    if (invalid || loading || !pricing) return;
    try {
      const data = buildServiceInspectorData(node.data, { label, description, pricing }, serviceType);
      onUpdateNode(node.id, data);
      onClose();
    } catch {
      setApplyError(text('Could not apply these settings. Check the values and try again.', '設定を適用できませんでした。値を確認して再試行してください。'));
    }
  }

  return (
    <div className="service-inspector-overlay" onClick={onClose}>
      <div className="service-inspector" ref={scopeRef} role="dialog" aria-modal="true"
        aria-labelledby="service-inspector-heading" tabIndex={-1} onClick={event => event.stopPropagation()}>
        <header>
          <div>
            <h2 id="service-inspector-heading">{text('Service inspector', 'サービスの詳細設定')}</h2>
            <p>{serviceType} · {region}</p>
          </div>
          <button type="button" className="service-inspector-close" onClick={onClose}
            aria-label={text('Close inspector', '詳細設定を閉じる')}><X size={20} /></button>
        </header>
        <form onSubmit={apply}>
          <div className="service-inspector-body">
            <label>
              {text('Label', 'ラベル')}
              <input data-modal-initial-focus required maxLength={200} value={label}
                onChange={event => setLabel(event.target.value)} />
            </label>
            <label>
              {text('Description', '説明')}
              <textarea aria-label={text('Description', '説明')} maxLength={4000} rows={3} value={description}
                onChange={event => setDescription(event.target.value)} />
            </label>
            <label>
              {text('Tier / SKU', 'レベル / SKU')}
              <select value={sku} disabled={loading || custom || !tiers.length}
                onChange={event => {
                  setSku(event.target.value); setTierChanged(true); setUsageAmount(undefined);
                }}>
                {!selected && <option value={sku}>{sku || text('No priced tier available', '価格情報のあるレベルがありません')}</option>}
                {tiers.map(tier => <option key={tier.id ?? tier.skuName} value={tier.id ?? tier.skuName}>
                  {tier.name === 'Included in capacity' ? text('Included in capacity', '容量に含まれる') : tier.name}
                </option>)}
              </select>
            </label>
            {loading && <p role="status">{text('Loading available prices…', '利用可能な価格を読み込み中…')}</p>}
            <label>
              {text('Quantity (instances / units)', '数量（インスタンス / 単位）')}
              <input type="number" required min={1} max={MAX_PRICING_QUANTITY} step={1}
                value={quantity} onChange={event => setQuantity(event.target.value)} />
            </label>
            <label className="service-inspector-checkbox">
              <input type="checkbox" checked={custom} onChange={event => {
                setCustom(event.target.checked); setUsageAmount(undefined);
              }} />
              {text('Use a custom monthly estimate', 'カスタム月額見積もりを使用')}
            </label>
            {custom && <label>
              {text('USD per unit / month', '単位あたりの月額（USD）')}
              <input type="number" required min={0} max={MAX_PRICING_AMOUNT} step="any" value={customPrice}
                onChange={event => setCustomPrice(event.target.value)} />
            </label>}
            {usage && <label>
              {text('Monthly usage per instance', 'インスタンスあたりの月間使用量')} ({usage.unit})
              <input type="number" min={0} max={MAX_PRICING_AMOUNT} step="any"
                placeholder={text('Unspecified', '未指定')}
                value={usageAmount ?? (usage.amount === null ? '' : String(usage.amount))}
                onChange={event => setUsageAmount(event.target.value)} />
              <small>{money(usage.unitPrice)} / {usage.unit} · {text('Enter billing units exactly as shown; other meters are excluded.', '表示された課金単位で入力してください。他のメーターは含まれません。')}</small>
            </label>}
            <section className="service-inspector-preview" aria-live="polite" aria-atomic="true">
              <h3>{text('Live estimate preview', '見積もりのプレビュー')}</h3>
              <span className="service-inspector-source">{provenanceLabels[provenance.kind]}</span>
              {pricing && pricing.estimatedCost !== null && provenance.kind !== 'unpriced' ? <>
                <p><strong>{money(pricing.estimatedCost * pricing.quantity)}</strong> {text('/ month', '/ 月')}</p>
                <p>{money(pricing.estimatedCost)} × {pricing.quantity} {text('units; quantity is applied once.', '単位。数量は1回だけ適用されます。')}</p>
              </> : <p>{provenance.kind === 'capacity'
                ? text('Included in shared capacity. Price the capacity separately; this is not free.', '共有容量に含まれます。容量は別途見積もってください。無料ではありません。')
                : text('No monthly price available. Excluded from the priced subtotal, not treated as zero.', '月額価格は未設定です。小計から除外され、ゼロとはみなされません。')}</p>}
              <p>{text('Price data date:', '価格データの日付:')} {provenance.asOf || text('Unknown', '不明')}</p>
              {provenance.snapshotAsOf && <p>{text('Regional file vintage (not the individual meter date):', '地域別ファイルの日付（個別メーターの日付ではありません）:')} {provenance.snapshotAsOf}</p>}
              {provenance.meterName && <p>{text('Meter:', 'メーター:')} {provenance.meterName}</p>}
              {provenance.meterId && <p>{text('Meter ID:', 'メーター ID:')} {provenance.meterId}</p>}
              {provenance.assumptions.length > 0 && <ul>
                {provenance.assumptions.map((assumption, index) =>
                  <li key={index}>{assumptionText[assumption.label] ?? assumption.label}: {assumption.value}{' '}
                    {assumptionText[assumption.unit] ?? assumption.unit}</li>)}
              </ul>}
              {provenance.kind === 'unknown' && <p>{text('Imported pricing has no verified source or usage assumptions.', 'インポート価格の出典や使用量の前提条件は確認されていません。')}</p>}
              {provenance.kind === 'fallback-estimate' && <p>{text('Bundled approximation, not a verified regional SKU quote. Detailed usage assumptions are unavailable.', '組み込みの概算であり、地域別のSKUの確定価格ではありません。詳細な使用量の前提条件は不明です。')}</p>}
              {provenance.kind === 'official-meter' && <p>{text('Selected meter only; additional meters and charges may apply.', '選択したメーターのみです。他のメーターや料金が適用される場合があります。')}</p>}
            </section>
            {outsideChange && <p className="service-inspector-error" role="alert">{text('This service changed while the inspector was open. Close and reopen it before applying.', '設定中にサービスが変更されました。閉じて再度開いてください。')}</p>}
            {preview.invalid && <p className="service-inspector-error" role="alert">{text(`Use a valid SKU, a whole quantity from 1 to ${MAX_PRICING_QUANTITY.toLocaleString('en-US')}, and finite nonnegative estimates.`, `有効なSKU、1〜${MAX_PRICING_QUANTITY.toLocaleString('ja-JP')}の整数の数量、および有限で非負の見積もり値を入力してください。`)}</p>}
            {applyError && <p className="service-inspector-error" role="alert">{applyError}</p>}
          </div>
          <footer>
            <button type="button" onClick={onClose}>{text('Cancel', 'キャンセル')}</button>
            <button type="submit" className="service-inspector-apply" disabled={invalid || loading}>{text('Apply', '適用')}</button>
          </footer>
        </form>
      </div>
    </div>
  );
}

export default function ServiceInspector(props: ServiceInspectorProps) {
  return <ServiceInspectorEditor key={props.node.id} {...props} />;
}
