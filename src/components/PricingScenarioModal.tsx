// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

import React, { useEffect, useMemo, useState } from 'react';
import type { Node } from 'reactflow';
import {
  CalendarDays,
  Copy,
  DollarSign,
  Download,
  Plus,
  RefreshCw,
  Trash2,
  TrendingDown,
  TrendingUp,
  X,
} from 'lucide-react';
import type { PricingCurrency, PricingScenario } from '../types/pricing';
import {
  DEFAULT_PRICING_SCENARIOS,
  calculatePricingScenario,
  createPricingScenario,
  defaultExchangeRate,
  exportPricingScenariosCsv,
  formatScenarioCurrency,
} from '../services/pricingScenarioService';
import { getPricingFreshness } from '../utils/pricingHelpers';
import { csvBlob } from '../utils/csvBlob';
import { useEscapeKey } from '../hooks/useEscapeKey';
import { useModalFocus } from '../hooks/useModalFocus';
import { useLanguage } from '../i18n/LanguageContext';
import { localize } from '../i18n/localization';
import './PricingScenarioModal.css';

interface PricingScenarioModalProps {
  isOpen: boolean;
  onClose: () => void;
  nodes: Node[];
  scenarios: PricingScenario[];
  onChange: (scenarios: PricingScenario[]) => void;
}

const CURRENCIES: PricingCurrency[] = ['USD', 'JPY', 'EUR', 'GBP'];

const PricingScenarioModal: React.FC<PricingScenarioModalProps> = ({
  isOpen,
  onClose,
  nodes,
  scenarios,
  onChange,
}) => {
  const { language, t } = useLanguage();
  const [selectedId, setSelectedId] = useState(scenarios[0]?.id || '');
  const [period, setPeriod] = useState<'monthly' | 'annual'>('monthly');
  const dialogRef = useModalFocus<HTMLDivElement>(isOpen);
  useEscapeKey(isOpen, onClose);

  useEffect(() => {
    if (!scenarios.some((scenario) => scenario.id === selectedId)) {
      setSelectedId(scenarios[0]?.id || '');
    }
  }, [scenarios, selectedId]);

  const results = useMemo(
    () => scenarios.map((scenario) => calculatePricingScenario(nodes, scenario)),
    [nodes, scenarios],
  );
  const selected = scenarios.find((scenario) => scenario.id === selectedId) || scenarios[0];
  const baseline = results.find((result) => result.scenario.kind === 'production')
    || results[0];
  const freshness = getPricingFreshness(
    results[0]?.pricesAsOf || '',
    new Date(),
    language,
  );
  const locale = language === 'ja' ? 'ja-JP' : 'en-US';

  const text = (en: string, ja: string) => localize(language, { en, ja });

  const updateSelected = (updates: Partial<PricingScenario>) => {
    if (!selected) return;
    onChange(scenarios.map((scenario) => (
      scenario.id === selected.id ? { ...scenario, ...updates } : scenario
    )));
  };

  const addScenario = (source?: PricingScenario) => {
    if (scenarios.length >= 12) return;
    const next = createPricingScenario(source);
    onChange([...scenarios, next]);
    setSelectedId(next.id);
  };

  const deleteSelected = () => {
    if (!selected || scenarios.length <= 1) return;
    if (!window.confirm(text(
      `Delete pricing scenario "${selected.name}"?`,
      `料金シナリオ「${selected.name}」を削除しますか？`,
    ))) return;
    const next = scenarios.filter((scenario) => scenario.id !== selected.id);
    onChange(next);
    setSelectedId(next[0]?.id || '');
  };

  const resetDefaults = () => {
    onChange(DEFAULT_PRICING_SCENARIOS.map((scenario) => ({ ...scenario })));
    setSelectedId(DEFAULT_PRICING_SCENARIOS[0].id);
  };

  const downloadComparison = () => {
    const csv = exportPricingScenariosCsv(results);
    const blob = csvBlob(csv);
    const url = URL.createObjectURL(blob);
    const link = document.createElement('a');
    link.href = url;
    link.download = `azurediagarm-pricing-scenarios-${new Date().toISOString().slice(0, 10)}.csv`;
    link.click();
    URL.revokeObjectURL(url);
  };

  if (!isOpen) return null;

  return (
    <div className="modal-overlay" onClick={onClose}>
      <div
        ref={dialogRef}
        className="modal-content pricing-scenario-modal"
        role="dialog"
        aria-modal="true"
        aria-label={text('Pricing scenario comparison', '料金シナリオ比較')}
        tabIndex={-1}
        onClick={(event) => event.stopPropagation()}
      >
        <div className="modal-header">
          <div>
            <h2><DollarSign size={24} /> {text('Pricing scenarios', '料金シナリオ')}</h2>
            <p>{text(
              'Compare development, production, and custom planning assumptions without changing the diagram.',
              '図面を変更せず、開発・本番・独自の計画条件を比較します。',
            )}</p>
          </div>
          <button className="modal-close" onClick={onClose} title={t('Close')} aria-label={t('Close')}>
            <X size={24} />
          </button>
        </div>

        <div className={`pricing-scenario-freshness ${freshness.isStale ? 'stale' : ''}`}>
          <CalendarDays size={16} />
          <span>
            {text('Azure prices as of', 'Azure価格の基準日')} {freshness.dateLabel}
            {' · '}
            {freshness.ageLabel}
          </span>
          <strong>{text(
            'FX rates are editable planning assumptions, not live market rates.',
            '為替レートは編集可能な計画値であり、リアルタイム市場レートではありません。',
          )}</strong>
        </div>

        <div className="pricing-period-toggle" role="group" aria-label={text('Display period', '表示期間')}>
          <button className={period === 'monthly' ? 'active' : ''} onClick={() => setPeriod('monthly')}>
            {text('Monthly', '月額')}
          </button>
          <button className={period === 'annual' ? 'active' : ''} onClick={() => setPeriod('annual')}>
            {text('Annual', '年額')}
          </button>
        </div>

        <div className="pricing-scenario-body">
          <section className="pricing-scenario-comparison">
            <div className="pricing-scenario-cards">
              {results.map((result) => {
                const value = period === 'monthly' ? result.totalMonthly : result.totalAnnual;
                const baselineUsd = period === 'monthly'
                  ? baseline?.totalMonthlyUsd || 0
                  : baseline?.totalAnnualUsd || 0;
                const currentUsd = period === 'monthly'
                  ? result.totalMonthlyUsd
                  : result.totalAnnualUsd;
                const delta = baselineUsd > 0
                  ? ((currentUsd - baselineUsd) / baselineUsd) * 100
                  : 0;
                return (
                  <button
                    key={result.scenario.id}
                    className={selected?.id === result.scenario.id ? 'selected' : ''}
                    onClick={() => setSelectedId(result.scenario.id)}
                  >
                    <span className="pricing-scenario-kind">{result.scenario.kind}</span>
                    <strong>{result.scenario.name}</strong>
                    <b>{formatScenarioCurrency(value, result.scenario.currency, locale)}</b>
                    <small>{period === 'monthly' ? text('per month', '月額') : text('per year', '年額')}</small>
                    {baseline && result.scenario.id !== baseline.scenario.id && (
                      <span className={`pricing-scenario-delta ${delta <= 0 ? 'saving' : 'increase'}`}>
                        {delta <= 0 ? <TrendingDown size={13} /> : <TrendingUp size={13} />}
                        {Math.abs(delta).toFixed(1)}% {text('vs production', '本番比')}
                      </span>
                    )}
                  </button>
                );
              })}
            </div>

            <div className="pricing-scenario-table-wrap">
              <table className="pricing-scenario-table">
                <thead>
                  <tr>
                    <th>{text('Scenario', 'シナリオ')}</th>
                    <th>{text('Term', '料金条件')}</th>
                    <th>{text('Capacity', '固定容量')}</th>
                    <th>{text('Usage', '使用量')}</th>
                    <th>{text('Discount', '割引')}</th>
                    <th>{text('Support', 'サポート')}</th>
                    <th>{period === 'monthly' ? text('Monthly', '月額') : text('Annual', '年額')}</th>
                  </tr>
                </thead>
                <tbody>
                  {results.map((result) => (
                    <tr key={result.scenario.id}>
                      <td>{result.scenario.name}</td>
                      <td>{result.scenario.pricingMode === 'reserved1yr' ? text('Savings 1yr', '1年Savings') : 'PAYG'}</td>
                      <td>{Math.round(result.scenario.capacityMultiplier * 100)}%</td>
                      <td>{Math.round(result.scenario.usageMultiplier * 100)}%</td>
                      <td>{result.scenario.discountPercent.toFixed(1)}%</td>
                      <td>{result.scenario.supportPercent.toFixed(1)}%</td>
                      <td>
                        {formatScenarioCurrency(
                          period === 'monthly' ? result.totalMonthly : result.totalAnnual,
                          result.scenario.currency,
                          locale,
                        )}
                      </td>
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
          </section>

          {selected && (
            <aside className="pricing-scenario-editor">
              <div className="pricing-scenario-editor-title">
                <h3>{text('Scenario assumptions', 'シナリオ条件')}</h3>
                <div>
                  <button onClick={() => addScenario(selected)} title={text('Duplicate', '複製')}>
                    <Copy size={15} />
                  </button>
                  <button onClick={deleteSelected} disabled={scenarios.length <= 1} title={text('Delete', '削除')}>
                    <Trash2 size={15} />
                  </button>
                </div>
              </div>

              <label>
                <span>{text('Name', '名前')}</span>
                <input
                  value={selected.name}
                  maxLength={80}
                  onChange={(event) => updateSelected({ name: event.target.value })}
                  onBlur={() => {
                    if (!selected.name.trim()) {
                      updateSelected({ name: text('Custom scenario', 'カスタム シナリオ') });
                    }
                  }}
                />
              </label>

              <label>
                <span>{text('Azure pricing term', 'Azure料金条件')}</span>
                <select
                  value={selected.pricingMode}
                  onChange={(event) => updateSelected({
                    pricingMode: event.target.value as PricingScenario['pricingMode'],
                  })}
                >
                  <option value="payg">{text('Pay-as-you-go', '従量課金')}</option>
                  <option value="reserved1yr">{text('1-year Savings Plan / representative reservation', '1年Savings Plan / 代表的な予約割引')}</option>
                </select>
              </label>

              <div className="pricing-scenario-field-grid">
                <label>
                  <span>{text('Fixed capacity', '固定容量')} (%)</span>
                  <input
                    type="number"
                    min={0}
                    max={2000}
                    step={5}
                    value={Math.round(selected.capacityMultiplier * 100)}
                    onChange={(event) => updateSelected({
                      capacityMultiplier: Math.max(
                        0,
                        Math.min(20, Number(event.target.value) / 100),
                      ),
                    })}
                  />
                </label>
                <label>
                  <span>{text('Usage level', '使用量')} (%)</span>
                  <input
                    type="number"
                    min={0}
                    max={2000}
                    step={5}
                    value={Math.round(selected.usageMultiplier * 100)}
                    onChange={(event) => updateSelected({
                      usageMultiplier: Math.max(
                        0,
                        Math.min(20, Number(event.target.value) / 100),
                      ),
                    })}
                  />
                </label>
                <label>
                  <span>{text('Negotiated discount', '契約割引')} (%)</span>
                  <input
                    type="number"
                    min={0}
                    max={90}
                    step={0.5}
                    value={selected.discountPercent}
                    onChange={(event) => updateSelected({
                      discountPercent: Math.max(0, Math.min(90, Number(event.target.value))),
                    })}
                  />
                </label>
                <label>
                  <span>{text('Support allowance', 'サポート加算')} (%)</span>
                  <input
                    type="number"
                    min={0}
                    max={50}
                    step={0.5}
                    value={selected.supportPercent}
                    onChange={(event) => updateSelected({
                      supportPercent: Math.max(0, Math.min(50, Number(event.target.value))),
                    })}
                  />
                </label>
              </div>

              <div className="pricing-scenario-field-grid">
                <label>
                  <span>{text('Currency', '通貨')}</span>
                  <select
                    value={selected.currency}
                    onChange={(event) => {
                      const currency = event.target.value as PricingCurrency;
                      updateSelected({
                        currency,
                        exchangeRate: defaultExchangeRate(currency),
                      });
                    }}
                  >
                    {CURRENCIES.map((currency) => <option key={currency}>{currency}</option>)}
                  </select>
                </label>
                <label>
                  <span>{text('Planning FX per USD', '1 USDあたり計画為替')}</span>
                  <input
                    type="number"
                    min={0.0001}
                    step={0.01}
                    disabled={selected.currency === 'USD'}
                    value={selected.currency === 'USD' ? 1 : selected.exchangeRate}
                    onChange={(event) => updateSelected({
                      exchangeRate: Math.max(
                        0.0001,
                        Math.min(100_000, Number(event.target.value)),
                      ),
                    })}
                  />
                </label>
              </div>

              <div className="pricing-scenario-formula">
                <strong>{text('Calculation order', '計算順序')}</strong>
                <span>{text(
                  'Azure term → capacity/usage scale → negotiated discount → support allowance → planning FX',
                  'Azure料金条件 → 容量/使用量倍率 → 契約割引 → サポート加算 → 計画為替',
                )}</span>
              </div>
            </aside>
          )}
        </div>

        <div className="modal-actions">
          <button className="btn-secondary" onClick={resetDefaults}>
            <RefreshCw size={17} />
            {text('Reset defaults', '既定値へ戻す')}
          </button>
          <button className="btn-secondary" onClick={() => addScenario()} disabled={scenarios.length >= 12}>
            <Plus size={17} />
            {text('Add scenario', 'シナリオを追加')}
          </button>
          <button className="btn-secondary" onClick={downloadComparison} disabled={results.length === 0}>
            <Download size={17} />
            {text('Export CSV', 'CSV出力')}
          </button>
          <button className="btn-primary" onClick={onClose}>{t('Close')}</button>
        </div>
      </div>
    </div>
  );
};

export default PricingScenarioModal;
