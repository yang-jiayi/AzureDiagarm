// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

import { ChevronDown, RotateCcw, SlidersHorizontal, Sparkles } from 'lucide-react';
import {
  FEATURE_CONFIG, getAvailableModels, getReasoningEffortLabel, getRecommendedModelSettings,
  getSupportedReasoningEfforts, updateFeatureOverride, useModelSettings, type FeatureType,
} from '../stores/modelSettingsStore';
import { useLanguage } from '../i18n/LanguageContext';
import { getBYOAISettings, useBYOAISettings } from '../stores/byoAISettingsStore';
import { localize } from '../i18n/localization';
import './ModelSettingsPopover.css';

export default function AstraReasoningSettings() {
  const { t, translate, language } = useLanguage();
  const [settings, updateSettings] = useModelSettings();
  useBYOAISettings();
  const byoActive = getBYOAISettings().activeProfileId !== null;
  const configured = getAvailableModels().includes('gpt-6-astra');
  const efforts = getSupportedReasoningEfforts('gpt-6-astra');
  const features = Object.keys(FEATURE_CONFIG) as FeatureType[];
  const hasOverrides = features.some(feature => settings.featureOverrides?.[feature] !== undefined);

  return (
    <div className="astra-reasoning-settings">
      <p className="msp-managed-note">
        <strong>GPT-6 Astra</strong>
        {' · '}{t('ai.astraOnlyDescription')}
      </p>
      {byoActive && <p className="msp-managed-note">{localize(language, {
        en: 'Managed-only settings. Your selected BYO profile uses its own reasoning, capabilities, and output limit; changing these Astra defaults does not change the active connection.',
        ja: '管理対象専用の設定です。選択中の BYO プロファイルは独自の推論・機能・出力上限を使用します。Astra の既定値を変更しても、選択中の接続は変わりません。',
      })}</p>}
      {!configured && <p className="azd-callout azd-callout--danger" role="alert">{t('ai.astraNotConfigured')}</p>}
      <div className="msp-reasoning-row">
        <span className="msp-reasoning-label">{t('ai.defaultReasoning')}</span>
        <div className="msp-reasoning-buttons" role="group" aria-label={t('ai.defaultReasoning')}>
          {efforts.map(level => (
            <button type="button" key={level}
              className={`msp-reasoning-btn${settings.reasoningEffort === level ? ' active' : ''}`}
              aria-pressed={settings.reasoningEffort === level}
              disabled={!configured}
              onClick={() => updateSettings({ model: 'gpt-6-astra', reasoningEffort: level })}
            >{t(getReasoningEffortLabel(level))}</button>
          ))}
        </div>
      </div>
      <details className="msp-advanced-settings">
        <summary>
          <span><SlidersHorizontal size={14} aria-hidden="true" />
            {t('ai.perFeatureReasoning')}</span>
          <span className="msp-advanced-summary-status">
            {t(hasOverrides ? 'ai.customized' : 'ai.usingDefaults')}
            <ChevronDown size={14} aria-hidden="true" />
          </span>
        </summary>
        <div className="msp-advanced-body">
          <div className="msp-portfolio">
            <button type="button" className="msp-portfolio-btn" disabled={!configured}
              onClick={() => updateSettings(getRecommendedModelSettings())}>
              <Sparkles size={12} />
              {t('ai.recommendedReasoning')}
            </button>
            {hasOverrides && <button type="button" className="msp-reset-btn" disabled={!configured}
              onClick={() => updateSettings({ featureOverrides: {} })}
              title={t('ai.resetFeatureReasoning')} aria-label={t('ai.resetFeatureReasoning')}><RotateCcw size={14} /></button>}
          </div>
          <div className="msp-features">
            {features.map(feature => {
              const name = translate(FEATURE_CONFIG[feature].displayName);
              const override = settings.featureOverrides?.[feature];
              return (
                <label className={`msp-feature-row${override ? ' overridden' : ''}`} key={feature}>
                  <span className="msp-feature-name">{name}</span>
                  <select className="msp-feature-select"
                    aria-label={`${name} - ${t('ai.reasoningEffort')}`}
                    value={override?.reasoningEffort ?? 'default'} disabled={!configured}
                    onChange={event => {
                      if (event.target.value === 'default') updateFeatureOverride(feature, null);
                      else {
                        const level = efforts.find(value => value === event.target.value);
                        if (level !== undefined) updateFeatureOverride(feature, { model: 'gpt-6-astra', reasoningEffort: level });
                      }
                    }}>
                    <option value="default">{t('ai.defaultReasoningChoice', { effort: t(getReasoningEffortLabel(settings.reasoningEffort)) })}</option>
                    {efforts.map(level => <option key={level} value={level}>{t(getReasoningEffortLabel(level))}</option>)}
                  </select>
                </label>
              );
            })}
          </div>
        </div>
      </details>
    </div>
  );
}
