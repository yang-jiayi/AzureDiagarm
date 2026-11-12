// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

import { ChevronDown, RotateCcw, SlidersHorizontal, Sparkles } from 'lucide-react';
import {
  FEATURE_CONFIG, getAvailableModels, getReasoningEffortLabel, getRecommendedModelSettings,
  getSupportedReasoningEfforts, updateFeatureOverride, useModelSettings, isModelType, normalizeReasoningEffort, MODEL_CONFIG, type FeatureType,
} from '../stores/modelSettingsStore';
import { useLanguage } from '../i18n/LanguageContext';
import { getBYOAISettings, useBYOAISettings } from '../stores/byoAISettingsStore';
import { localize } from '../i18n/localization';
import './ModelSettingsPopover.css';

export default function ManagedModelSettings() {
  const { t, translate, language } = useLanguage();
  const [settings, updateSettings] = useModelSettings();
  useBYOAISettings();
  const byoActive = getBYOAISettings().activeProfileId !== null;
  const availableModels = getAvailableModels();
  const models = Object.keys(MODEL_CONFIG).filter(isModelType);
  const configured = availableModels.includes(settings.model);
  const efforts = getSupportedReasoningEfforts(settings.model);
  const features = Object.keys(FEATURE_CONFIG) as FeatureType[];
  const hasOverrides = features.some(feature => settings.featureOverrides?.[feature] !== undefined);

  return (
    <div className="managed-model-settings">
      <p className="msp-managed-note">
        {t('ai.managedDescription')}
      </p>
      <label className="msp-feature-row">
        <span>{localize(language, { en: 'Managed model', ja: '管理対象モデル' })}</span>
        <select className="msp-feature-select" disabled={availableModels.length === 0} aria-label={localize(language, { en: 'Managed model', ja: '管理対象モデル' })}
          value={settings.model} onChange={event => {
            if (isModelType(event.target.value)) updateSettings({ model: event.target.value });
          }}>
          {models.map(model => <option key={model} value={model} disabled={!availableModels.includes(model)}>
            {MODEL_CONFIG[model].displayName}{availableModels.includes(model) ? '' : localize(language, { en: ' (not configured)', ja: '（未設定）' })}
          </option>)}
        </select>
      </label>
      {byoActive && <p className="msp-managed-note">{localize(language, {
        en: 'Managed-only settings. Your selected BYO profile uses its own reasoning, capabilities, and output limit; changing these defaults does not change the active connection.',
        ja: '管理対象専用の設定です。選択中の BYO プロファイルは独自の推論・機能・出力上限を使用します。この既定値を変更しても、選択中の接続は変わりません。',
      })}</p>}
      {!configured && <p className="azd-callout azd-callout--danger" role="alert">{t('ai.managedNotConfigured')}</p>}
      <div className="msp-reasoning-row">
        <span className="msp-reasoning-label">{t('ai.defaultReasoning')}</span>
        <div className="msp-reasoning-buttons" role="group" aria-label={t('ai.defaultReasoning')}>
          {efforts.map(level => (
            <button type="button" key={level}
              className={`msp-reasoning-btn${settings.reasoningEffort === level ? ' active' : ''}`}
              aria-pressed={settings.reasoningEffort === level}
              disabled={!configured}
              onClick={() => updateSettings({ reasoningEffort: level })}
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
              const model = override?.model ?? settings.model;
              const featureEfforts = getSupportedReasoningEfforts(model);
              return (
                <div className={`msp-feature-row${override ? ' overridden' : ''}`} key={feature}>
                  <span className="msp-feature-name">{name}</span>
                  <select className="msp-feature-select" disabled={availableModels.length === 0} aria-label={`${name} - ${localize(language, { en: 'Model', ja: 'モデル' })}`}
                    value={override?.model ?? 'default'} onChange={event => {
                      if (event.target.value === 'default') updateFeatureOverride(feature, null);
                      else if (isModelType(event.target.value)) updateFeatureOverride(feature, {
                        model: event.target.value, reasoningEffort: override?.reasoningEffort,
                      });
                    }}>
                    <option value="default">{localize(language, { en: 'Default model', ja: '既定のモデル' })}</option>
                    {models.map(item => <option key={item} value={item} disabled={!availableModels.includes(item)}>{MODEL_CONFIG[item].displayName}</option>)}
                  </select>
                  <select className="msp-feature-select"
                    aria-label={`${name} - ${t('ai.reasoningEffort')}`}
                    value={override?.reasoningEffort ?? 'default'} disabled={!availableModels.includes(model)}
                    onChange={event => {
                      if (event.target.value === 'default') updateFeatureOverride(feature, override ? { model: override.model } : null);
                      else {
                        const level = featureEfforts.find(value => value === event.target.value);
                        if (level !== undefined) updateFeatureOverride(feature, { model, reasoningEffort: level });
                      }
                    }}>
                    <option value="default">{t('ai.defaultReasoningChoice', { effort: t(getReasoningEffortLabel(normalizeReasoningEffort(model, settings.reasoningEffort))) })}</option>
                    {featureEfforts.map(level => <option key={level} value={level}>{t(getReasoningEffortLabel(level))}</option>)}
                  </select>
                </div>
              );
            })}
          </div>
        </div>
      </details>
    </div>
  );
}
