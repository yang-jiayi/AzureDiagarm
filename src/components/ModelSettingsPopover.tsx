// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

/**
 * Model Settings Popover
 * Toolbar dropdown for AI model selection across all features.
 * Follows the same toolbar-dropdown pattern as Layout/Export menus.
 */

import { forwardRef } from 'react';
import { Brain, RotateCcw, ChevronDown, Cpu, Layers, Sparkles, Zap } from 'lucide-react';
import {
  useModelSettings,
  MODEL_CONFIG,
  ModelType,
  ReasoningEffort,
  FeatureType,
  FEATURE_CONFIG,
  getAvailableModels,
  updateFeatureOverride,
  hasFeatureOverride,
} from '../stores/modelSettingsStore';
import './ModelSettingsPopover.css';
import { useLanguage } from '../i18n/LanguageContext';
import type { TranslationKey } from '../i18n/LanguageContext';

interface ModelSettingsPopoverProps {
  isOpen: boolean;
  onToggle: () => void;
}

const REASONING_LABEL_KEYS: Record<ReasoningEffort, TranslationKey> = {
  none: 'None',
  low: 'Low',
  medium: 'Medium',
  high: 'High',
};

const ModelSettingsPopover = forwardRef<HTMLDivElement, ModelSettingsPopoverProps>(
  ({ isOpen, onToggle }, ref) => {
    const { t } = useLanguage();
    const [settings, updateSettings] = useModelSettings();
    const availableModels = getAvailableModels();
    const currentConfig = MODEL_CONFIG[settings.model];

    const hasAnyOverride = (Object.keys(FEATURE_CONFIG) as FeatureType[]).some(hasFeatureOverride);

    const handleModelChange = (model: ModelType) => {
      const config = MODEL_CONFIG[model];
      // Auto-set reasoning effort to model's default when switching
      if (config.defaultReasoningEffort !== undefined) {
        updateSettings({ model, reasoningEffort: config.defaultReasoningEffort });
      } else {
        // For models without a default, keep current effort but ensure it's valid
        const currentEffort = settings.reasoningEffort;
        if (!config.isReasoning || (currentEffort === 'none' && !config.defaultReasoningEffort)) {
          updateSettings({ model, reasoningEffort: 'medium' });
        } else {
          updateSettings({ model });
        }
      }
    };

    const handleReasoningChange = (reasoning: ReasoningEffort) => {
      updateSettings({ reasoningEffort: reasoning });
    };

    const handleFeatureModelChange = (feature: FeatureType, value: string) => {
      if (value === 'default') {
        updateFeatureOverride(feature, null);
      } else {
        const model = value as ModelType;
        const currentOverride = settings.featureOverrides?.[feature];
        updateFeatureOverride(feature, {
          model,
          reasoningEffort: currentOverride?.reasoningEffort,
        });
      }
    };

    const handleFeatureReasoningChange = (feature: FeatureType, value: ReasoningEffort) => {
      const currentOverride = settings.featureOverrides?.[feature];
      if (currentOverride) {
        updateFeatureOverride(feature, {
          ...currentOverride,
          reasoningEffort: value,
        });
      }
    };

    const resetAllOverrides = () => {
      updateSettings({ featureOverrides: {} });
    };

    const getModelIcon = (model: ModelType) => {
      switch (model) {
        case 'gpt-5.1':
          return <Cpu size={14} />;
        case 'gpt-5.2':
          return <Brain size={14} />;
        case 'gpt-5.6-sol':
          return <Brain size={14} />;
        case 'gpt-5.6-terra':
          return <Brain size={14} />;
        case 'gpt-5.6-luna':
          return <Brain size={14} />;
        case 'deepseek-v3.2-speciale':
          return <Layers size={14} />;
        case 'grok-4.1-fast':
          return <Zap size={14} />;
        case 'gpt-5.4-mini':
          return <Sparkles size={14} />;
        default:
          return <Cpu size={14} />;
      }
    };

    const getFeatureCurrentModel = (feature: FeatureType): string => {
      const override = settings.featureOverrides?.[feature];
      return override ? override.model : 'default';
    };

    const getFeatureCurrentReasoning = (feature: FeatureType): ReasoningEffort => {
      const override = settings.featureOverrides?.[feature];
      return override?.reasoningEffort || settings.reasoningEffort;
    };

    // Compute the effective model for each feature (for display)
    const getEffectiveModel = (feature: FeatureType): ModelType => {
      const override = settings.featureOverrides?.[feature];
      return override ? override.model : settings.model;
    };

    const getEffectiveReasoning = (feature: FeatureType): ReasoningEffort | null => {
      const model = getEffectiveModel(feature);
      if (!MODEL_CONFIG[model].isReasoning) return null;
      const override = settings.featureOverrides?.[feature];
      return override?.reasoningEffort || settings.reasoningEffort;
    };

    return (
      <div className="toolbar-dropdown" ref={ref}>
        <button
          onClick={onToggle}
          className="btn btn-secondary model-popover-trigger"
          title={t("AI model settings")}
          aria-haspopup="menu"
          aria-expanded={isOpen}
        >
          {getModelIcon(settings.model)}
          <span className="model-popover-label">{currentConfig.displayName}</span>
          {currentConfig.isReasoning && (
            <span className="model-popover-reasoning">{t(REASONING_LABEL_KEYS[settings.reasoningEffort])}</span>
          )}
          {hasAnyOverride && <span className="model-popover-override-dot" />}
          <ChevronDown size={14} style={{ marginLeft: 2 }} />
        </button>

        {isOpen && (
          <div
            className="toolbar-dropdown-menu toolbar-dropdown-menu--model-settings"
            role="menu"
            aria-label={t("AI model settings")}
          >
            {/* Default Model */}
            <div className="toolbar-dropdown-heading">{t("Default Model")}</div>
            <div className="msp-model-buttons">
              {availableModels.map((model) => (
                <button
                  key={model}
                  className={`msp-model-btn ${settings.model === model ? 'active' : ''}`}
                  onClick={() => handleModelChange(model)}
                  title={MODEL_CONFIG[model].description}
                >
                  {getModelIcon(model)}
                  <span>{MODEL_CONFIG[model].displayName}</span>
                </button>
              ))}
            </div>

            {/* Reasoning effort (only shown when reasoning model selected) */}
            {currentConfig.isReasoning && (
              <>
                <div className="msp-reasoning-row">
                  <span className="msp-reasoning-label">{t("Reasoning")}</span>
                  <div className="msp-reasoning-buttons">
                    {(['none', 'low', 'medium', 'high'] as ReasoningEffort[]).map((level) => (
                      <button
                        key={level}
                        className={`msp-reasoning-btn ${settings.reasoningEffort === level ? 'active' : ''}`}
                        onClick={() => handleReasoningChange(level)}
                        title={level === 'none' ? t("No reasoning - fastest response") : undefined}
                      >
                        {t(REASONING_LABEL_KEYS[level])}
                      </button>
                    ))}
                  </div>
                </div>
              </>
            )}

            <div className="toolbar-dropdown-separator" role="separator" />

            {/* Per-Feature Overrides */}
            <div className="toolbar-dropdown-heading">
              {' '}{t("Per-Feature Settings")}{' '}{hasAnyOverride && (
                <button className="msp-reset-btn" onClick={resetAllOverrides} title={t("Reset all to default")}>
                  <RotateCcw size={11} />
                </button>
              )}
            </div>

            <div className="msp-features">
              {(Object.keys(FEATURE_CONFIG) as FeatureType[]).map((feature) => {
                const featureConfig = FEATURE_CONFIG[feature];
                const currentModel = getFeatureCurrentModel(feature);
                const currentReasoning = getFeatureCurrentReasoning(feature);
                const isOverridden = currentModel !== 'default';
                const selectedModelConfig = isOverridden ? MODEL_CONFIG[currentModel as ModelType] : null;
                const effectiveModel = getEffectiveModel(feature);
                const effectiveReasoning = getEffectiveReasoning(feature);

                return (
                  <div key={feature} className={`msp-feature-row ${isOverridden ? 'overridden' : ''}`}>
                    <div className="msp-feature-info">
                      <span className="msp-feature-name">{featureConfig.displayName}</span>
                      <span className="msp-feature-effective">
                        {MODEL_CONFIG[effectiveModel].displayName}
                        {effectiveReasoning && ` (${t(REASONING_LABEL_KEYS[effectiveReasoning])})`}
                      </span>
                    </div>
                    <div className="msp-feature-controls">
                      <select
                        value={currentModel}
                        onChange={(e) => handleFeatureModelChange(feature, e.target.value)}
                        className="msp-feature-select"
                      >
                        <option value="default">{t("Default")}</option>
                        {availableModels.map((model) => (
                          <option key={model} value={model}>
                            {MODEL_CONFIG[model].displayName}
                          </option>
                        ))}
                      </select>

                      {isOverridden && selectedModelConfig?.isReasoning && (
                        <select
                          value={currentReasoning}
                          onChange={(e) =>
                            handleFeatureReasoningChange(feature, e.target.value as ReasoningEffort)
                          }
                          className="msp-reasoning-select"
                        >
                          <option value="none">{t("None")}</option>
                          <option value="low">{t("Low")}</option>
                          <option value="medium">{t("Med")}</option>
                          <option value="high">{t("High")}</option>
                        </select>
                      )}
                    </div>
                  </div>
                );
              })}
            </div>

            <div className="toolbar-dropdown-separator" role="separator" />

            <div className="toolbar-dropdown-hint">
              {' '}{t("Recommended: GPT-5.4 Mini (low) for cost-efficient generation and validation")}{' '}</div>
          </div>
        )}
      </div>
    );
  }
);

ModelSettingsPopover.displayName = 'ModelSettingsPopover';

export default ModelSettingsPopover;
