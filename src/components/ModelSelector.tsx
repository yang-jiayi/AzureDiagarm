// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

/**
 * Model Selector Component
 * Allows users to select AI model and reasoning effort level
 * Includes Advanced section for per-feature model overrides
 */

import React, { useState } from 'react';
import { Settings, Brain, ChevronDown, ChevronRight, RotateCcw, Cpu, Layers, Sparkles, Zap } from 'lucide-react';
import { 
  useModelSettings, 
  MODEL_CONFIG, 
  ModelType, 
  ReasoningEffort,
  FeatureType,
  FEATURE_CONFIG,
  getAvailableModels,
  getReasoningEffortLabel,
  getSupportedReasoningEfforts,
  updateFeatureOverride,
  hasFeatureOverride
} from '../stores/modelSettingsStore';
import './ModelSelector.css';
import { useLanguage } from '../i18n/LanguageContext';

interface ModelSelectorProps {
  compact?: boolean;
}

const ModelSelector: React.FC<ModelSelectorProps> = ({ compact = false }) => {
  const { t, translate } = useLanguage();
  const [settings, updateSettings] = useModelSettings();
  const [showAdvanced, setShowAdvanced] = useState(false);
  const availableModels = getAvailableModels();
  
  const currentConfig = MODEL_CONFIG[settings.model];
  const supportedReasoningEfforts = getSupportedReasoningEfforts(settings.model);
  
  // Check if any feature has overrides
  const hasAnyOverride = (Object.keys(FEATURE_CONFIG) as FeatureType[]).some(hasFeatureOverride);
  
  const handleModelChange = (e: React.ChangeEvent<HTMLSelectElement>) => {
    const model = e.target.value as ModelType;
    updateSettings({
      model,
      reasoningEffort: MODEL_CONFIG[model].defaultReasoningEffort ?? settings.reasoningEffort,
    });
  };
  
  const handleReasoningChange = (e: React.ChangeEvent<HTMLSelectElement>) => {
    updateSettings({ reasoningEffort: e.target.value as ReasoningEffort });
  };

  const handleFeatureModelChange = (feature: FeatureType, value: string) => {
    if (value === 'default') {
      updateFeatureOverride(feature, null);
    } else {
      const model = value as ModelType;
      const currentOverride = settings.featureOverrides?.[feature];
      updateFeatureOverride(feature, {
        model,
        reasoningEffort: currentOverride?.reasoningEffort
      });
    }
  };

  const handleFeatureReasoningChange = (feature: FeatureType, value: ReasoningEffort) => {
    const currentOverride = settings.featureOverrides?.[feature];
    if (currentOverride) {
      updateFeatureOverride(feature, {
        ...currentOverride,
        reasoningEffort: value
      });
    }
  };

  const resetAllOverrides = () => {
    updateSettings({ featureOverrides: {} });
  };

  const getModelIcon = (model: ModelType) => {
    switch (model) {
      case 'gpt-5.1':
        return <Cpu size={16} />;
      case 'gpt-5.2':
        return <Brain size={16} />;
      case 'gpt-5.6-sol':
        return <Brain size={16} />;
      case 'gpt-5.6-terra':
        return <Brain size={16} />;
      case 'gpt-5.6-luna':
        return <Brain size={16} />;
      case 'deepseek-v3.2-speciale':
        return <Layers size={16} />;
      case 'grok-4.1-fast':
        return <Zap size={16} />;
      case 'gpt-5.4-mini':
        return <Sparkles size={16} />;

      default:
        return <Settings size={16} />;
    }
  };

  const getFeatureCurrentModel = (feature: FeatureType): string => {
    const override = settings.featureOverrides?.[feature];
    if (!override) return 'default';
    return override.model;
  };

  const getFeatureCurrentReasoning = (feature: FeatureType): ReasoningEffort => {
    const override = settings.featureOverrides?.[feature];
    return override?.reasoningEffort || settings.reasoningEffort;
  };

  if (compact) {
    return (
      <div className="model-selector compact">
        <div className="model-selector-with-help">
          <div className="model-selector-controls">
            <div className="model-selector-row">
              <select 
                value={settings.model} 
                onChange={handleModelChange}
                className="model-select"
                title={translate(currentConfig.description)}
              >
                {availableModels.map(model => (
                  <option key={model} value={model}>
                    {MODEL_CONFIG[model].displayName}
                  </option>
                ))}
              </select>
              
              {currentConfig.isReasoning && (
                <select
                  value={settings.reasoningEffort}
                  onChange={handleReasoningChange}
                  className="reasoning-select"
                  title={t("Reasoning effort level")}
                >
                  {supportedReasoningEfforts.map(level => (
                    <option key={level} value={level}>
                      {t(getReasoningEffortLabel(level))}
                    </option>
                  ))}
                </select>
              )}
              
              <button
                className={`advanced-toggle-compact ${showAdvanced ? 'expanded' : ''} ${hasAnyOverride ? 'has-overrides' : ''}`}
                onClick={() => setShowAdvanced(!showAdvanced)}
                title={t("Per-feature model settings")}
              >
                <Settings size={14} />
                {hasAnyOverride && <span className="override-dot" />}
              </button>
            </div>
            
            {showAdvanced && (
              <div className="compact-advanced-panel">
                <div className="compact-advanced-header">
                  <span>{t("Per-Feature Model Overrides")}</span>
                  {hasAnyOverride && (
                    <button className="reset-overrides-compact" onClick={resetAllOverrides} title={t("Reset all to default")}>
                      <RotateCcw size={12} />
                    </button>
                  )}
                </div>
                
                {(Object.keys(FEATURE_CONFIG) as FeatureType[]).map(feature => {
                  const featureConfig = FEATURE_CONFIG[feature];
                  const currentModel = getFeatureCurrentModel(feature);
                  const currentReasoning = getFeatureCurrentReasoning(feature);
                  const isOverridden = currentModel !== 'default';
                  const selectedModelConfig = isOverridden ? MODEL_CONFIG[currentModel as ModelType] : null;
                  
                  return (
                    <div key={feature} className={`compact-feature-row ${isOverridden ? 'overridden' : ''}`}>
                      <span className="compact-feature-label">{translate(featureConfig.displayName)}</span>
                      <div className="compact-feature-controls">
                        <select
                          value={currentModel}
                          onChange={(e) => handleFeatureModelChange(feature, e.target.value)}
                          className="compact-feature-select"
                        >
                          <option value="default">{t("Default (")}{MODEL_CONFIG[settings.model].displayName}{t(")")}</option>
                          {availableModels.map(model => (
                            <option key={model} value={model}>
                              {MODEL_CONFIG[model].displayName}
                            </option>
                          ))}
                        </select>
                        
                        {isOverridden && selectedModelConfig?.isReasoning && (
                          <select
                            value={currentReasoning}
                            onChange={(e) => handleFeatureReasoningChange(feature, e.target.value as ReasoningEffort)}
                            className="compact-reasoning-select"
                          >
                            {getSupportedReasoningEfforts(currentModel as ModelType).map(level => (
                              <option key={level} value={level}>
                                {t(getReasoningEffortLabel(level))}
                              </option>
                            ))}
                          </select>
                        )}
                      </div>
                    </div>
                  );
                })}
                
                <div className="compact-advanced-footer">
                  <span className="compact-hint">{t("Rec:")}{' '}{FEATURE_CONFIG.architectureGeneration.recommendedModel} {' '}{t("for generation")}</span>
                </div>
              </div>
            )}
          </div>
          
          <div className="model-help-panel">
            <div className="help-item">
              <span className="help-label">{t("Reasoning:")}</span>
              <span className="help-text">{t("None = fastest; Low to Max increases reasoning depth")}</span>
            </div>
            <div className="help-item">
              <span className="help-icon"><Settings size={10} /></span>
              <span className="help-text">{t("Click gear for per-feature model overrides")}</span>
            </div>
            <div className="help-divider" />
            <div className="help-subtitle">{t("Recommended settings (from testing):")}</div>
            <div className="help-defaults">
              <span>{t("• Architecture: GPT-5.6 Sol (low)")}</span>
              <span>{t("• Validation: GPT-5.6 Sol (low)")}</span>
              <span>{t("• Deployment: GPT-5.6 Sol (low)")}</span>
            </div>
          </div>
        </div>
      </div>
    );
  }

  return (
    <div className="model-selector">
      <div className="model-selector-header">
        <Settings size={16} />
        <span>{t("AI Model Settings")}</span>
      </div>
      
      <div className="model-selector-content">
        <div className="model-selector-group">
          <label className="model-label">{t("Default Model")}</label>
          <div className="model-buttons">
            {availableModels.map(model => (
              <button
                key={model}
                className={`model-button ${settings.model === model ? 'active' : ''}`}
                onClick={() => updateSettings({ model })}
                title={translate(MODEL_CONFIG[model].description)}
              >
                {getModelIcon(model)}
                <span>{MODEL_CONFIG[model].displayName}</span>
              </button>
            ))}
          </div>
          <p className="model-description">{translate(currentConfig.description)}</p>
        </div>
        
        {currentConfig.isReasoning && (
          <div className="model-selector-group">
            <label className="model-label">{t("Default Reasoning Effort")}</label>
            <div className="reasoning-buttons">
              {supportedReasoningEfforts.map(level => (
                <button
                  key={level}
                  className={`reasoning-button ${settings.reasoningEffort === level ? 'active' : ''}`}
                  onClick={() => updateSettings({ reasoningEffort: level })}
                >
                  {t(getReasoningEffortLabel(level))}
                </button>
              ))}
            </div>
            <p className="reasoning-hint">
              {settings.reasoningEffort === 'none' && t("No reasoning - fastest response, lowest cost")}
              {settings.reasoningEffort === 'low' && t("Faster responses, less detailed analysis")}
              {settings.reasoningEffort === 'medium' && t("Balanced speed and depth")}
              {settings.reasoningEffort === 'high' && t("Thorough analysis, may take longer")}
              {settings.reasoningEffort === 'xhigh' && t("Extra-thorough analysis with higher latency and token use")}
              {settings.reasoningEffort === 'max' && t("Maximum reasoning depth with the longest latency and highest token use")}
            </p>
          </div>
        )}

        {/* Advanced Per-Feature Settings */}
        <div className="advanced-section">
          <button 
            className={`advanced-toggle ${showAdvanced ? 'expanded' : ''}`}
            onClick={() => setShowAdvanced(!showAdvanced)}
          >
            {showAdvanced ? <ChevronDown size={14} /> : <ChevronRight size={14} />}
            <span>{t("Advanced (per-feature)")}</span>
            {hasAnyOverride && <span className="override-badge">{Object.keys(settings.featureOverrides || {}).length}</span>}
          </button>
          
          {showAdvanced && (
            <div className="advanced-content">
              <p className="advanced-hint">
                {' '}{t("Override the default model for specific features. Leave as \"Use default\" to use the settings above.")}{' '}</p>
              
              {(Object.keys(FEATURE_CONFIG) as FeatureType[]).map(feature => {
                const featureConfig = FEATURE_CONFIG[feature];
                const currentModel = getFeatureCurrentModel(feature);
                const currentReasoning = getFeatureCurrentReasoning(feature);
                const isOverridden = currentModel !== 'default';
                const selectedModelConfig = isOverridden ? MODEL_CONFIG[currentModel as ModelType] : null;
                
                return (
                  <div key={feature} className={`feature-override ${isOverridden ? 'active' : ''}`}>
                    <div className="feature-header">
                      <span className="feature-name">{translate(featureConfig.displayName)}</span>
                      <span className="feature-desc">{translate(featureConfig.description)}</span>
                    </div>
                    <div className="feature-controls">
                      <select
                        value={currentModel}
                        onChange={(e) => handleFeatureModelChange(feature, e.target.value)}
                        className="feature-model-select"
                      >
                        <option value="default">{t("Use default")}</option>
                        {availableModels.map(model => (
                          <option key={model} value={model}>
                            {MODEL_CONFIG[model].displayName}
                          </option>
                        ))}
                      </select>
                      
                      {isOverridden && selectedModelConfig?.isReasoning && (
                        <select
                          value={currentReasoning}
                          onChange={(e) => handleFeatureReasoningChange(feature, e.target.value as ReasoningEffort)}
                          className="feature-reasoning-select"
                        >
                          {getSupportedReasoningEfforts(currentModel as ModelType).map(level => (
                            <option key={level} value={level}>
                              {t(getReasoningEffortLabel(level))}
                            </option>
                          ))}
                        </select>
                      )}
                    </div>
                    <div className="feature-recommended">
                      {' '}{t("Recommended:")}{' '}{MODEL_CONFIG[featureConfig.recommendedModel].displayName}
                      {featureConfig.recommendedReasoning && ` (${featureConfig.recommendedReasoning})`}
                    </div>
                  </div>
                );
              })}
              
              {hasAnyOverride && (
                <button className="reset-overrides" onClick={resetAllOverrides}>
                  <RotateCcw size={12} />
                  {' '}{t("Reset all to default")}{' '}</button>
              )}
            </div>
          )}
        </div>
      </div>
    </div>
  );
};

export default ModelSelector;
