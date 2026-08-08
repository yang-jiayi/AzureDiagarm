// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

/**
 * Model Settings Popover
 * Toolbar dropdown for AI model selection across all features.
 * Follows the same toolbar-dropdown pattern as Layout/Export menus.
 */

import { forwardRef } from 'react';
import {
  AlertTriangle,
  Brain,
  CheckCircle2,
  ChevronDown,
  Cpu,
  Globe2,
  KeyRound,
  Layers,
  Moon,
  PlugZap,
  RotateCcw,
  SlidersHorizontal,
  Sparkles,
  Sun,
  X,
  Zap,
} from 'lucide-react';
import {
  useModelSettings,
  MODEL_CONFIG,
  ModelType,
  ReasoningEffort,
  FeatureType,
  FEATURE_CONFIG,
  getAvailableModels,
  getReasoningEffortLabel,
  getRecommendedModelSettings,
  getSupportedReasoningEfforts,
  updateFeatureOverride,
  hasFeatureOverride,
} from '../stores/modelSettingsStore';
import {
  getBYOAIModelLabel,
  useBYOAISettings,
} from '../stores/byoAISettingsStore';
import { useRuntimeConfig } from '../services/runtimeConfig';
import './ModelSettingsPopover.css';
import { useLanguage } from '../i18n/LanguageContext';
import { localize } from '../i18n/localization';

interface ModelSettingsPopoverProps {
  isOpen: boolean;
  onToggle: () => void;
  onOpenBYOSettings: () => void;
}

const ModelSettingsPopover = forwardRef<HTMLDivElement, ModelSettingsPopoverProps>(
  ({ isOpen, onToggle, onOpenBYOSettings }, ref) => {
    const { t, translate, language } = useLanguage();
    const text = (en: string, ja: string) => localize(language, { en, ja });
    const [settings, updateSettings] = useModelSettings();
    const byoSnapshot = useBYOAISettings();
    const runtimeConfig = useRuntimeConfig();
    const availableModels = getAvailableModels();
    const currentConfig = MODEL_CONFIG[settings.model];
    const byoConfigured = byoSnapshot.settings.enabled;
    const byoServerAvailable = runtimeConfig.status === 'ready'
      && runtimeConfig.bringYourOwnAI;
    const byoActive = byoConfigured
      && byoServerAvailable
      && byoSnapshot.verified;
    const byoStatus = !byoConfigured
      ? 'not-connected'
      : runtimeConfig.status === 'unknown' || runtimeConfig.status === 'loading'
        ? 'checking'
        : runtimeConfig.status === 'error'
          ? 'server-error'
        : !byoServerAvailable
          ? 'unavailable'
          : byoSnapshot.connectionState;
    const byoStatusLabel = (() => {
      switch (byoStatus) {
        case 'verified':
          return text('Verified', '確認済み');
        case 'key-required':
          return text('Key required', 'キーが必要');
        case 'unverified':
          return text('Verification required', '接続確認が必要');
        case 'unavailable':
          return text('Disabled by admin', '管理者により無効');
        case 'server-error':
          return text('Server unavailable', 'サーバー利用不可');
        case 'checking':
          return text('Checking server', 'サーバー確認中');
        default:
          return text('Not connected', '未接続');
      }
    })();

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

    const applyRecommendedPortfolio = () => {
      updateSettings(getRecommendedModelSettings());
    };

    const getModelIcon = (model: ModelType) => {
      switch (model) {
        case 'gpt-5.1':
          return <Cpu size={14} />;
        case 'gpt-5.2':
          return <Brain size={14} />;
        case 'gpt-5.6-sol':
          return <Sun size={14} />;
        case 'gpt-5.6-terra':
          return <Globe2 size={14} />;
        case 'gpt-5.6-luna':
          return <Moon size={14} />;
        case 'claude-opus-5':
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
          {byoConfigured ? <PlugZap size={14} /> : getModelIcon(settings.model)}
          <span className="model-popover-label">
            {byoConfigured ? `Custom: ${byoSnapshot.settings.model}` : currentConfig.displayName}
          </span>
          {byoConfigured ? (
            <span className={`model-popover-reasoning model-popover-reasoning--${byoStatus}`}>
              {byoStatusLabel}
            </span>
          ) : currentConfig.isReasoning && (
            <span className="model-popover-reasoning">{t(getReasoningEffortLabel(settings.reasoningEffort))}</span>
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
            <div className="toolbar-dropdown-heading">
              <span>{t("AI model settings")}</span>
              <button
                className="msp-close-btn"
                onClick={onToggle}
                title={t("Close")}
                aria-label={t("Close")}
              >
                <X size={13} />
              </button>
            </div>
            <div className={`msp-byo-card msp-byo-card--${byoStatus}`}>
              <div className="msp-byo-copy">
                <span className="msp-byo-icon" aria-hidden="true">
                  {byoStatus === 'verified'
                    ? <CheckCircle2 size={16} />
                    : byoStatus === 'key-required'
                      ? <KeyRound size={16} />
                      : byoStatus === 'unavailable' || byoStatus === 'server-error'
                        ? <AlertTriangle size={16} />
                        : <PlugZap size={16} />}
                </span>
                <span>
                  <strong>{byoConfigured ? getBYOAIModelLabel() : text(
                    'Bring your own AI endpoint',
                    '独自の AI エンドポイントを使用',
                  )}</strong>
                  <small>
                    {byoActive
                      ? text(
                          `Custom endpoint active with ${getReasoningEffortLabel(byoSnapshot.settings.reasoningEffort).toLowerCase()} reasoning.`,
                          `カスタム エンドポイントを使用中です。推論強度: ${getReasoningEffortLabel(byoSnapshot.settings.reasoningEffort)}`,
                        )
                      : byoStatus === 'key-required'
                        ? text(
                            'The saved endpoint remains selected. Re-enter and verify the API key to resume.',
                            '保存済みエンドポイントは選択されたままです。API キーを再入力して確認すると再開できます。',
                          )
                        : byoStatus === 'unavailable'
                          ? text(
                              'The server kill switch is off. Disconnect or contact the administrator.',
                              'サーバーのキル スイッチがオフです。接続解除するか管理者に連絡してください。',
                            )
                          : byoStatus === 'server-error'
                            ? text(
                                'Server availability could not be confirmed. Requests remain blocked.',
                                'サーバーの利用可否を確認できないため、リクエストはブロックされています。',
                              )
                            : byoConfigured
                            ? text(
                                'Test this connection before it can be used.',
                                '使用する前に接続テストを実行してください。',
                              )
                            : text(
                                'Use your Azure OpenAI, Microsoft Foundry, or official OpenAI endpoint.',
                                '独自の Azure OpenAI、Microsoft Foundry、または公式 OpenAI エンドポイントを使用します。',
                              )}
                  </small>
                  <span className={`msp-byo-status msp-byo-status--${byoStatus}`}>
                    {byoStatusLabel}
                  </span>
                </span>
              </div>
              <button
                type="button"
                className="msp-byo-button"
                onClick={() => {
                  onToggle();
                  onOpenBYOSettings();
                }}
              >
                {byoStatus === 'key-required'
                  ? text('Enter key', 'キーを入力')
                  : byoConfigured
                    ? text('Configure', '設定')
                    : text('Connect', '接続')}
              </button>
            </div>

            <div className="toolbar-dropdown-separator" role="separator" />
            <div className="toolbar-dropdown-heading">
              <span>{text('Managed models', '管理モデル')}</span>
            </div>
            {byoConfigured && (
              <p className="msp-managed-note">
                {text(
                  'Managed settings remain saved but are paused while a custom endpoint is selected, including while its key is missing.',
                  '管理モデル設定は保存されたままですが、キー未入力時を含め、カスタム エンドポイントが選択されている間は一時停止します。',
                )}
              </p>
            )}
            <div className="msp-model-buttons">
              {availableModels.map((model) => (
                <button
                  key={model}
                  className={`msp-model-btn ${settings.model === model ? 'active' : ''}`}
                  onClick={() => handleModelChange(model)}
                  title={translate(MODEL_CONFIG[model].description)}
                  disabled={byoConfigured}
                >
                  <span className="msp-model-btn-main">
                    {getModelIcon(model)}
                    <span>{MODEL_CONFIG[model].displayName}</span>
                  </span>
                  {MODEL_CONFIG[model].recommendedUse && (
                    <span className="msp-model-role">
                      {translate(MODEL_CONFIG[model].recommendedUse)}
                    </span>
                  )}
                </button>
              ))}
            </div>

            {/* Reasoning effort (only shown when reasoning model selected) */}
            {currentConfig.isReasoning && (
              <>
                <div className="msp-reasoning-row">
                  <span className="msp-reasoning-label">{t("Reasoning")}</span>
                  <div className="msp-reasoning-buttons">
                    {getSupportedReasoningEfforts(settings.model).map((level) => (
                      <button
                        key={level}
                        className={`msp-reasoning-btn ${settings.reasoningEffort === level ? 'active' : ''}`}
                        onClick={() => handleReasoningChange(level)}
                        title={level === 'none' ? t("No reasoning - fastest response") : undefined}
                        disabled={byoConfigured}
                      >
                        {t(getReasoningEffortLabel(level))}
                      </button>
                    ))}
                  </div>
                </div>
              </>
            )}

            <details className="msp-advanced-settings">
              <summary>
                <span>
                  <SlidersHorizontal size={14} aria-hidden="true" />
                  {text('Advanced feature overrides', '高度な機能別設定')}
                </span>
                <span className="msp-advanced-summary-status">
                  {hasAnyOverride
                    ? text('Customized', 'カスタム済み')
                    : text('Using defaults', '既定値を使用')}
                  <ChevronDown size={14} aria-hidden="true" />
                </span>
              </summary>
              <div className="msp-advanced-body">
                <div className="msp-portfolio">
                  <div className="msp-portfolio-copy">
                    <strong>{t("Recommended portfolio")}</strong>
                    <span>{t("Sol for architecture • Terra for validation and deployment • Luna for blueprints")}</span>
                  </div>
                  <button
                    className="msp-portfolio-btn"
                    onClick={applyRecommendedPortfolio}
                    title={t("Use recommended portfolio")}
                    disabled={byoConfigured}
                  >
                    <Sparkles size={12} />
                    {t("Apply")}
                  </button>
                </div>

                <div className="toolbar-dropdown-heading">
                  {' '}{t("Per-Feature Settings")}{' '}{hasAnyOverride && (
                    <button
                      className="msp-reset-btn"
                      onClick={resetAllOverrides}
                      title={t("Reset all to default")}
                      disabled={byoConfigured}
                    >
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
                          <span className="msp-feature-name">{translate(featureConfig.displayName)}</span>
                          <span className="msp-feature-effective">
                            {MODEL_CONFIG[effectiveModel].displayName}
                            {effectiveReasoning && ` (${t(getReasoningEffortLabel(effectiveReasoning))})`}
                          </span>
                        </div>
                        <div className="msp-feature-controls">
                          <select
                            value={currentModel}
                            onChange={(e) => handleFeatureModelChange(feature, e.target.value)}
                            className="msp-feature-select"
                            disabled={byoConfigured}
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
                              disabled={byoConfigured}
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
                </div>
              </div>
            </details>
          </div>
        )}
      </div>
    );
  }
);

ModelSettingsPopover.displayName = 'ModelSettingsPopover';

export default ModelSettingsPopover;
