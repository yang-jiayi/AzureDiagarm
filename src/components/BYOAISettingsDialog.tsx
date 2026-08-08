// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

import { useEffect, useMemo, useState } from 'react';
import {
  AlertTriangle,
  CheckCircle2,
  ChevronDown,
  CircleDashed,
  Info,
  KeyRound,
  Loader2,
  PlugZap,
  ShieldCheck,
  X,
} from 'lucide-react';
import {
  applyBYOAIAutomaticCapabilities,
  BYOAI_REASONING_EFFORTS,
  DEFAULT_BYO_AI_SETTINGS,
  disableBYOAI,
  getBYOAIApiKey,
  getBYOAISettings,
  isBYOAIVerified,
  saveBYOAIConfiguration,
  validateBYOAISettings,
  type BYOAICapabilityMode,
  type BYOAIProvider,
  type BYOAISettings,
  type BYOAIValidationResult,
} from '../stores/byoAISettingsStore';
import { getReasoningEffortLabel } from '../stores/modelSettingsStore';
import { buildRequestBody, callAzureOpenAIProxy } from '../services/apiHelper';
import { useRuntimeConfig } from '../services/runtimeConfig';
import { useEscapeKey } from '../hooks/useEscapeKey';
import { useModalFocus } from '../hooks/useModalFocus';
import { useLanguage } from '../i18n/LanguageContext';
import { localize } from '../i18n/localization';
import './BYOAISettingsDialog.css';

interface BYOAISettingsDialogProps {
  isOpen: boolean;
  onClose: () => void;
}

type ConnectionStatus =
  | { kind: 'idle'; message: '' }
  | { kind: 'success' | 'error'; message: string };

const OPENAI_ENDPOINT = 'https://api.openai.com';

function configurationSignature(settings: BYOAISettings, apiKey: string): string {
  const { enabled: _enabled, ...configuration } = settings;
  return `${JSON.stringify(configuration)}\u0000${apiKey.trim()}`;
}

export default function BYOAISettingsDialog({
  isOpen,
  onClose,
}: BYOAISettingsDialogProps) {
  const { language } = useLanguage();
  const runtimeConfig = useRuntimeConfig();
  const dialogRef = useModalFocus<HTMLElement>(isOpen);
  const [draft, setDraft] = useState<BYOAISettings>(DEFAULT_BYO_AI_SETTINGS);
  const [apiKey, setApiKey] = useState('');
  const [errors, setErrors] = useState<BYOAIValidationResult['errors']>({});
  const [status, setStatus] = useState<ConnectionStatus>({ kind: 'idle', message: '' });
  const [verifiedSignature, setVerifiedSignature] = useState('');
  const [isTesting, setIsTesting] = useState(false);
  const text = (en: string, ja: string) => localize(language, { en, ja });

  const validation = useMemo(
    () => validateBYOAISettings(draft, apiKey),
    [apiKey, draft],
  );
  const currentSignature = validation.settings
    ? configurationSignature(validation.settings, apiKey)
    : '';
  const serverAvailable = runtimeConfig.status === 'ready'
    && runtimeConfig.bringYourOwnAI;
  const serverChecking = runtimeConfig.status === 'unknown'
    || runtimeConfig.status === 'loading';
  const connectionVerified = Boolean(
    currentSignature && verifiedSignature === currentSignature,
  );
  const canTest = serverAvailable && validation.valid && !isTesting;
  const canSave = canTest && connectionVerified;

  useEffect(() => {
    if (!isOpen) return;
    const settings = getBYOAISettings();
    const key = getBYOAIApiKey();
    setDraft(settings);
    setApiKey(key);
    setErrors({});
    setStatus({ kind: 'idle', message: '' });
    setVerifiedSignature(
      isBYOAIVerified() ? configurationSignature(settings, key) : '',
    );
  }, [isOpen]);

  useEscapeKey(isOpen && !isTesting, onClose);

  if (!isOpen) return null;

  const resetVerification = () => {
    setVerifiedSignature('');
    setStatus({ kind: 'idle', message: '' });
  };

  const updateDraft = <K extends keyof BYOAISettings>(
    key: K,
    value: BYOAISettings[K],
  ) => {
    setDraft(current => {
      const next = { ...current, [key]: value };
      return key === 'model' || key === 'capabilityMode'
        ? applyBYOAIAutomaticCapabilities(next)
        : next;
    });
    setErrors(current => ({ ...current, [key]: undefined }));
    resetVerification();
  };

  const handleProviderChange = (provider: BYOAIProvider) => {
    setDraft(current => applyBYOAIAutomaticCapabilities({
      ...current,
      provider,
      endpoint: provider === 'openai' ? OPENAI_ENDPOINT : '',
    }));
    setErrors({});
    resetVerification();
  };

  const validate = () => {
    const result = validateBYOAISettings(draft, apiKey);
    setErrors(result.errors);
    return result;
  };

  const handleSave = () => {
    const result = validate();
    if (!result.valid || !result.settings) return;
    const signature = configurationSignature(result.settings, apiKey);
    if (!serverAvailable || verifiedSignature !== signature) {
      setStatus({
        kind: 'error',
        message: text(
          'Test this exact connection before saving it.',
          '保存する前に、この接続内容で接続テストを実行してください。',
        ),
      });
      return;
    }
    saveBYOAIConfiguration(
      { ...result.settings, enabled: true },
      apiKey,
      { verified: true },
    );
    onClose();
  };

  const handleTest = async () => {
    const result = validate();
    if (!result.valid || !result.settings || !serverAvailable) return;

    const settings = result.settings;
    setIsTesting(true);
    setStatus({ kind: 'idle', message: '' });
    try {
      const body = buildRequestBody({
        deployment: settings.model,
        messages: [
          {
            role: 'user',
            content: 'Reply with JSON only: {"status":"ok"}',
          },
        ],
        maxTokens: 128,
        apiFormat: settings.apiFormat,
        isReasoning: settings.isReasoning,
        reasoningEffort: settings.isReasoning ? settings.reasoningEffort : 'none',
      });
      const response = await callAzureOpenAIProxy({
        apiFormat: settings.apiFormat,
        deployment: settings.model,
        body,
        byo: {
          provider: settings.provider,
          endpoint: settings.endpoint,
          apiKey: apiKey.trim(),
        },
      });
      if (!response.ok) {
        throw new Error(response.error?.message || text(
          'The endpoint rejected the connection test.',
          'エンドポイントが接続テストを拒否しました。',
        ));
      }
      setVerifiedSignature(configurationSignature(settings, apiKey));
      setStatus({
        kind: 'success',
        message: text(
          'Connection verified. You can now save and use it.',
          '接続を確認しました。保存して使用できます。',
        ),
      });
    } catch (error) {
      setVerifiedSignature('');
      setStatus({
        kind: 'error',
        message: error instanceof Error
          ? error.message
          : text('Connection test failed.', '接続テストに失敗しました。'),
      });
    } finally {
      setIsTesting(false);
    }
  };

  const handleDisconnect = () => {
    disableBYOAI();
    onClose();
  };

  const statusContent = (() => {
    if (status.kind !== 'idle') return status;
    if (serverChecking) {
      return {
        kind: 'info' as const,
        message: text(
          'Checking whether custom AI connections are enabled on this server.',
          'このサーバーでカスタム AI 接続が有効か確認しています。',
        ),
      };
    }
    if (runtimeConfig.status === 'error') {
      return {
        kind: 'error' as const,
        message: text(
          'The application could not confirm server availability. Custom AI remains blocked until the server can be reached.',
          'サーバーの利用可否を確認できませんでした。サーバーへ接続できるまでカスタム AI はブロックされます。',
        ),
      };
    }
    if (!serverAvailable) {
      return {
        kind: 'warning' as const,
        message: text(
          'Custom AI connections are disabled by the application administrator.',
          'カスタム AI 接続はアプリケーション管理者によって無効化されています。',
        ),
      };
    }
    if (draft.enabled && !apiKey) {
      return {
        kind: 'warning' as const,
        message: text(
          'API key required: the saved connection remains selected, but requests stay blocked until you re-enter and verify the key.',
          'API キーが必要です。保存済みの接続は選択されたままですが、キーを再入力して確認するまでリクエストはブロックされます。',
        ),
      };
    }
    if (connectionVerified) {
      return {
        kind: 'success' as const,
        message: text('Connection verified.', '接続確認済みです。'),
      };
    }
    return {
      kind: 'info' as const,
      message: text(
        'Not connected. Complete the fields and run a connection test.',
        '未接続です。項目を入力して接続テストを実行してください。',
      ),
    };
  })();

  const statusLabel = connectionVerified
    ? text('Verified', '確認済み')
    : draft.enabled && !apiKey
      ? text('Key required', 'キーが必要')
      : runtimeConfig.status === 'error'
        ? text('Server unavailable', 'サーバー利用不可')
      : serverAvailable
        ? text('Not connected', '未接続')
        : text('Unavailable', '利用不可');

  return (
    <div className="byo-ai-dialog-overlay" onClick={isTesting ? undefined : onClose}>
      <section
        ref={dialogRef}
        className="byo-ai-dialog"
        role="dialog"
        aria-modal="true"
        aria-labelledby="byo-ai-dialog-title"
        aria-busy={isTesting}
        tabIndex={-1}
        onClick={event => event.stopPropagation()}
      >
        <header className="byo-ai-dialog-header">
          <div className="byo-ai-dialog-heading">
            <span className="byo-ai-dialog-icon" aria-hidden="true">
              <PlugZap size={22} />
            </span>
            <div>
              <span className="byo-ai-dialog-eyebrow">
                {text('Optional AI connection', '任意の AI 接続')}
              </span>
              <h2 id="byo-ai-dialog-title">
                {text('Bring your own AI endpoint', '独自の AI エンドポイントを使用')}
              </h2>
              <span className={`byo-ai-connection-badge byo-ai-connection-badge--${statusContent.kind}`}>
                {connectionVerified
                  ? <CheckCircle2 size={14} aria-hidden="true" />
                  : <CircleDashed size={14} aria-hidden="true" />}
                {statusLabel}
              </span>
            </div>
          </div>
          <button
            type="button"
            className="byo-ai-dialog-close"
            onClick={onClose}
            disabled={isTesting}
            aria-label={text('Close custom AI settings', 'カスタム AI 設定を閉じる')}
          >
            <X size={20} />
          </button>
        </header>

        <div className="byo-ai-dialog-body">
          <section className="byo-ai-section" aria-labelledby="byo-ai-connection-heading">
            <div className="byo-ai-section-heading">
              <div>
                <span className="byo-ai-section-step">1</span>
                <h3 id="byo-ai-connection-heading">
                  {text('Connection', '接続')}
                </h3>
              </div>
              <p>
                {text(
                  'Choose a provider and enter the resource details you control.',
                  '利用するプロバイダーと、ご自身で管理するリソース情報を入力します。',
                )}
              </p>
            </div>

            <div className="byo-ai-form-grid">
              <label>
                <span>{text('Provider', 'プロバイダー')}</span>
                <select
                  value={draft.provider}
                  onChange={event => handleProviderChange(event.target.value as BYOAIProvider)}
                  disabled={isTesting}
                >
                  <option value="azure-openai">Azure OpenAI / Microsoft Foundry</option>
                  <option value="openai">OpenAI</option>
                </select>
              </label>

              <label>
                <span>
                  {draft.provider === 'azure-openai'
                    ? text('Deployment name', 'デプロイ名')
                    : text('Model', 'モデル')}
                </span>
                <input
                  type="text"
                  value={draft.model}
                  onChange={event => updateDraft('model', event.target.value)}
                  placeholder={draft.provider === 'azure-openai' ? 'gpt-5-deployment' : 'gpt-5'}
                  disabled={isTesting}
                  aria-invalid={Boolean(errors.model)}
                  aria-describedby={errors.model ? 'byo-ai-model-error' : undefined}
                />
                {errors.model && (
                  <small id="byo-ai-model-error" className="byo-ai-field-error">
                    {errors.model}
                  </small>
                )}
              </label>

              <label className="byo-ai-field-wide">
                <span>{text('Endpoint', 'エンドポイント')}</span>
                <input
                  type="url"
                  value={draft.endpoint}
                  onChange={event => updateDraft('endpoint', event.target.value)}
                  placeholder="https://your-resource.openai.azure.com"
                  readOnly={draft.provider === 'openai'}
                  disabled={isTesting}
                  aria-invalid={Boolean(errors.endpoint)}
                  aria-describedby={errors.endpoint ? 'byo-ai-endpoint-error' : 'byo-ai-endpoint-help'}
                />
                <small id="byo-ai-endpoint-help" className="byo-ai-field-help">
                  {draft.provider === 'azure-openai'
                    ? text(
                        'Supports Azure OpenAI and Microsoft Foundry resource origins. Do not include /openai/v1.',
                        'Azure OpenAI と Microsoft Foundry のリソース オリジンに対応します。/openai/v1 は含めないでください。',
                      )
                    : text(
                        'Official OpenAI requests are fixed to api.openai.com.',
                        'OpenAI のリクエスト先は api.openai.com に固定されます。',
                      )}
                </small>
                {errors.endpoint && (
                  <small id="byo-ai-endpoint-error" className="byo-ai-field-error">
                    {errors.endpoint}
                  </small>
                )}
              </label>

              <label className="byo-ai-field-wide">
                <span>{text('API key', 'API キー')}</span>
                <div className="byo-ai-secret-field">
                  <KeyRound size={17} aria-hidden="true" />
                  <input
                    type="password"
                    value={apiKey}
                    onChange={event => {
                      setApiKey(event.target.value);
                      setErrors(current => ({ ...current, apiKey: undefined }));
                      resetVerification();
                    }}
                    autoComplete="off"
                    spellCheck={false}
                    placeholder={text(
                      'Kept only for this browser tab',
                      'このブラウザー タブ内でのみ保持',
                    )}
                    disabled={isTesting}
                    aria-invalid={Boolean(errors.apiKey)}
                    aria-describedby={errors.apiKey ? 'byo-ai-key-error' : 'byo-ai-key-help'}
                  />
                </div>
                <small id="byo-ai-key-help" className="byo-ai-field-help">
                  {text(
                    'Closing or reloading this tab clears the key and requires verification again.',
                    'このタブを閉じるか再読み込みするとキーが消去され、再確認が必要になります。',
                  )}
                </small>
                {errors.apiKey && (
                  <small id="byo-ai-key-error" className="byo-ai-field-error">
                    {errors.apiKey}
                  </small>
                )}
              </label>
            </div>
          </section>

          <div className="byo-ai-security-note">
            <ShieldCheck size={19} aria-hidden="true" />
            <p>
              <strong>{text('Credential handling', '資格情報の取り扱い')}</strong>
              <span>
                {text(
                  'Your key is sent through this application server only to reach the selected provider. It stays in this browser tab memory and is not saved to local storage, diagrams, telemetry, URLs, or server logs.',
                  'キーは選択したプロバイダーへ接続するために、このアプリケーション サーバーを経由します。ブラウザー タブのメモリにのみ保持され、ローカル ストレージ、図、テレメトリ、URL、サーバー ログには保存されません。',
                )}
              </span>
            </p>
          </div>

          <details className="byo-ai-advanced">
            <summary>
              <span>
                <span className="byo-ai-section-step">2</span>
                <strong>{text('Advanced model behavior', '高度なモデル動作')}</strong>
              </span>
              <span className="byo-ai-advanced-summary">
                {text('Optional', '任意')}
                <ChevronDown size={16} aria-hidden="true" />
              </span>
            </summary>
            <div className="byo-ai-advanced-body">
              <div className="byo-ai-form-grid">
                <label>
                  <span>{text('API format', 'API 形式')}</span>
                  <select
                    value={draft.apiFormat}
                    onChange={event => updateDraft(
                      'apiFormat',
                      event.target.value as BYOAISettings['apiFormat'],
                    )}
                    disabled={isTesting}
                  >
                    <option value="responses">Responses API (recommended)</option>
                    <option value="chat-completions">Chat Completions</option>
                  </select>
                  <small className="byo-ai-field-help">
                    {text(
                      'Azure v1 routing is used automatically; no dated API version is required.',
                      'Azure v1 ルーティングを自動使用するため、日付付き API バージョンは不要です。',
                    )}
                  </small>
                </label>

                <label>
                  <span>{text('Capability setup', '機能設定')}</span>
                  <select
                    value={draft.capabilityMode}
                    onChange={event => updateDraft(
                      'capabilityMode',
                      event.target.value as BYOAICapabilityMode,
                    )}
                    disabled={isTesting}
                  >
                    <option value="auto">
                      {text('Automatic from model name', 'モデル名から自動判定')}
                    </option>
                    <option value="manual">
                      {text('Manual override', '手動設定')}
                    </option>
                  </select>
                  <small className="byo-ai-field-help">
                    {text(
                      'Use manual override when an Azure deployment uses a custom alias.',
                      'Azure デプロイが独自の別名を使用する場合は手動設定を選択してください。',
                    )}
                  </small>
                </label>

                {draft.isReasoning && (
                  <label>
                    <span>{text('Reasoning effort', '推論強度')}</span>
                    <select
                      value={draft.reasoningEffort}
                      onChange={event => updateDraft(
                        'reasoningEffort',
                        event.target.value as BYOAISettings['reasoningEffort'],
                      )}
                      disabled={isTesting}
                    >
                      {BYOAI_REASONING_EFFORTS.map(effort => (
                        <option key={effort} value={effort}>
                          {text(
                            getReasoningEffortLabel(effort),
                            getReasoningEffortLabel(effort),
                          )}
                        </option>
                      ))}
                    </select>
                    <small className="byo-ai-field-help">
                      {text(
                        'The same effort is used for connection testing and real requests.',
                        '接続テストと実際のリクエストで同じ推論強度を使用します。',
                      )}
                    </small>
                  </label>
                )}
              </div>

              {draft.capabilityMode === 'auto' ? (
                <div className="byo-ai-detected-capabilities">
                  <Info size={17} aria-hidden="true" />
                  <span>
                    <strong>{text('Detected capabilities', '検出された機能')}</strong>
                    <small>
                      {draft.isReasoning
                        ? text('Reasoning enabled', '推論を有効化')
                        : text('Standard model', '標準モデル')}
                      {' · '}
                      {draft.supportsVision
                        ? text('Image input enabled', '画像入力を有効化')
                        : text('Text input only', 'テキスト入力のみ')}
                    </small>
                  </span>
                </div>
              ) : (
                <fieldset className="byo-ai-capabilities" disabled={isTesting}>
                  <legend>{text('Manual capabilities', '手動機能設定')}</legend>
                  <label>
                    <input
                      type="checkbox"
                      checked={draft.isReasoning}
                      onChange={event => updateDraft('isReasoning', event.target.checked)}
                    />
                    <span>
                      <strong>{text('Reasoning model', '推論モデル')}</strong>
                      <small>{text(
                        'Send reasoning effort and reasoning-compatible token limits.',
                        '推論強度と推論モデル対応のトークン上限を送信します。',
                      )}</small>
                    </span>
                  </label>
                  <label>
                    <input
                      type="checkbox"
                      checked={draft.supportsVision}
                      onChange={event => updateDraft('supportsVision', event.target.checked)}
                    />
                    <span>
                      <strong>{text('Image input', '画像入力')}</strong>
                      <small>{text(
                        'Allow architecture image analysis with this model.',
                        'このモデルでアーキテクチャ画像の分析を許可します。',
                      )}</small>
                    </span>
                  </label>
                </fieldset>
              )}
            </div>
          </details>

          <section className="byo-ai-test-section" aria-labelledby="byo-ai-test-heading">
            <div className="byo-ai-section-heading">
              <div>
                <span className="byo-ai-section-step">3</span>
                <h3 id="byo-ai-test-heading">{text('Verify and save', '確認して保存')}</h3>
              </div>
              <p>
                {text(
                  'The test sends a small JSON prompt and may consume provider tokens or incur a small charge.',
                  'テストでは小さな JSON プロンプトを送信するため、プロバイダーのトークン消費または少額の料金が発生する場合があります。',
                )}
              </p>
            </div>

            <div
              className={`byo-ai-status byo-ai-status--${statusContent.kind}`}
              role={statusContent.kind === 'error' ? 'alert' : 'status'}
            >
              {statusContent.kind === 'success' && <CheckCircle2 size={17} aria-hidden="true" />}
              {statusContent.kind === 'warning' && <AlertTriangle size={17} aria-hidden="true" />}
              {statusContent.kind === 'info' && <Info size={17} aria-hidden="true" />}
              {statusContent.kind === 'error' && <AlertTriangle size={17} aria-hidden="true" />}
              <span>{statusContent.message}</span>
            </div>
          </section>
        </div>

        <footer className="byo-ai-dialog-footer">
          <button
            type="button"
            className="btn btn-secondary byo-ai-disconnect"
            onClick={handleDisconnect}
            disabled={isTesting || (!draft.enabled && !getBYOAIApiKey())}
          >
            {text('Disconnect', '接続解除')}
          </button>
          <span className="byo-ai-dialog-footer-spacer" />
          <button
            type="button"
            className="btn btn-secondary"
            onClick={() => void handleTest()}
            disabled={!canTest}
          >
            {isTesting ? <Loader2 size={16} className="spin" /> : <PlugZap size={16} />}
            {serverChecking
              ? text('Checking server…', 'サーバー確認中…')
              : text('Test connection', '接続テスト')}
          </button>
          <button
            type="button"
            className="btn btn-primary"
            onClick={handleSave}
            disabled={!canSave}
            title={!canSave
              ? text(
                  'Complete and verify the connection before saving.',
                  '接続内容を入力し、確認してから保存してください。',
                )
              : undefined}
          >
            <CheckCircle2 size={16} aria-hidden="true" />
            {text('Save verified connection', '確認済み接続を保存')}
          </button>
        </footer>
      </section>
    </div>
  );
}
