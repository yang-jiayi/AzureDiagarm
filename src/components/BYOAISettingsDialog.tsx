// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

import { useEffect, useState } from 'react';
import { CheckCircle2, KeyRound, Loader2, PlugZap, ShieldCheck, X } from 'lucide-react';
import {
  DEFAULT_BYO_AI_SETTINGS,
  disableBYOAI,
  getBYOAIApiKey,
  getBYOAISettings,
  saveBYOAIConfiguration,
  validateBYOAISettings,
  type BYOAIProvider,
  type BYOAISettings,
  type BYOAIValidationResult,
} from '../stores/byoAISettingsStore';
import { buildRequestBody, callAzureOpenAIProxy } from '../services/apiHelper';
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

export default function BYOAISettingsDialog({
  isOpen,
  onClose,
}: BYOAISettingsDialogProps) {
  const { language } = useLanguage();
  const dialogRef = useModalFocus<HTMLElement>(isOpen);
  const [draft, setDraft] = useState<BYOAISettings>(DEFAULT_BYO_AI_SETTINGS);
  const [apiKey, setApiKey] = useState('');
  const [errors, setErrors] = useState<BYOAIValidationResult['errors']>({});
  const [status, setStatus] = useState<ConnectionStatus>({ kind: 'idle', message: '' });
  const [isTesting, setIsTesting] = useState(false);
  const text = (en: string, ja: string) => localize(language, { en, ja });

  useEffect(() => {
    if (!isOpen) return;
    setDraft(getBYOAISettings());
    setApiKey(getBYOAIApiKey());
    setErrors({});
    setStatus({ kind: 'idle', message: '' });
  }, [isOpen]);

  useEscapeKey(isOpen && !isTesting, onClose);

  if (!isOpen) return null;

  const updateDraft = <K extends keyof BYOAISettings>(
    key: K,
    value: BYOAISettings[K],
  ) => {
    setDraft(current => ({ ...current, [key]: value }));
    setErrors(current => ({ ...current, [key]: undefined }));
    setStatus({ kind: 'idle', message: '' });
  };

  const handleProviderChange = (provider: BYOAIProvider) => {
    setDraft(current => ({
      ...current,
      provider,
      endpoint: provider === 'openai' ? OPENAI_ENDPOINT : '',
      apiVersion: provider === 'openai'
        ? ''
        : (current.apiVersion || DEFAULT_BYO_AI_SETTINGS.apiVersion),
    }));
    setErrors({});
    setStatus({ kind: 'idle', message: '' });
  };

  const validate = () => {
    const result = validateBYOAISettings(draft, apiKey);
    setErrors(result.errors);
    return result;
  };

  const handleSave = () => {
    const result = validate();
    if (!result.valid || !result.settings) return;
    saveBYOAIConfiguration({ ...result.settings, enabled: true }, apiKey);
    onClose();
  };

  const handleTest = async () => {
    const result = validate();
    if (!result.valid || !result.settings) return;

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
        maxTokens: 64,
        apiFormat: settings.apiFormat,
        isReasoning: settings.isReasoning,
        reasoningEffort: settings.isReasoning ? 'low' : 'none',
      });
      const response = await callAzureOpenAIProxy({
        apiFormat: settings.apiFormat,
        deployment: settings.model,
        body,
        byo: {
          provider: settings.provider,
          endpoint: settings.endpoint,
          apiKey: apiKey.trim(),
          ...(settings.provider === 'azure-openai'
            ? { apiVersion: settings.apiVersion }
            : {}),
        },
      });
      if (!response.ok) {
        throw new Error(response.error?.message || text(
          'The endpoint rejected the connection test.',
          'エンドポイントが接続テストを拒否しました。',
        ));
      }
      setStatus({
        kind: 'success',
        message: text('Connection successful.', '接続に成功しました。'),
      });
    } catch (error) {
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

          <div className="byo-ai-form-grid">
            <label>
              <span>{text('Provider', 'プロバイダー')}</span>
              <select
                value={draft.provider}
                onChange={event => handleProviderChange(event.target.value as BYOAIProvider)}
                disabled={isTesting}
              >
                <option value="azure-openai">Azure OpenAI</option>
                <option value="openai">OpenAI</option>
              </select>
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
                aria-describedby={errors.endpoint ? 'byo-ai-endpoint-error' : undefined}
              />
              {errors.endpoint && (
                <small id="byo-ai-endpoint-error" className="byo-ai-field-error">
                  {errors.endpoint}
                </small>
              )}
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
                placeholder={draft.provider === 'azure-openai' ? 'gpt-5.6-sol' : 'gpt-5'}
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
                <option value="responses">Responses API</option>
                <option value="chat-completions">Chat Completions</option>
              </select>
            </label>

            {draft.provider === 'azure-openai' && (
              <label>
                <span>{text('API version', 'API バージョン')}</span>
                <input
                  type="text"
                  value={draft.apiVersion}
                  onChange={event => updateDraft('apiVersion', event.target.value)}
                  placeholder="2024-05-01-preview"
                  disabled={isTesting}
                  aria-invalid={Boolean(errors.apiVersion)}
                  aria-describedby={errors.apiVersion ? 'byo-ai-version-error' : undefined}
                />
                {errors.apiVersion && (
                  <small id="byo-ai-version-error" className="byo-ai-field-error">
                    {errors.apiVersion}
                  </small>
                )}
              </label>
            )}

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
                    setStatus({ kind: 'idle', message: '' });
                  }}
                  autoComplete="off"
                  spellCheck={false}
                  placeholder={text('Kept only for this browser tab', 'このブラウザー タブ内でのみ保持')}
                  disabled={isTesting}
                  aria-invalid={Boolean(errors.apiKey)}
                  aria-describedby={errors.apiKey ? 'byo-ai-key-error' : 'byo-ai-key-help'}
                />
              </div>
              <small id="byo-ai-key-help" className="byo-ai-field-help">
                {text(
                  'Closing or reloading this tab clears the key.',
                  'このタブを閉じるか再読み込みすると、キーは消去されます。',
                )}
              </small>
              {errors.apiKey && (
                <small id="byo-ai-key-error" className="byo-ai-field-error">
                  {errors.apiKey}
                </small>
              )}
            </label>
          </div>

          <fieldset className="byo-ai-capabilities" disabled={isTesting}>
            <legend>{text('Model capabilities', 'モデル機能')}</legend>
            <label>
              <input
                type="checkbox"
                checked={draft.isReasoning}
                onChange={event => updateDraft('isReasoning', event.target.checked)}
              />
              <span>
                <strong>{text('Reasoning model', '推論モデル')}</strong>
                <small>{text(
                  'Send the selected reasoning effort when supported.',
                  '対応している場合、選択された推論強度を送信します。',
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

          {status.kind !== 'idle' && (
            <div className={`byo-ai-status byo-ai-status--${status.kind}`} role="status">
              {status.kind === 'success' && <CheckCircle2 size={17} aria-hidden="true" />}
              <span>{status.message}</span>
            </div>
          )}
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
            disabled={isTesting}
          >
            {isTesting ? <Loader2 size={16} className="spin" /> : <PlugZap size={16} />}
            {text('Test connection', '接続テスト')}
          </button>
          <button
            type="button"
            className="btn btn-primary"
            onClick={handleSave}
            disabled={isTesting}
          >
            {text('Save and use', '保存して使用')}
          </button>
        </footer>
      </section>
    </div>
  );
}
