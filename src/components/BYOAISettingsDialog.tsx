// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

import { useEffect, useRef, useState } from 'react';
import { CheckCircle2, Eye, EyeOff, KeyRound, Loader2, PlugZap, Plus, Trash2, X } from 'lucide-react';
import {
  getBYOAIConnectionState, getBYOAISettings, removeBYOAIProfile, selectBYOAIProfile,
  setBYOAIApiKey, upsertBYOAIProfile, useBYOAISettings, invalidateBYOAIProfile,
  BYOAI_REASONING_EFFORTS, MAX_BYO_AI_PROFILES,
  type BYOAIProfile,
} from '../stores/byoAISettingsStore';
import { testBYOAIConnection } from '../services/byoAIConnection';
import { loadRuntimeConfig, useRuntimeConfig } from '../services/runtimeConfig';
import { getReasoningEffortLabel } from '../stores/modelSettingsStore';
import { useLanguage } from '../i18n/LanguageContext';
import { localize } from '../i18n/localization';
import AIConnectionSelector, { connectionStateLabel } from './AIConnectionSelector';
import AstraReasoningSettings from './AstraReasoningSettings';
import ModalScaffold from './ModalScaffold';
import './BYOAISettingsDialog.css';

interface BYOAISettingsDialogProps {
  isOpen: boolean;
  onClose: () => void;
  returnFocusTarget?: HTMLElement | null;
}

function emptyProfile(): BYOAIProfile {
  return {
    id: crypto.randomUUID(), name: '', provider: 'azure-openai', endpoint: '', model: '',
    apiFormat: 'responses', reasoningEffort: 'none', isReasoning: false,
    supportsVision: false, maxCompletionTokens: 8000,
  };
}

export default function BYOAISettingsDialog({
  isOpen, onClose, returnFocusTarget,
}: BYOAISettingsDialogProps) {
  const { language, t, translate } = useLanguage();
  const { storageError } = useBYOAISettings();
  const settings = getBYOAISettings();
  const policy = useRuntimeConfig();
  const [draft, setDraft] = useState<BYOAIProfile>(emptyProfile);
  const [dirty, setDirty] = useState(false);
  const [error, setError] = useState('');
  const [notice, setNotice] = useState('');
  const [confirmDelete, setConfirmDelete] = useState(false);
  const [showKey, setShowKey] = useState(false);
  const [keyValue, setKeyValue] = useState('');
  const testRef = useRef<AbortController | null>(null);
  const [testing, setTesting] = useState(false);
  const nameRef = useRef<HTMLInputElement>(null);
  const text = (en: string, ja: string) => localize(language, { en, ja });
  const saved = settings.profiles.find(profile => profile.id === draft.id);
  const connection = saved ? getBYOAIConnectionState(saved.id) : null;
  const state = connection?.status ?? 'key-required';
  const active = saved?.id === settings.activeProfileId;
  const serverAllowed = policy.status === 'ready' && policy.bringYourOwnAI;
  const limitReached = settings.profiles.length >= MAX_BYO_AI_PROFILES;
  const displayState = storageError ? 'storage-unavailable' : policy.status === 'error' ? 'policy-unavailable'
    : policy.status !== 'ready' ? 'policy-checking' : !policy.bringYourOwnAI ? 'admin-disabled' : state;

  const cancelTest = () => {
    const controller = testRef.current;
    testRef.current = null;
    controller?.abort();
    setTesting(false);
    if (controller) setNotice(text('Connection test cancelled. Test again before activation.', '接続テストをキャンセルしました。有効化する前に再テストしてください。'));
  };
  useEffect(() => {
    if (!isOpen) return;
    const snapshot = getBYOAISettings();
    setDraft(snapshot.profiles.find(profile => profile.id === snapshot.activeProfileId)
      ?? snapshot.profiles[0] ?? emptyProfile());
    setDirty(false);
    setError('');
    setNotice('');
    setConfirmDelete(false);
    setShowKey(false);
    setKeyValue('');
    return () => {
      testRef.current?.abort();
      testRef.current = null;
    };
  }, [isOpen]);

  const chooseDraft = (profile?: BYOAIProfile) => {
    cancelTest();
    setDraft(profile ? { ...profile } : emptyProfile());
    setDirty(false);
    setError('');
    setNotice('');
    setKeyValue('');
    setShowKey(false);
    setConfirmDelete(false);
    window.requestAnimationFrame(() => nameRef.current?.focus());
  };
  const updateDraft = <K extends keyof BYOAIProfile>(key: K, value: BYOAIProfile[K]) => {
    if (key !== 'name') cancelTest();
    setShowKey(false);
    setDraft(current => ({ ...current, [key]: value }));
    setDirty(true);
    setError('');
    setNotice('');
    setConfirmDelete(false);
    if (saved && key !== 'name') {
      try { invalidateBYOAIProfile(saved.id); }
      catch (cause) { safeFailure(cause); }
    }
  };
  const safeFailure = (cause: unknown) => setError(translate(
    cause instanceof Error ? cause.message : 'The AI connection is unavailable.',
  ));
  const save = () => {
    cancelTest();
    try {
      upsertBYOAIProfile(draft);
      setDraft(getBYOAISettings().profiles.find(profile => profile.id === draft.id)!);
      setDirty(false);
      setError('');
      setNotice(text(
        'Public profile saved. Saving does not activate it. Connection changes require a successful test before you choose Use this profile.',
        '公開設定を保存しました。保存だけでは有効化されません。接続内容を変更した場合は、テストを成功させてから「このプロファイルを使用」を選択してください。',
      ));
    } catch (cause) { safeFailure(cause); }
  };
  const testConnection = async () => {
    if (!saved || dirty || testing || storageError || !serverAllowed) return;
    const controller = new AbortController();
    testRef.current = controller;
    setTesting(true);
    setError('');
    setNotice('');
    try {
      await testBYOAIConnection(saved.id, { signal: controller.signal });
    } catch (cause) {
      if (testRef.current === controller && !controller.signal.aborted
        && !getBYOAIConnectionState(saved.id).error) safeFailure(cause);
    } finally {
      if (testRef.current === controller) {
        testRef.current = null;
        setTesting(false);
      }
    }
  };
  const activate = () => {
    if (!saved || dirty || testing || storageError || state !== 'verified' || !serverAllowed) return;
    try {
      selectBYOAIProfile(saved.id);
      setError('');
      setNotice(text('This profile is now selected for all AI features.', 'このプロファイルをすべての AI 機能に選択しました。'));
    } catch (cause) { safeFailure(cause); }
  };
  const deleteProfile = () => {
    if (!saved || active || testing || getBYOAISettings().activeProfileId === saved.id) return;
    try {
      removeBYOAIProfile(saved.id);
      chooseDraft(getBYOAISettings().profiles[0]);
    } catch (cause) { safeFailure(cause); }
  };

  return (
    <ModalScaffold isOpen={isOpen} onClose={onClose} returnFocusTarget={returnFocusTarget}
      className="byo-ai-dialog" overlayClassName="byo-ai-dialog-overlay"
      ariaLabelledBy="byo-ai-dialog-title" closeOnBackdrop={!testing}
      closeOnEscape={!testing} aria-busy={testing}>
      <header className="byo-ai-dialog-header">
        <div><PlugZap size={23} aria-hidden="true" /><h2 id="byo-ai-dialog-title">{text('AI connections', 'AI 接続')}</h2></div>
        <button type="button" className="byo-ai-icon-button" disabled={testing} onClick={onClose}
          aria-label={text('Close AI connections', 'AI 接続を閉じる')}><X size={20} /></button>
      </header>
      <div className="byo-ai-dialog-body">
        <p className="byo-ai-privacy-note"><KeyRound size={17} aria-hidden="true" />{text(
          'Only public profiles are saved on this device. Keys and verification stay in this tab’s memory, never browser storage. Re-enter and test keys after reload. A connection test sends a small request to your selected provider and may incur charges.',
          'このデバイスに保存されるのは公開プロファイルのみです。キーと確認状態はタブのメモリ内にのみ保持され、ブラウザーのストレージには保存されません。再読み込み後はキーを再入力してテストしてください。接続テストでは選択したプロバイダーに少量のリクエストを送信するため、料金が発生する場合があります。',
        )}</p>
        <AIConnectionSelector disabled={testing} />
        {!serverAllowed && <div className="byo-ai-policy-note">
          <p role="status">{policy.status === 'ready'
          ? text(
            'BYO testing and activation are disabled by the application administrator. You can still add, edit, and delete inactive profiles. Managed GPT-6 Astra remains available if configured.',
            '管理者が BYO のテストと有効化を無効にしています。プロファイルの追加・編集・非アクティブなプロファイルの削除は可能です。設定済みの場合は管理対象の GPT-6 Astra を使用できます。',
          )
          : text(
            'Server policy is not yet available. Profile editing is available, but testing and activation stay blocked until the server confirms permission.',
            'サーバー ポリシーを確認できていません。プロファイルの編集は可能ですが、サーバーが許可を確認するまでテストと有効化はブロックされます。',
          )}</p>
          <button type="button" onClick={() => void loadRuntimeConfig(true)}
            disabled={testing || policy.status === 'loading'}>
            {text('Refresh server policy', 'サーバー ポリシーを再確認')}
          </button>
        </div>}
        <section className="byo-ai-profiles" aria-label={text('Saved BYO profiles', '保存済み BYO プロファイル')}>
          <div className="byo-ai-section-heading">
            <h3>{text('Your connection profiles', '接続プロファイル')} <span>({settings.profiles.length}/{MAX_BYO_AI_PROFILES})</span></h3>
            <button type="button" onClick={() => chooseDraft()} disabled={limitReached || testing}>
              <Plus size={16} aria-hidden="true" />{text('Add profile', 'プロファイルを追加')}
            </button>
          </div>
          {limitReached && <p>{text('Maximum 10 profiles. Delete an inactive profile to add another.', 'プロファイルは最大 10 件です。追加するには非アクティブなプロファイルを削除してください。')}</p>}
          <div className="byo-ai-profile-list">
            {settings.profiles.map(profile => (
              <button type="button" key={profile.id} aria-pressed={profile.id === draft.id}
                onClick={() => chooseDraft(profile)} disabled={testing}>
                <strong>{profile.name}</strong>
                <span>{profile.provider === 'openai' ? 'OpenAI' : 'Azure OpenAI'} · {profile.model}</span>
                <span>{connectionStateLabel(getBYOAIConnectionState(profile.id).status, language)}
                  {settings.activeProfileId === profile.id && <> · {text('Selected', '選択中')}</>}</span>
              </button>
            ))}
          </div>
        </section>
        <form className="byo-ai-editor" onSubmit={event => { event.preventDefault(); save(); }}>
          <h3>{saved ? text('Edit profile', 'プロファイルを編集') : text('New profile', '新しいプロファイル')}</h3>
          <p>{text(
            'Profile fields are public: never paste keys or credentials here. Editing connection details, capabilities, or keys invalidates verification immediately; renaming only does not. Save edits before testing. The selected connection is not automatically switched.',
            'プロファイル項目は公開設定です。キーや資格情報を貼り付けないでください。接続内容・機能・キーを編集すると確認状態は直ちに無効になりますが、名前だけの変更では無効になりません。編集内容を保存してからテストしてください。選択中の接続は自動で切り替わりません。',
          )}</p>
          <div className="byo-ai-form-grid">
            <label><span>{text('Profile name', 'プロファイル名')}</span>
              <input ref={nameRef} value={draft.name} maxLength={80} required autoComplete="off"
                onChange={event => updateDraft('name', event.target.value)} /></label>
            <label><span id="byo-provider-label">{text('Provider', 'プロバイダー')}</span>
              <select aria-labelledby="byo-provider-label" value={draft.provider} onChange={event => {
                const provider = event.target.value as BYOAIProfile['provider'];
                updateDraft('provider', provider);
                setDraft(current => ({ ...current, endpoint: provider === 'openai' ? 'https://api.openai.com' : '' }));
              }}>
                <option value="azure-openai">Azure OpenAI / Microsoft Foundry</option>
                <option value="openai">{text('Official OpenAI', '公式 OpenAI')}</option>
              </select></label>
            <label className="byo-ai-full-row"><span id="byo-endpoint-label">{text('Endpoint origin', 'エンドポイントのオリジン')}</span>
              <input value={draft.endpoint} type="url" required autoComplete="off" spellCheck={false}
                readOnly={draft.provider === 'openai'} placeholder="https://your-resource.openai.azure.com"
                onChange={event => updateDraft('endpoint', event.target.value)}
                aria-labelledby="byo-endpoint-label" aria-describedby="byo-endpoint-help" />
              <small id="byo-endpoint-help">{text(
                'HTTPS Azure OpenAI / Microsoft Foundry resource origins or https://api.openai.com only. No custom hosts, API paths, query strings, ports, or credentials.',
                'HTTPS の Azure OpenAI / Microsoft Foundry リソースのオリジン、または https://api.openai.com のみ使用できます。任意のホスト、API パス、クエリ、ポート、資格情報は使用できません。',
              )}</small></label>
            <label><span>{text('Model / deployment', 'モデル / デプロイ名')}</span>
              <input value={draft.model} maxLength={128} required autoComplete="off" spellCheck={false}
                onChange={event => updateDraft('model', event.target.value)} /></label>
            <label><span id="byo-format-label">{text('API format', 'API 形式')}</span>
              <select aria-labelledby="byo-format-label" value={draft.apiFormat} onChange={event => updateDraft('apiFormat', event.target.value as BYOAIProfile['apiFormat'])}>
                <option value="responses">Responses API</option><option value="chat-completions">Chat Completions API</option>
              </select></label>
          </div>
          <fieldset className="byo-ai-capabilities">
            <legend>{text('Profile capabilities and limits', 'プロファイルの機能と上限')}</legend>
            <p>{text(
              'Set capabilities supported by this exact model/deployment. These settings control every BYO request, independently of managed Astra reasoning. Image analysis is blocked when vision is off.',
              'このモデル / デプロイが対応する機能を指定してください。管理対象の Astra の推論設定とは独立して、すべての BYO リクエストに適用されます。画像対応をオフにすると画像分析はブロックされます。',
            )}</p>
            <label className="byo-ai-checkbox"><input type="checkbox" checked={draft.isReasoning}
              onChange={event => {
                updateDraft('isReasoning', event.target.checked);
                if (!event.target.checked) setDraft(current => ({ ...current, reasoningEffort: 'none' }));
              }} />
              <span>{text('Supports reasoning', '推論に対応')}</span></label>
            <label className="byo-ai-checkbox"><input type="checkbox" checked={draft.supportsVision}
              onChange={event => updateDraft('supportsVision', event.target.checked)} />
              <span>{text('Supports image / vision input', '画像 / Vision 入力に対応')}</span></label>
            <div className="byo-ai-form-grid">
              <label><span id="byo-reasoning-label">{text('Profile reasoning effort', 'プロファイルの推論強度')}</span>
                <select aria-labelledby="byo-reasoning-label" value={draft.reasoningEffort} disabled={!draft.isReasoning}
                  onChange={event => updateDraft('reasoningEffort', event.target.value as BYOAIProfile['reasoningEffort'])}>
                  {BYOAI_REASONING_EFFORTS.map(effort => (
                    <option key={effort} value={effort}>{t(getReasoningEffortLabel(effort))}</option>
                  ))}
                </select></label>
              <label><span id="byo-output-label">{text('Maximum output tokens', '最大出力トークン数')}</span>
                <input type="number" min={1} max={32768} step={1} value={Number.isFinite(draft.maxCompletionTokens) ? draft.maxCompletionTokens : ''} required
                  aria-labelledby="byo-output-label" aria-describedby="byo-output-help"
                  onChange={event => updateDraft('maxCompletionTokens', event.target.valueAsNumber)} />
                <small id="byo-output-help">{text('1–32,768 tokens. The provider may impose a lower limit.', '1～32,768 トークン。プロバイダーの上限がこれより低い場合があります。')}</small></label>
            </div>
          </fieldset>
          {dirty && <p className="byo-ai-policy-note" role="status">{connection?.verified
            ? text('Unsaved profile name. Save it to keep the new name; this connection remains verified.', 'プロファイル名が未保存です。新しい名前を保持するには保存してください。接続の確認状態は保持されています。')
            : text('Unsaved changes. Save and test before activation.', '未保存の変更があります。保存とテストを行ってから有効化してください。')}</p>}
          <div className="byo-ai-actions">
            <button type="submit" disabled={testing || (!saved && limitReached)}>{text('Save profile', 'プロファイルを保存')}</button>
            {saved && <button type="button" onClick={() => setConfirmDelete(true)} disabled={active || testing}>
              <Trash2 size={16} aria-hidden="true" />{text('Delete profile', 'プロファイルを削除')}
            </button>}
          </div>
          {active && <p>{text(
            'This is the active profile. Explicitly switch to managed Astra or another verified profile before deleting it.',
            'このプロファイルは選択中です。削除する前に管理対象の Astra または別の確認済みプロファイルに明示的に切り替えてください。',
          )}</p>}
          {confirmDelete && !active && <div className="byo-ai-delete-confirm">
            <p>{text('Delete this inactive profile and its in-memory key?', 'この非アクティブなプロファイルとメモリ内のキーを削除しますか？')}</p>
            <button type="button" onClick={deleteProfile}>{text('Confirm delete', '削除を確定')}</button>
            <button type="button" onClick={() => setConfirmDelete(false)}>{text('Keep profile', 'プロファイルを保持')}</button>
          </div>}
        </form>
        <section className="byo-ai-test-section" aria-labelledby="byo-ai-test-heading">
          <h3 id="byo-ai-test-heading">{text('Key, test, then activate', 'キー入力・テスト・有効化')}</h3>
          <label htmlFor="byo-ai-key">{text('API key (tab memory only)', 'API キー（タブのメモリ内のみ）')}</label>
          <div className="byo-ai-key-row">
            <input id="byo-ai-key" type={showKey ? 'text' : 'password'} value={keyValue} autoComplete="off"
              spellCheck={false} autoCapitalize="none" maxLength={512} disabled={!saved || dirty}
              placeholder={state === 'key-required' ? text('Enter API key', 'API キーを入力') : text('Replace or re-enter key', 'キーを置き換え / 再入力')}
              onChange={event => {
                cancelTest();
                const value = event.target.value;
                setKeyValue(value);
                setError('');
                setNotice('');
                if (saved) {
                  try {
                    invalidateBYOAIProfile(saved.id);
                    setBYOAIApiKey(saved.id, value);
                  } catch (cause) {
                    if (getBYOAISettings().profiles.some(profile => profile.id === saved.id)) {
                      setBYOAIApiKey(saved.id, '');
                    }
                    safeFailure(cause);
                  }
                }
              }} />
            <button type="button" onClick={() => setShowKey(current => !current)} aria-pressed={showKey}
              disabled={!saved || dirty} aria-label={showKey ? text('Hide API key', 'API キーを非表示') : text('Show API key', 'API キーを表示')}>
              {showKey ? <EyeOff size={18} aria-hidden="true" /> : <Eye size={18} aria-hidden="true" />}
            </button>
            <button type="button" disabled={!saved || dirty} onClick={() => {
              cancelTest();
              if (saved) {
                try { setBYOAIApiKey(saved.id, ''); }
                catch (cause) { safeFailure(cause); }
              }
              setKeyValue('');
              setShowKey(false);
            }}>{text('Clear key', 'キーを消去')}</button>
          </div>
          {!saved && <p>{text('Save a public profile first, then enter its key.', 'まず公開プロファイルを保存してから、キーを入力してください。')}</p>}
          <p className={`byo-ai-status byo-ai-status--${displayState}`} role="status" aria-live="polite">
            {state === 'verified' ? <CheckCircle2 size={17} aria-hidden="true" /> : state === 'testing' ? <Loader2 size={17} className="spin" aria-hidden="true" /> : <KeyRound size={17} aria-hidden="true" />}
            {connectionStateLabel(displayState, language)}
          </p>
          {connection?.error && <p className="byo-ai-error" role="alert">
            {translate(connection.error.message)}
            {connection.error.requestId && !connection.error.message.includes('Request ID:')
              && <> {text('Request ID:', 'リクエスト ID:')} {connection.error.requestId}</>}
          </p>}
          {error && <p className="byo-ai-error" role="alert">{error}</p>}
          {notice && <p role="status">{notice}</p>}
          <div className="byo-ai-actions">
            {testing
              ? <button type="button" onClick={cancelTest}>{text('Cancel test', 'テストをキャンセル')}</button>
              : <button type="button" onClick={() => void testConnection()}
                disabled={!saved || dirty || !!storageError || !serverAllowed || state === 'key-required'}>
                {text('Test connection', '接続をテスト')}</button>}
            <button type="button" onClick={activate}
              disabled={!saved || dirty || testing || !!storageError || !serverAllowed || state !== 'verified' || active}>
              {active ? text('Profile selected', 'プロファイルを選択中') : text('Use this profile', 'このプロファイルを使用')}
            </button>
          </div>
        </section>
        <details className="byo-ai-managed-settings">
          <summary>{text('Managed Astra reasoning settings', '管理対象の Astra の推論設定')}</summary>
          <AstraReasoningSettings />
        </details>
      </div>
      <footer className="byo-ai-dialog-footer">
        <button type="button" onClick={onClose} disabled={testing}>{text('Done', '完了')}</button>
      </footer>
    </ModalScaffold>
  );
}
