// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

import { useId, useState } from 'react';
import { PlugZap } from 'lucide-react';
import {
  getBYOAIConnectionState, getBYOAISettings, selectBYOAIProfile, useBYOAISettings,
} from '../stores/byoAISettingsStore';
import { useRuntimeConfig } from '../services/runtimeConfig';
import { useLanguage, type Language } from '../i18n/LanguageContext';
import { localize } from '../i18n/localization';
import './AIConnectionSelector.css';

export function connectionStateLabel(state: string, language: Language): string {
  const labels: Record<string, { en: string; ja: string }> = {
    'key-required': { en: 'Key required', ja: 'キーが必要' },
    unverified: { en: 'Unverified', ja: '未確認' },
    testing: { en: 'Testing', ja: 'テスト中' },
    verified: { en: 'Verified', ja: '確認済み' },
    failed: { en: 'Failed', ja: '失敗' },
    'admin-disabled': { en: 'Administrator disabled', ja: '管理者によって無効' },
    'missing-profile': { en: 'Missing profile', ja: 'プロファイルが見つかりません' },
    'policy-checking': { en: 'Checking server policy', ja: 'サーバー ポリシーを確認中' },
    'policy-unavailable': { en: 'Server policy unavailable', ja: 'サーバー ポリシーを確認できません' },
    'storage-unavailable': { en: 'Storage unavailable', ja: 'ストレージを利用できません' },
  };
  return localize(language, labels[state] ?? { en: 'Unavailable', ja: '利用不可' });
}

interface AIConnectionSelectorProps {
  onConfigureConnections?: () => void;
  disabled?: boolean;
  compact?: boolean;
}

export default function AIConnectionSelector({
  onConfigureConnections, disabled = false, compact = false,
}: AIConnectionSelectorProps) {
  const { language, translate } = useLanguage();
  const { storageError } = useBYOAISettings();
  const settings = getBYOAISettings();
  const policy = useRuntimeConfig();
  const [error, setError] = useState('');
  const id = useId();
  const text = (en: string, ja: string) => localize(language, { en, ja });
  const active = settings.profiles.find(profile => profile.id === settings.activeProfileId);
  const serverAllowed = policy.status === 'ready' && policy.bringYourOwnAI;
  const activeState = settings.activeProfileId ? getBYOAIConnectionState(settings.activeProfileId) : null;
  const displayState = (status: string) => storageError ? 'storage-unavailable' : policy.status === 'error' ? 'policy-unavailable'
    : policy.status !== 'ready' ? 'policy-checking' : !serverAllowed ? 'admin-disabled' : status;
  const choose = (profileId: string | null) => {
    try {
      selectBYOAIProfile(profileId);
      setError('');
    } catch (cause) {
      setError(translate(cause instanceof Error ? cause.message : 'The AI connection is unavailable.'));
    }
  };

  return (
    <section className={`ai-connection-selector${compact ? ' ai-connection-selector--compact' : ''}`}
      aria-label={text('AI connection', 'AI 接続')}>
      <label htmlFor={id}>{text('AI connection', 'AI 接続')}</label>
      <div className="ai-connection-selector-controls">
        <select id={id} value={settings.activeProfileId ?? ''} disabled={disabled}
          onChange={event => choose(event.target.value || null)}>
          <option value="">{text('Managed GPT-6 Astra', '管理対象の GPT-6 Astra')}</option>
          {settings.activeProfileId && !active && (
            <option value={settings.activeProfileId} disabled>{text('Missing BYO connection', 'BYO 接続が見つかりません')}</option>
          )}
          {settings.profiles.map(profile => {
            const status = getBYOAIConnectionState(profile.id);
            return (
              <option key={profile.id} value={profile.id}
                disabled={!!storageError || !serverAllowed || !status.verified}>
                {profile.name} · {profile.model} — {connectionStateLabel(displayState(status.status), language)}
              </option>
            );
          })}
        </select>
        {onConfigureConnections && (
          <button type="button" onClick={onConfigureConnections} disabled={disabled}>
            <PlugZap size={16} aria-hidden="true" />
            {text('AI connections', 'AI 接続の設定')}
          </button>
        )}
      </div>
      {settings.activeProfileId !== null && (
        <div className="ai-connection-current">
          <span>{text('BYO profile', 'BYO プロファイル')}: <strong>{active?.name ?? text('Missing profile', 'プロファイルが見つかりません')}</strong>
            {' · '}{connectionStateLabel(displayState(activeState?.status ?? 'missing-profile'), language)}</span>
          {(!activeState?.verified || !serverAllowed || storageError) && (
            <p role="status">{text(
              'This selected connection cannot run. Re-enter its key and test it, or explicitly choose managed Astra. No automatic fallback.',
              '選択中の接続は実行できません。キーを再入力してテストするか、管理対象の Astra を明示的に選択してください。自動切り替えは行いません。',
            )}</p>
          )}
          <button type="button" onClick={() => choose(null)} disabled={disabled}>
            {text('Use managed Astra', '管理対象の Astra を使用')}
          </button>
        </div>
      )}
      {storageError && <div className="ai-connection-error">
        <p role="alert">{translate(storageError.message)}</p>
        {settings.activeProfileId === null && <button type="button" onClick={() => choose(null)} disabled={disabled}>
          {text('Use managed Astra', '管理対象の Astra を使用')}
        </button>}
      </div>}
      {error && <p className="ai-connection-error" role="alert">{error}</p>}
    </section>
  );
}
