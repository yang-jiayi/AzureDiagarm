// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

import { AlertTriangle, CheckCircle2, Cloud, CloudOff, Loader2, Lock } from 'lucide-react';
import type { CloudSyncStatus } from '../hooks/useCloudDiagramSync';
import { useLanguage } from '../i18n/LanguageContext';
import { localize } from '../i18n/localization';
import './DocumentStatus.css';

interface DocumentStatusProps {
  status: CloudSyncStatus;
  lastSavedAt: string | null;
  hasCloudDocument: boolean;
  onOpen: () => void;
}

export default function DocumentStatus({
  status,
  lastSavedAt,
  hasCloudDocument,
  onOpen,
}: DocumentStatusProps) {
  const { language } = useLanguage();
  const savedTime = lastSavedAt
    ? new Intl.DateTimeFormat(language === 'ja' ? 'ja-JP' : 'en-US', {
      hour: '2-digit',
      minute: '2-digit',
    }).format(new Date(lastSavedAt))
    : null;

  const presentation = (() => {
    switch (status) {
      case 'saving':
        return {
          tone: 'working',
          icon: <Loader2 className="document-status-spinner" size={15} />,
          label: localize(language, { en: 'Saving', ja: '保存中' }),
        };
      case 'saved':
        return {
          tone: 'success',
          icon: <CheckCircle2 size={15} />,
          label: localize(language, { en: 'Saved', ja: '保存済み' }),
        };
      case 'readonly':
        return {
          tone: 'readonly',
          icon: <Lock size={15} />,
          label: localize(language, { en: 'Read-only', ja: '閲覧のみ' }),
        };
      case 'conflict':
      case 'error':
        return {
          tone: 'danger',
          icon: <AlertTriangle size={15} />,
          label: localize(language, { en: 'Sync issue', ja: '同期の問題' }),
        };
      case 'offline':
      case 'unavailable':
        return {
          tone: 'offline',
          icon: <CloudOff size={15} />,
          label: localize(language, { en: 'Local only', ja: 'ローカルのみ' }),
        };
      case 'loading':
        return {
          tone: 'working',
          icon: <Loader2 className="document-status-spinner" size={15} />,
          label: localize(language, { en: 'Connecting', ja: '接続中' }),
        };
      default:
        return {
          tone: hasCloudDocument ? 'neutral' : 'local',
          icon: <Cloud size={15} />,
          label: hasCloudDocument
            ? localize(language, { en: 'Cloud', ja: 'クラウド' })
            : localize(language, { en: 'Local draft', ja: 'ローカル下書き' }),
        };
    }
  })();

  const detail = savedTime
    ? localize(language, { en: `at ${savedTime}`, ja: `${savedTime}` })
    : localize(language, { en: 'Open workspace', ja: 'ワークスペースを開く' });

  return (
    <button
      type="button"
      className={`document-status document-status--${presentation.tone}`}
      onClick={onOpen}
      aria-label={localize(language, {
        en: `Document status: ${presentation.label}. ${detail}`,
        ja: `ドキュメント状態: ${presentation.label}。${detail}`,
      })}
      title={localize(language, {
        en: `${presentation.label} · ${detail}`,
        ja: `${presentation.label} · ${detail}`,
      })}
    >
      <span aria-hidden="true">{presentation.icon}</span>
      <span className="document-status-copy" aria-live="polite">
        <strong>{presentation.label}</strong>
        <small>{detail}</small>
      </span>
    </button>
  );
}
