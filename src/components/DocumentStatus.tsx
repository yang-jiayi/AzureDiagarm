// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

import {
  AlertTriangle,
  CheckCircle2,
  Cloud,
  CloudOff,
  Loader2,
  Lock,
} from 'lucide-react';
import type { CloudDiagramAccess, CloudDiagramRole } from '../services/cloudDiagramService';
import type { DraftSaveStatus } from '../hooks/useDraftAutosave';
import type { CloudSyncStatus } from '../hooks/useCloudDiagramSync';
import { useLanguage } from '../i18n/LanguageContext';
import { localize } from '../i18n/localization';
import './DocumentStatus.css';

interface DocumentStatusProps {
  documentName: string;
  status: CloudSyncStatus;
  localSaveStatus: DraftSaveStatus;
  localSaveError?: Error | null;
  localSavedAt?: number;
  lastSavedAt: string | null;
  hasCloudDocument: boolean;
  cloudAccess?: CloudDiagramAccess;
  cloudRole?: CloudDiagramRole;
  shareCount?: number;
  onOpen: () => void;
  onRetryLocalSave: () => void;
}

export default function DocumentStatus({
  documentName,
  status,
  localSaveStatus,
  localSaveError,
  localSavedAt,
  lastSavedAt,
  hasCloudDocument,
  cloudAccess,
  cloudRole,
  shareCount,
  onOpen,
  onRetryLocalSave,
}: DocumentStatusProps) {
  const { language } = useLanguage();
  const formatTime = (time: string | number) => new Intl.DateTimeFormat(
    language === 'ja' ? 'ja-JP' : 'en-US',
    { hour: '2-digit', minute: '2-digit' },
  ).format(new Date(time));

  const cloudState = (() => {
    switch (status) {
      case 'saving':
        return {
          tone: 'working',
          icon: <Loader2 className="document-status-spinner" size={15} />,
          label: localize(language, { en: 'Saving to cloud', ja: 'クラウドに保存中' }),
        };
      case 'saved':
        return {
          tone: 'success',
          icon: <CheckCircle2 size={15} />,
          label: localize(language, { en: 'Saved to cloud', ja: 'クラウド保存済み' }),
        };
      case 'readonly':
        return {
          tone: 'readonly',
          icon: <Lock size={15} />,
          label: localize(language, { en: 'Cloud · read-only', ja: 'クラウド · 閲覧のみ' }),
        };
      case 'conflict':
      case 'error':
        return {
          tone: 'danger',
          icon: <AlertTriangle size={15} />,
          label: localize(language, { en: 'Cloud sync issue', ja: 'クラウド同期の問題' }),
        };
      case 'offline':
      case 'unavailable':
        return {
          tone: 'offline',
          icon: <CloudOff size={15} />,
          label: localize(language, {
            en: 'Cloud unavailable · local copy only',
            ja: 'クラウド利用不可 · ローカルのみ',
          }),
        };
      case 'loading':
        return {
          tone: 'working',
          icon: <Loader2 className="document-status-spinner" size={15} />,
          label: localize(language, { en: 'Connecting to cloud', ja: 'クラウドに接続中' }),
        };
      default:
        return {
          tone: 'neutral',
          icon: <Cloud size={15} />,
          label: localize(language, { en: 'Cloud document', ja: 'クラウド ドキュメント' }),
        };
    }
  })();

  const localState = (() => {
    switch (localSaveStatus) {
      case 'loading':
        return localize(language, { en: 'Opening local draft', ja: 'ローカル下書きを開いています' });
      case 'saving':
        return localize(language, { en: 'Saving local copy', ja: 'ローカル コピーを保存中' });
      case 'saved':
        return localize(language, { en: 'Local copy saved', ja: 'ローカル コピー保存済み' });
      case 'error':
        return localize(language, { en: 'Local draft not saved', ja: 'ローカル下書き未保存' });
      default:
        return localize(language, {
          en: 'Local draft · autosave ready',
          ja: 'ローカル下書き · 自動保存の準備完了',
        });
    }
  })();
  const localSaveErrorMessage = localSaveStatus === 'error'
    ? localSaveError?.message.trim().replace(/\s+/g, ' ')
    : undefined;
  const localSaveErrorDisclosure = localSaveErrorMessage
    ? localize(language, {
        en: `Local save error: ${localSaveErrorMessage}`,
        ja: `ローカル保存エラー: ${localSaveErrorMessage}`,
      })
    : undefined;

  const cloudBacked = hasCloudDocument || status !== 'idle';
  const locationState = cloudBacked ? cloudState : {
    tone: localSaveStatus === 'error'
      ? 'danger'
      : localSaveStatus === 'saving' || localSaveStatus === 'loading'
        ? 'working'
        : localSaveStatus === 'saved'
          ? 'success'
          : 'local',
    icon: localSaveStatus === 'error'
      ? <AlertTriangle size={15} />
      : localSaveStatus === 'saving' || localSaveStatus === 'loading'
        ? <Loader2 className="document-status-spinner" size={15} />
        : localSaveStatus === 'saved'
          ? <CheckCircle2 size={15} />
          : <CloudOff size={15} />,
    label: localState,
  };

  const sharingState = !hasCloudDocument
    ? localize(language, { en: 'Not shared', ja: '未共有' })
    : cloudAccess === 'shared'
      ? cloudRole === 'viewer'
        ? localize(language, { en: 'Shared · view access', ja: '共有済み · 閲覧権限' })
        : cloudRole === 'editor'
          ? localize(language, { en: 'Shared · edit access', ja: '共有済み · 編集権限' })
          : localize(language, { en: 'Shared access', ja: '共有アクセス' })
      : shareCount === undefined
        ? localize(language, { en: 'Sharing status unavailable', ja: '共有状態を取得できません' })
        : shareCount > 0
          ? localize(language, {
              en: `Shared · ${shareCount} link${shareCount === 1 ? '' : 's'}`,
              ja: `共有済み · ${shareCount} 件のリンク`,
            })
          : localize(language, { en: 'Not shared', ja: '未共有' });

  const savedAt = hasCloudDocument ? lastSavedAt : localSavedAt;
  const savedTime = savedAt !== null && savedAt !== undefined
    ? localize(language, {
        en: `${hasCloudDocument ? 'Last cloud save' : 'Last local save'} ${formatTime(savedAt)}`,
        ja: `${hasCloudDocument ? '最終クラウド保存' : '最終ローカル保存'} ${formatTime(savedAt)}`,
      })
    : null;
  const currentName = documentName.trim() || localize(language, {
    en: 'Untitled Architecture',
    ja: '無題のアーキテクチャ',
  });
  const secondaryDetail = cloudBacked
    ? `${sharingState} · ${localState}`
    : sharingState;
  const accessibleSummary = [
    localize(language, {
      en: `Document status: ${locationState.label}`,
      ja: `ドキュメント状態: ${locationState.label}`,
    }),
    localize(language, { en: `Document: ${currentName}`, ja: `ドキュメント: ${currentName}` }),
    savedTime,
    secondaryDetail,
    localSaveErrorDisclosure,
  ].filter(Boolean).join('. ');
  const openWorkspaceLabel = localize(language, {
    en: 'Open document and sharing workspace',
    ja: 'ドキュメントと共有ワークスペースを開く',
  });

  return (
    <div className="document-status-group">
      <button
        type="button"
        className={`document-status document-status--${locationState.tone}`}
        onClick={onOpen}
        aria-label={`${accessibleSummary}. ${openWorkspaceLabel}`}
        title={`${accessibleSummary}. ${openWorkspaceLabel}`}
      >
        <span aria-hidden="true">{locationState.icon}</span>
        <span className="document-status-copy" aria-live="polite">
          <strong title={currentName}>{currentName}</strong>
          <small>{[locationState.label, savedTime].filter(Boolean).join(' · ')}</small>
          <small>{secondaryDetail}</small>
          {localSaveErrorDisclosure && (
            <small className="document-status-error" aria-hidden="true" title={localSaveErrorMessage}>
              {localSaveErrorDisclosure}
            </small>
          )}
        </span>
      </button>
      {localSaveStatus === 'error' && (
        <button
          type="button"
          className="document-status-retry"
          onClick={onRetryLocalSave}
          aria-label={localize(language, {
            en: localSaveErrorMessage
              ? `Retry saving the local draft: ${localSaveErrorMessage}`
              : 'Retry saving the local draft',
            ja: localSaveErrorMessage
              ? `ローカル下書きの保存を再試行: ${localSaveErrorMessage}`
              : 'ローカル下書きの保存を再試行',
          })}
          title={localSaveErrorMessage || localize(language, {
            en: 'Retry local save',
            ja: 'ローカル保存を再試行',
          })}
        >
          {localize(language, { en: 'Retry', ja: '再試行' })}
        </button>
      )}
    </div>
  );
}
