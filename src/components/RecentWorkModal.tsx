// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

import { useCallback, useEffect, useMemo, useRef, useState } from 'react';
import {
  AlertTriangle,
  Cloud,
  CloudOff,
  FolderOpen,
  History,
  Loader2,
  RefreshCw,
  Trash2,
  X,
} from 'lucide-react';
import type { CloudDiagramSummary } from '../services/cloudDiagramService';
import { listCloudDiagrams } from '../services/cloudDiagramService';
import {
  deleteRecentWork,
  isRecentWorkUnsynced,
  listRecentWork,
  type RecentWorkRecord,
} from '../services/recentWorkService';
import { useMediaQuery } from '../hooks/useMediaQuery';
import { useLanguage } from '../i18n/LanguageContext';
import { localize } from '../i18n/localization';
import { MEDIA_QUERIES } from '../styles/breakpoints';
import { OperationGeneration } from '../utils/operationGeneration';
import ResponsiveDrawer from './ResponsiveDrawer';
import './RecentWorkModal.css';

interface RecentWorkModalProps {
  isOpen: boolean;
  currentSessionId: string;
  currentLineageId: string;
  onClose: () => void;
  onResumeLocal: (record: RecentWorkRecord) => Promise<boolean>;
  onOpenCloud: (summary: CloudDiagramSummary) => Promise<boolean>;
}

function counts(record: RecentWorkRecord) {
  return {
    services: record.payload.nodes.filter(node => node?.type === 'azureNode').length,
    connections: record.payload.edges.length,
  };
}

export default function RecentWorkModal({
  isOpen,
  currentSessionId,
  currentLineageId,
  onClose,
  onResumeLocal,
  onOpenCloud,
}: RecentWorkModalProps) {
  const { language } = useLanguage();
  const text = useCallback(
    (en: string, ja: string) => localize(language, { en, ja }),
    [language],
  );
  const compact = useMediaQuery(MEDIA_QUERIES.narrow);
  const [localWork, setLocalWork] = useState<RecentWorkRecord[]>([]);
  const [cloudWork, setCloudWork] = useState<CloudDiagramSummary[]>([]);
  const [isLoading, setIsLoading] = useState(false);
  const [localError, setLocalError] = useState('');
  const [cloudError, setCloudError] = useState('');
  const [operation, setOperation] = useState('');
  const isOpenRef = useRef(isOpen);
  const generationRef = useRef(new OperationGeneration());
  isOpenRef.current = isOpen;

  const loadWork = useCallback(async () => {
    const generation = generationRef.current.advance();
    setIsLoading(true);
    setLocalError('');
    setCloudError('');
    const [localResult, cloudResult] = await Promise.allSettled([
      listRecentWork(),
      listCloudDiagrams(),
    ]);
    if (!isOpenRef.current || !generationRef.current.isCurrent(generation)) return;

    if (localResult.status === 'fulfilled') {
      setLocalWork(localResult.value);
    } else {
      console.error('Failed to load locally recoverable work:', localResult.reason);
      setLocalWork([]);
      setLocalError(text(
        'Local recovery storage could not be read.',
        'ローカル復旧ストレージを読み込めませんでした。',
      ));
    }

    if (cloudResult.status === 'fulfilled') {
      setCloudWork(cloudResult.value);
    } else {
      console.error('Failed to load recent cloud diagrams:', cloudResult.reason);
      setCloudWork([]);
      setCloudError(text(
        'Cloud diagrams are currently unavailable. Local recovery remains available.',
        'クラウド図面は現在利用できません。ローカル復旧は引き続き利用できます。',
      ));
    }
    setIsLoading(false);
  }, [text]);

  useEffect(() => {
    if (isOpen) {
      void loadWork();
      return;
    }
    generationRef.current.advance();
    setOperation('');
    setIsLoading(false);
  }, [isOpen, loadWork]);

  const closeModal = useCallback(() => {
    generationRef.current.advance();
    setOperation('');
    onClose();
  }, [onClose]);

  const attentionWork = useMemo(
    () => localWork.filter(isRecentWorkUnsynced),
    [localWork],
  );
  const attentionIds = useMemo(
    () => new Set(attentionWork.map(record => record.id)),
    [attentionWork],
  );
  const recoveredWork = useMemo(
    () => localWork.filter(record => (
      !attentionIds.has(record.id)
      && record.sessionId !== currentSessionId
    )),
    [attentionIds, currentSessionId, localWork],
  );
  const localDrafts = useMemo(
    () => localWork.filter(record => (
      !attentionIds.has(record.id)
      && record.sessionId === currentSessionId
      && !record.cloudDocumentId
    )),
    [attentionIds, currentSessionId, localWork],
  );

  const formatDate = useCallback((timestamp: number | string) => (
    new Intl.DateTimeFormat(language === 'ja' ? 'ja-JP' : 'en-US', {
      month: 'short',
      day: 'numeric',
      year: 'numeric',
      hour: '2-digit',
      minute: '2-digit',
    }).format(new Date(timestamp))
  ), [language]);

  const resumeLocal = async (record: RecentWorkRecord) => {
    setOperation(`local-${record.id}`);
    try {
      if (await onResumeLocal(record)) closeModal();
    } finally {
      if (isOpenRef.current) setOperation('');
    }
  };

  const openCloud = async (summary: CloudDiagramSummary) => {
    setOperation(`cloud-${summary.id}`);
    try {
      if (await onOpenCloud(summary)) closeModal();
    } finally {
      if (isOpenRef.current) setOperation('');
    }
  };

  const removeLocal = async (record: RecentWorkRecord) => {
    if (!window.confirm(text(
      `Delete the local recovery copy "${record.diagramName}"?`,
      `ローカル復旧コピー「${record.diagramName}」を削除しますか？`,
    ))) return;
    setOperation(`delete-${record.id}`);
    try {
      await deleteRecentWork(record.id);
      if (isOpenRef.current) {
        setLocalWork(items => items.filter(item => item.id !== record.id));
      }
    } catch (error) {
      console.error('Failed to delete recoverable work:', error);
      if (isOpenRef.current) {
        setLocalError(text(
          'The local recovery copy could not be deleted.',
          'ローカル復旧コピーを削除できませんでした。',
        ));
      }
    } finally {
      if (isOpenRef.current) setOperation('');
    }
  };

  const renderLocalCard = (
    record: RecentWorkRecord,
    status: { label: string; tone: string },
  ) => {
    const recordCounts = counts(record);
    const isCurrent = record.lineageId === currentLineageId;
    return (
      <article className="recent-work-card" key={record.id}>
        <div className="recent-work-card-main">
          <div className="recent-work-card-title">
            <strong>{record.diagramName || text('Untitled Architecture', '無題のアーキテクチャ')}</strong>
            <span className={`recent-work-badge recent-work-badge--${status.tone}`}>
              {status.label}
            </span>
            {isCurrent && (
              <span className="recent-work-badge recent-work-badge--current">
                {text('Current', '現在')}
              </span>
            )}
          </div>
          <p>{formatDate(record.updatedAt)}</p>
          <small>
            {recordCounts.services} {text('services', 'サービス')}
            {' · '}
            {recordCounts.connections} {text('connections', '接続')}
          </small>
        </div>
        <div className="recent-work-card-actions">
          <button
            type="button"
            className="recent-work-primary-action"
            onClick={() => void resumeLocal(record)}
            disabled={Boolean(operation)}
          >
            {operation === `local-${record.id}`
              ? <Loader2 className="recent-work-spin" size={16} />
              : <FolderOpen size={16} />}
            {isCurrent ? text('Return', '戻る') : text('Resume', '再開')}
          </button>
          <button
            type="button"
            className="recent-work-delete"
            onClick={() => void removeLocal(record)}
            disabled={Boolean(operation) || isCurrent}
            title={isCurrent
              ? text('The active recovery copy cannot be deleted.', '使用中の復旧コピーは削除できません。')
              : text('Delete local recovery copy', 'ローカル復旧コピーを削除')}
            aria-label={text(
              `Delete local recovery copy ${record.diagramName}`,
              `ローカル復旧コピー ${record.diagramName} を削除`,
            )}
          >
            <Trash2 size={16} />
          </button>
        </div>
      </article>
    );
  };

  const hasAnyWork = (
    attentionWork.length > 0
    || recoveredWork.length > 0
    || localDrafts.length > 0
    || cloudWork.length > 0
  );

  return (
    <ResponsiveDrawer
      isOpen={isOpen}
      modal
      placement={compact ? 'bottom' : 'center'}
      className="recent-work-modal"
      backdropClassName="recent-work-overlay"
      ariaLabel={text('Recent work', '最近の作業')}
      onClose={closeModal}
      backgroundSelectors={[
        '.app > .app-header',
        '.app > .workspace',
        '.app > .arch-chat-panel',
      ]}
    >
      <header className="recent-work-header">
        <div>
          <History size={24} aria-hidden="true" />
          <span>
            <h2>{text('Resume recent work', '最近の作業を再開')}</h2>
            <p>{text(
              'Continue local drafts, recover interrupted sessions, or open a cloud diagram.',
              'ローカル下書き、前回中断したセッション、クラウド図面を続行します。',
            )}</p>
          </span>
        </div>
        <button
          type="button"
          className="recent-work-close"
          onClick={closeModal}
          aria-label={text('Close recent work', '最近の作業を閉じる')}
        >
          <X size={22} />
        </button>
      </header>

      <div className="recent-work-toolbar">
        <span>{text(
          'Recovery copies stay in this browser for up to 45 days. AI endpoint credentials are never included.',
          '復旧コピーはこのブラウザーに最大45日間保持されます。AIエンドポイントの資格情報は含まれません。',
        )}</span>
        <button type="button" onClick={() => void loadWork()} disabled={isLoading || Boolean(operation)}>
          <RefreshCw className={isLoading ? 'recent-work-spin' : ''} size={16} />
          {text('Refresh', '更新')}
        </button>
      </div>

      {(localError || cloudError) && (
        <div className="recent-work-errors" role="status">
          {localError && <p><AlertTriangle size={16} />{localError}</p>}
          {cloudError && <p><CloudOff size={16} />{cloudError}</p>}
        </div>
      )}

      <main className="recent-work-body">
        {isLoading && !hasAnyWork ? (
          <div className="recent-work-empty">
            <Loader2 className="recent-work-spin" size={30} />
            <p>{text('Finding recent work...', '最近の作業を検索しています...')}</p>
          </div>
        ) : !hasAnyWork ? (
          <div className="recent-work-empty">
            <History size={42} />
            <h3>{text('No recent work yet', '最近の作業はまだありません')}</h3>
            <p>{text(
              'Start a diagram and it will appear here automatically.',
              '図面を開始すると、自動的にここへ表示されます。',
            )}</p>
          </div>
        ) : (
          <div className="recent-work-sections">
            {attentionWork.length > 0 && (
              <section>
                <div className="recent-work-section-title recent-work-section-title--attention">
                  <AlertTriangle size={18} />
                  <span>
                    <h3>{text('Needs attention', '確認が必要')}</h3>
                    <p>{text('Unsynced or conflicted work preserved locally.', '未同期または競合した作業をローカルに保持しています。')}</p>
                  </span>
                </div>
                <div className="recent-work-list">
                  {attentionWork.map(record => renderLocalCard(record, {
                    label: text('Unsynced', '未同期'),
                    tone: 'warning',
                  }))}
                </div>
              </section>
            )}

            {recoveredWork.length > 0 && (
              <section>
                <div className="recent-work-section-title">
                  <History size={18} />
                  <span>
                    <h3>{text('Recovered sessions', '復旧したセッション')}</h3>
                    <p>{text('Copies left by another browser session.', '別のブラウザーセッションが残したコピーです。')}</p>
                  </span>
                </div>
                <div className="recent-work-list">
                  {recoveredWork.map(record => renderLocalCard(record, {
                    label: text('Recovered', '復旧'),
                    tone: 'recovered',
                  }))}
                </div>
              </section>
            )}

            {localDrafts.length > 0 && (
              <section>
                <div className="recent-work-section-title">
                  <FolderOpen size={18} />
                  <span>
                    <h3>{text('Local drafts', 'ローカル下書き')}</h3>
                    <p>{text('Work saved only in this browser.', 'このブラウザーだけに保存された作業です。')}</p>
                  </span>
                </div>
                <div className="recent-work-list">
                  {localDrafts.map(record => renderLocalCard(record, {
                    label: text('Local', 'ローカル'),
                    tone: 'local',
                  }))}
                </div>
              </section>
            )}

            {cloudWork.length > 0 && (
              <section>
                <div className="recent-work-section-title">
                  <Cloud size={18} />
                  <span>
                    <h3>{text('Cloud diagrams', 'クラウド図面')}</h3>
                    <p>{text('Your most recently updated cloud work.', '最近更新したクラウド作業です。')}</p>
                  </span>
                </div>
                <div className="recent-work-list">
                  {cloudWork.map(summary => (
                    <article className="recent-work-card" key={`${summary.access}-${summary.id}`}>
                      <div className="recent-work-card-main">
                        <div className="recent-work-card-title">
                          <strong>{summary.diagramName}</strong>
                          <span className="recent-work-badge recent-work-badge--cloud">
                            {summary.access === 'shared'
                              ? text('Shared', '共有')
                              : text('Cloud', 'クラウド')}
                          </span>
                        </div>
                        <p>{formatDate(summary.updatedAt)}</p>
                        <small>
                          {summary.serviceCount} {text('services', 'サービス')}
                          {' · '}
                          {summary.connectionCount} {text('connections', '接続')}
                          {' · '}
                          {text(`Revision ${summary.revision}`, `版 ${summary.revision}`)}
                        </small>
                      </div>
                      <div className="recent-work-card-actions">
                        <button
                          type="button"
                          className="recent-work-primary-action"
                          onClick={() => void openCloud(summary)}
                          disabled={Boolean(operation)}
                        >
                          {operation === `cloud-${summary.id}`
                            ? <Loader2 className="recent-work-spin" size={16} />
                            : <Cloud size={16} />}
                          {text('Open', '開く')}
                        </button>
                      </div>
                    </article>
                  ))}
                </div>
              </section>
            )}
          </div>
        )}
      </main>

      <footer className="recent-work-footer">
        <span>{text(
          'Local copies never leave this browser unless you explicitly open or save them to the cloud.',
          'ローカルコピーは、明示的に開くかクラウド保存しない限り、このブラウザーの外へ送信されません。',
        )}</span>
        <button type="button" onClick={closeModal}>{text('Close', '閉じる')}</button>
      </footer>
    </ResponsiveDrawer>
  );
}
