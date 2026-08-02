// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

import React, { useCallback, useEffect, useMemo, useState } from 'react';
import {
  AlertTriangle,
  Check,
  Clock,
  Cloud,
  Copy,
  Edit3,
  Eye,
  FolderOpen,
  Link2,
  Loader,
  MessageSquare,
  RefreshCw,
  Trash2,
  X,
} from 'lucide-react';
import {
  CloudDiagramDocument,
  CloudDiagramShare,
  CloudDiagramSummary,
  CloudDiagramVersion,
  CloudDocumentContext,
  addCloudComment,
  createCloudShare,
  deleteCloudDiagram,
  getCloudDiagram,
  getCloudVersion,
  listCloudDiagrams,
  listCloudShares,
  listCloudVersions,
  revokeCloudShare,
} from '../services/cloudDiagramService';
import type { CloudSyncStatus } from '../hooks/useCloudDiagramSync';
import { useEscapeKey } from '../hooks/useEscapeKey';
import { useLanguage } from '../i18n/LanguageContext';
import { localize } from '../i18n/localization';
import './CloudWorkspaceModal.css';

interface CloudWorkspaceModalProps {
  isOpen: boolean;
  onClose: () => void;
  currentDocument: CloudDiagramDocument | null;
  currentContext: CloudDocumentContext | null;
  syncStatus: CloudSyncStatus;
  syncError: string;
  lastSavedAt: string | null;
  onOpenDocument: (
    document: CloudDiagramDocument,
    context: CloudDocumentContext,
  ) => void;
  onRestoreVersion: (
    version: CloudDiagramVersion,
    document: CloudDiagramDocument,
    context: CloudDocumentContext,
  ) => void;
  onDocumentUpdated: (document: CloudDiagramDocument) => void;
  onResetCurrent: () => void;
  onReloadRemote: () => Promise<CloudDiagramDocument | null>;
  onSaveAsCopy: () => Promise<CloudDiagramDocument | null>;
}

function ownerContext(documentId: string): CloudDocumentContext {
  return {
    documentId,
    access: 'owner',
    role: 'owner',
  };
}

const CloudWorkspaceModal: React.FC<CloudWorkspaceModalProps> = ({
  isOpen,
  onClose,
  currentDocument,
  currentContext,
  syncStatus,
  syncError,
  lastSavedAt,
  onOpenDocument,
  onRestoreVersion,
  onDocumentUpdated,
  onResetCurrent,
  onReloadRemote,
  onSaveAsCopy,
}) => {
  const { language, t } = useLanguage();
  const text = useCallback(
    (en: string, ja: string) => localize(language, { en, ja }),
    [language],
  );
  const [documents, setDocuments] = useState<CloudDiagramSummary[]>([]);
  const [selectedDocument, setSelectedDocument] = useState<CloudDiagramDocument | null>(null);
  const [versions, setVersions] = useState<CloudDiagramVersion[]>([]);
  const [shares, setShares] = useState<CloudDiagramShare[]>([]);
  const [isLoading, setIsLoading] = useState(false);
  const [isDetailsLoading, setIsDetailsLoading] = useState(false);
  const [operation, setOperation] = useState('');
  const [error, setError] = useState('');
  const [comment, setComment] = useState('');
  const [shareRole, setShareRole] = useState<'viewer' | 'editor'>('viewer');
  const [newShareUrl, setNewShareUrl] = useState('');

  useEscapeKey(isOpen, onClose);

  const selectedContext = useMemo<CloudDocumentContext | null>(() => {
    if (!selectedDocument) return null;
    if (
      currentDocument?.id === selectedDocument.id
      && currentContext?.access === 'shared'
    ) {
      return currentContext;
    }
    return ownerContext(selectedDocument.id);
  }, [currentContext, currentDocument?.id, selectedDocument]);

  const statusText = useMemo(() => {
    switch (syncStatus) {
      case 'saving':
        return text('Saving changes to the cloud...', '変更をクラウドへ保存しています...');
      case 'saved':
        return text('Cloud copy is up to date', 'クラウド版は最新です');
      case 'readonly':
        return text('Shared as read-only', '読み取り専用で共有されています');
      case 'conflict':
        return text('A newer cloud revision exists', 'クラウドに新しい版があります');
      case 'unavailable':
        return text('Cloud storage is not configured', 'クラウド保存が構成されていません');
      case 'offline':
        return text('Cloud storage is temporarily unavailable', 'クラウド保存を一時的に利用できません');
      case 'error':
        return text('Cloud synchronization needs attention', 'クラウド同期を確認してください');
      case 'loading':
        return text('Connecting to cloud storage...', 'クラウド保存へ接続しています...');
      default:
        return text('Cloud autosave starts when the diagram has content', '図に内容が追加されるとクラウド自動保存が始まります');
    }
  }, [syncStatus, text]);

  const statusClass = ['conflict', 'offline', 'unavailable', 'error'].includes(syncStatus)
    ? 'warning'
    : syncStatus === 'saved'
      ? 'success'
      : '';

  const formatDate = useCallback((value: string) => {
    if (!value) return '';
    return new Date(value).toLocaleString(language === 'ja' ? 'ja-JP' : 'en-US', {
      year: 'numeric',
      month: 'short',
      day: 'numeric',
      hour: '2-digit',
      minute: '2-digit',
    });
  }, [language]);

  const loadDetails = useCallback(async (
    document: CloudDiagramDocument,
    context: CloudDocumentContext,
  ) => {
    setSelectedDocument(document);
    setIsDetailsLoading(true);
    setError('');
    try {
      const [nextVersions, nextShares] = await Promise.all([
        listCloudVersions(context),
        context.access === 'owner' ? listCloudShares(document.id) : Promise.resolve([]),
      ]);
      setVersions(nextVersions);
      setShares(nextShares);
    } catch (loadError) {
      setVersions([]);
      setShares([]);
      setError(loadError instanceof Error ? loadError.message : text(
        'Failed to load cloud details.',
        'クラウドの詳細を読み込めませんでした。',
      ));
    } finally {
      setIsDetailsLoading(false);
    }
  }, [text]);

  const refreshDocuments = useCallback(async (preferredId?: string) => {
    setIsLoading(true);
    setError('');
    try {
      const owned = await listCloudDiagrams();
      const merged = [...owned];
      if (
        currentDocument
        && currentContext?.access === 'shared'
        && !merged.some((item) => item.id === currentDocument.id)
      ) {
        merged.unshift({
          id: currentDocument.id,
          diagramName: currentDocument.diagramName,
          createdAt: currentDocument.createdAt,
          updatedAt: currentDocument.updatedAt,
          revision: currentDocument.revision,
          serviceCount: currentDocument.payload.nodes.length,
          connectionCount: currentDocument.payload.edges.length,
          commentCount: currentDocument.comments.length,
          shareCount: 0,
          access: 'shared',
          role: currentContext.role,
          etag: currentDocument.etag,
        });
      }
      setDocuments(merged);

      const targetId = preferredId || currentDocument?.id || merged[0]?.id;
      if (!targetId) {
        setSelectedDocument(null);
        setVersions([]);
        setShares([]);
        return;
      }

      if (currentDocument?.id === targetId) {
        await loadDetails(
          currentDocument,
          currentContext || ownerContext(currentDocument.id),
        );
      } else {
        const nextDocument = await getCloudDiagram(targetId);
        await loadDetails(nextDocument, ownerContext(targetId));
      }
    } catch (loadError) {
      setDocuments([]);
      setSelectedDocument(currentDocument);
      setError(loadError instanceof Error ? loadError.message : text(
        'Failed to load cloud diagrams.',
        'クラウド図面を読み込めませんでした。',
      ));
    } finally {
      setIsLoading(false);
    }
  }, [currentContext, currentDocument, loadDetails, text]);

  useEffect(() => {
    if (isOpen) void refreshDocuments();
  }, [isOpen, refreshDocuments]);

  const handleSelect = async (summary: CloudDiagramSummary) => {
    if (selectedDocument?.id === summary.id) return;
    setIsDetailsLoading(true);
    setError('');
    try {
      if (currentDocument?.id === summary.id) {
        await loadDetails(
          currentDocument,
          currentContext || ownerContext(currentDocument.id),
        );
      } else {
        const nextDocument = await getCloudDiagram(summary.id);
        await loadDetails(nextDocument, ownerContext(summary.id));
      }
    } catch (loadError) {
      setError(loadError instanceof Error ? loadError.message : text(
        'Failed to open the cloud diagram.',
        'クラウド図面を開けませんでした。',
      ));
    } finally {
      setIsDetailsLoading(false);
    }
  };

  const handleDelete = async () => {
    if (!selectedDocument || selectedContext?.access !== 'owner') return;
    if (!window.confirm(text(
      `Delete "${selectedDocument.diagramName}" and all of its cloud snapshots?`,
      `「${selectedDocument.diagramName}」とクラウドスナップショットをすべて削除しますか？`,
    ))) return;

    setOperation('delete');
    setError('');
    try {
      await deleteCloudDiagram(selectedDocument.id, selectedDocument.etag);
      if (currentDocument?.id === selectedDocument.id) onResetCurrent();
      setSelectedDocument(null);
      await refreshDocuments();
    } catch (deleteError) {
      setError(deleteError instanceof Error ? deleteError.message : text(
        'Failed to delete the cloud diagram.',
        'クラウド図面を削除できませんでした。',
      ));
    } finally {
      setOperation('');
    }
  };

  const handleAddComment = async () => {
    const message = comment.trim();
    if (!selectedDocument || !selectedContext || !message) return;
    setOperation('comment');
    setError('');
    try {
      const updated = await addCloudComment(selectedContext, message);
      setSelectedDocument(updated);
      setComment('');
      if (currentDocument?.id === updated.id) onDocumentUpdated(updated);
      setDocuments((items) => items.map((item) => (
        item.id === updated.id
          ? {
              ...item,
              updatedAt: updated.updatedAt,
              revision: updated.revision,
              commentCount: updated.comments.length,
            }
          : item
      )));
    } catch (commentError) {
      setError(commentError instanceof Error ? commentError.message : text(
        'Failed to add the comment.',
        'コメントを追加できませんでした。',
      ));
    } finally {
      setOperation('');
    }
  };

  const handleCreateShare = async () => {
    if (!selectedDocument || selectedContext?.access !== 'owner') return;
    setOperation('share');
    setError('');
    setNewShareUrl('');
    try {
      const result = await createCloudShare(selectedDocument.id, shareRole);
      setNewShareUrl(result.url);
      setShares((items) => [result.share, ...items]);
      const refreshed = await getCloudDiagram(selectedDocument.id);
      setSelectedDocument(refreshed);
      if (currentDocument?.id === refreshed.id) onDocumentUpdated(refreshed);
      try {
        await navigator.clipboard.writeText(result.url);
      } catch {
        // The URL remains visible for manual copying when clipboard access is denied.
      }
    } catch (shareError) {
      setError(shareError instanceof Error ? shareError.message : text(
        'Failed to create the share link.',
        '共有リンクを作成できませんでした。',
      ));
    } finally {
      setOperation('');
    }
  };

  const handleCopyShare = async () => {
    if (!newShareUrl) return;
    try {
      await navigator.clipboard.writeText(newShareUrl);
    } catch {
      setError(text(
        'Clipboard access was denied. Copy the displayed URL manually.',
        'クリップボードへのアクセスが拒否されました。表示されたURLを手動でコピーしてください。',
      ));
    }
  };

  const handleRevokeShare = async (shareId: string) => {
    if (!selectedDocument || selectedContext?.access !== 'owner') return;
    setOperation(`revoke-${shareId}`);
    setError('');
    try {
      await revokeCloudShare(selectedDocument.id, shareId);
      setShares((items) => items.filter((share) => share.shareId !== shareId));
      const refreshed = await getCloudDiagram(selectedDocument.id);
      setSelectedDocument(refreshed);
      if (currentDocument?.id === refreshed.id) onDocumentUpdated(refreshed);
    } catch (revokeError) {
      setError(revokeError instanceof Error ? revokeError.message : text(
        'Failed to revoke the share link.',
        '共有リンクを無効化できませんでした。',
      ));
    } finally {
      setOperation('');
    }
  };

  const handleRestoreVersion = async (versionId: string) => {
    if (!selectedDocument || !selectedContext) return;
    if (!window.confirm(text(
      'Restore this cloud snapshot to the canvas?',
      'このクラウドスナップショットをキャンバスへ復元しますか？',
    ))) return;
    setOperation(`version-${versionId}`);
    setError('');
    try {
      const version = await getCloudVersion(selectedContext, versionId);
      onRestoreVersion(version, selectedDocument, selectedContext);
      onClose();
    } catch (versionError) {
      setError(versionError instanceof Error ? versionError.message : text(
        'Failed to restore the cloud snapshot.',
        'クラウドスナップショットを復元できませんでした。',
      ));
    } finally {
      setOperation('');
    }
  };

  const handleConflictAction = async (action: 'reload' | 'copy') => {
    setOperation(action);
    setError('');
    try {
      const updated = action === 'reload'
        ? await onReloadRemote()
        : await onSaveAsCopy();
      if (updated) await refreshDocuments(updated.id);
    } catch (actionError) {
      setError(actionError instanceof Error ? actionError.message : text(
        'The synchronization action failed.',
        '同期処理に失敗しました。',
      ));
    } finally {
      setOperation('');
    }
  };

  if (!isOpen) return null;

  return (
    <div className="modal-overlay" onClick={onClose}>
      <div
        className="modal-content cloud-workspace-modal"
        role="dialog"
        aria-modal="true"
        aria-label={text('Cloud workspace', 'クラウド ワークスペース')}
        tabIndex={-1}
        autoFocus
        onClick={(event) => event.stopPropagation()}
      >
        <div className="modal-header">
          <h2>
            <Cloud size={24} />
            {text('Cloud workspace', 'クラウド ワークスペース')}
          </h2>
          <button className="modal-close" onClick={onClose} title={t('Close')} aria-label={t('Close')}>
            <X size={24} />
          </button>
        </div>

        <div className={`cloud-sync-banner ${statusClass}`}>
          <span className="cloud-sync-icon">
            {syncStatus === 'saved' ? <Check size={18} /> : syncStatus === 'saving' ? <Loader size={18} className="spin" /> : <AlertTriangle size={18} />}
          </span>
          <span>
            <strong>{statusText}</strong>
            {(syncError || lastSavedAt) && (
              <small>
                {syncError || text(
                  `Last saved ${formatDate(lastSavedAt || '')}`,
                  `最終保存 ${formatDate(lastSavedAt || '')}`,
                )}
              </small>
            )}
          </span>
          {syncStatus === 'conflict' && (
            <span className="cloud-conflict-actions">
              <button onClick={() => void handleConflictAction('reload')} disabled={Boolean(operation)}>
                {text('Load cloud copy', 'クラウド版を読み込む')}
              </button>
              <button onClick={() => void handleConflictAction('copy')} disabled={Boolean(operation)}>
                {text('Save as copy', 'コピーとして保存')}
              </button>
            </span>
          )}
        </div>

        {error && <div className="cloud-workspace-error" role="alert">{error}</div>}

        <div className="cloud-workspace-body">
          <aside className="cloud-document-sidebar">
            <div className="cloud-section-heading">
              <span>{text('Cloud diagrams', 'クラウド図面')}</span>
              <button
                onClick={() => void refreshDocuments(selectedDocument?.id)}
                title={text('Refresh cloud diagrams', 'クラウド図面を更新')}
                aria-label={text('Refresh cloud diagrams', 'クラウド図面を更新')}
                disabled={isLoading}
              >
                <RefreshCw size={16} className={isLoading ? 'spin' : ''} />
              </button>
            </div>

            {isLoading && documents.length === 0 ? (
              <div className="cloud-empty-state"><Loader size={24} className="spin" /></div>
            ) : documents.length === 0 ? (
              <div className="cloud-empty-state">
                <Cloud size={32} />
                <p>{text(
                  'Your first cloud copy is created automatically after you add content.',
                  '内容を追加すると、最初のクラウド版が自動的に作成されます。',
                )}</p>
              </div>
            ) : (
              <div className="cloud-document-list">
                {documents.map((item) => (
                  <button
                    key={`${item.access}-${item.id}`}
                    className={selectedDocument?.id === item.id ? 'selected' : ''}
                    onClick={() => void handleSelect(item)}
                  >
                    <span className="cloud-document-name">{item.diagramName}</span>
                    <span className="cloud-document-meta">
                      {item.access === 'shared' ? text('Shared', '共有') : text('Owned', '所有')}
                      {' · '}
                      {formatDate(item.updatedAt)}
                    </span>
                    <span className="cloud-document-counts">
                      {item.serviceCount} {text('services', 'サービス')}
                      {' · '}
                      {text(`r${item.revision}`, `版 ${item.revision}`)}
                    </span>
                  </button>
                ))}
              </div>
            )}
          </aside>

          <main className="cloud-document-details">
            {!selectedDocument ? (
              <div className="cloud-empty-state large">
                <FolderOpen size={42} />
                <p>{text('Select a cloud diagram to view its details.', '詳細を表示するクラウド図面を選択してください。')}</p>
              </div>
            ) : (
              <>
                <div className="cloud-document-toolbar">
                  <div>
                    <h3>{selectedDocument.diagramName}</h3>
                    <p>
                      {selectedContext?.access === 'shared'
                        ? text(
                            `Shared ${selectedContext.role === 'editor' ? 'with editing' : 'as read-only'}`,
                            selectedContext.role === 'editor' ? '編集可能で共有' : '読み取り専用で共有',
                          )
                        : text('You own this diagram', 'この図面を所有しています')}
                      {' · '}
                      {text(`Revision ${selectedDocument.revision}`, `版 ${selectedDocument.revision}`)}
                    </p>
                  </div>
                  <div className="cloud-document-toolbar-actions">
                    <button
                      className="btn-secondary"
                      onClick={() => {
                        if (selectedContext) onOpenDocument(selectedDocument, selectedContext);
                        onClose();
                      }}
                    >
                      <FolderOpen size={16} />
                      {text('Open', '開く')}
                    </button>
                    {selectedContext?.access === 'owner' && (
                      <button
                        className="btn-secondary danger"
                        onClick={() => void handleDelete()}
                        disabled={operation === 'delete'}
                      >
                        <Trash2 size={16} />
                        {text('Delete', '削除')}
                      </button>
                    )}
                  </div>
                </div>

                {isDetailsLoading ? (
                  <div className="cloud-empty-state large"><Loader size={28} className="spin" /></div>
                ) : (
                  <div className="cloud-details-grid">
                    <section className="cloud-detail-card cloud-versions-card">
                      <div className="cloud-card-title">
                        <Clock size={18} />
                        <h4>{text('Cloud snapshots', 'クラウド スナップショット')}</h4>
                        <span>{versions.length}</span>
                      </div>
                      {versions.length === 0 ? (
                        <p className="cloud-card-empty">{text(
                          'Use Snapshot in the ribbon to preserve a named version.',
                          'リボンの「スナップショット」で名前付きの版を保存できます。',
                        )}</p>
                      ) : (
                        <div className="cloud-version-list">
                          {versions.map((version) => (
                            <div key={version.versionId} className="cloud-version-row">
                              <div>
                                <strong>{version.notes || text('Snapshot', 'スナップショット')}</strong>
                                <small>
                                  {formatDate(version.createdAt)}
                                  {' · '}
                                  {text(`Revision ${version.sourceRevision}`, `版 ${version.sourceRevision}`)}
                                </small>
                              </div>
                              <button
                                onClick={() => void handleRestoreVersion(version.versionId)}
                                disabled={operation === `version-${version.versionId}`}
                              >
                                {text('Restore', '復元')}
                              </button>
                            </div>
                          ))}
                        </div>
                      )}
                    </section>

                    <section className="cloud-detail-card">
                      <div className="cloud-card-title">
                        <MessageSquare size={18} />
                        <h4>{text('Comments', 'コメント')}</h4>
                        <span>{selectedDocument.comments.length}</span>
                      </div>
                      <div className="cloud-comment-list">
                        {selectedDocument.comments.length === 0 ? (
                          <p className="cloud-card-empty">{text('No comments yet.', 'コメントはまだありません。')}</p>
                        ) : selectedDocument.comments.map((item) => (
                          <article key={item.commentId}>
                            <header>
                              <strong>{item.authorEmail}</strong>
                              <time>{formatDate(item.createdAt)}</time>
                            </header>
                            <p>{item.message}</p>
                          </article>
                        ))}
                      </div>
                      <div className="cloud-comment-composer">
                        <textarea
                          value={comment}
                          onChange={(event) => setComment(event.target.value)}
                          maxLength={2000}
                          rows={2}
                          placeholder={text('Add a review comment...', 'レビューコメントを追加...')}
                        />
                        <button
                          onClick={() => void handleAddComment()}
                          disabled={!comment.trim() || operation === 'comment'}
                        >
                          {text('Comment', 'コメント')}
                        </button>
                      </div>
                    </section>

                    {selectedContext?.access === 'owner' && (
                      <section className="cloud-detail-card cloud-sharing-card">
                        <div className="cloud-card-title">
                          <Link2 size={18} />
                          <h4>{text('Secure share links', '安全な共有リンク')}</h4>
                          <span>{shares.length}</span>
                        </div>
                        <p className="cloud-share-note">
                          {text(
                            'Recipients must still be assigned to AzureDiagarm in Microsoft Entra ID.',
                            '受信者はMicrosoft Entra IDでAzureDiagarmへの割り当ても必要です。',
                          )}
                        </p>
                        <div className="cloud-share-controls">
                          <label>
                            <span>{text('Permission', '権限')}</span>
                            <select
                              value={shareRole}
                              onChange={(event) => setShareRole(event.target.value as 'viewer' | 'editor')}
                            >
                              <option value="viewer">{text('Viewer', '閲覧者')}</option>
                              <option value="editor">{text('Editor', '編集者')}</option>
                            </select>
                          </label>
                          <button onClick={() => void handleCreateShare()} disabled={operation === 'share'}>
                            <Link2 size={16} />
                            {text('Create link', 'リンクを作成')}
                          </button>
                        </div>
                        {newShareUrl && (
                          <div className="cloud-new-share">
                            <input value={newShareUrl} readOnly aria-label={text('New share URL', '新しい共有URL')} />
                            <button onClick={() => void handleCopyShare()} title={text('Copy link', 'リンクをコピー')}>
                              <Copy size={16} />
                            </button>
                          </div>
                        )}
                        <div className="cloud-share-list">
                          {shares.map((share) => (
                            <div key={share.shareId}>
                              <span className={`cloud-share-role ${share.role}`}>
                                {share.role === 'editor' ? <Edit3 size={14} /> : <Eye size={14} />}
                                {share.role === 'editor' ? text('Editor', '編集者') : text('Viewer', '閲覧者')}
                              </span>
                              <span>{formatDate(share.createdAt)}</span>
                              <button
                                onClick={() => void handleRevokeShare(share.shareId)}
                                disabled={operation === `revoke-${share.shareId}`}
                              >
                                {text('Revoke', '無効化')}
                              </button>
                            </div>
                          ))}
                        </div>
                      </section>
                    )}
                  </div>
                )}
              </>
            )}
          </main>
        </div>

        <div className="modal-actions">
          <span className="cloud-security-note">
            {text(
              'Encrypted in Azure Storage; access is enforced by Easy Auth and per-link permissions.',
              'Azure Storageで暗号化され、Easy Authとリンク権限でアクセスを制御します。',
            )}
          </span>
          <button className="btn-secondary" onClick={onClose}>{t('Close')}</button>
        </div>
      </div>
    </div>
  );
};

export default CloudWorkspaceModal;
