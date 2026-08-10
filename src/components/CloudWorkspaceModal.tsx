// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

import React, { useCallback, useEffect, useMemo, useRef, useState } from 'react';
import {
  AlertTriangle,
  Check,
  CheckCircle2,
  Clock,
  Cloud,
  Copy,
  Download,
  Edit3,
  Eye,
  FolderOpen,
  Link2,
  Loader,
  MapPin,
  MessageSquare,
  Plus,
  RefreshCw,
  RotateCcw,
  Send,
  ThumbsUp,
  Trash2,
  X,
} from 'lucide-react';
import {
  CloudCommentAnchor,
  CloudDiagramApiError,
  CloudDiagramDocument,
  CloudDiagramReview,
  CloudReviewAction,
  CloudDiagramShare,
  CloudDiagramSummary,
  CloudDiagramVersion,
  CloudDiagramVersionSummary,
  CloudDocumentContext,
  addCloudComment,
  createCloudShare,
  createCloudVersion,
  deleteCloudDiagram,
  getCloudDiagram,
  getSharedCloudDiagram,
  getCloudVersion,
  listCloudDiagrams,
  listCloudShares,
  listCloudVersions,
  revokeCloudShare,
  setCloudCommentResolved,
  updateCloudReview,
} from '../services/cloudDiagramService';
import type { CloudSyncStatus } from '../hooks/useCloudDiagramSync';
import { useMediaQuery } from '../hooks/useMediaQuery';
import { useLanguage } from '../i18n/LanguageContext';
import { localize } from '../i18n/localization';
import { OperationGeneration } from '../utils/operationGeneration';
import { MEDIA_QUERIES } from '../styles/breakpoints';
import { buildCloudReviewReport } from '../utils/cloudReviewReport';
import { toFileNameSegment } from '../utils/fileName';
import ResponsiveDrawer from './ResponsiveDrawer';
import './CloudWorkspaceModal.css';

interface CloudWorkspaceModalProps {
  isOpen: boolean;
  onClose: () => void;
  currentDocument: CloudDiagramDocument | null;
  currentContext: CloudDocumentContext | null;
  syncStatus: CloudSyncStatus;
  syncError: string;
  lastSavedAt: string | null;
  hasLocalDraft: boolean;
  isCloseBlocked?: boolean;
  onOpenDocument: (
    document: CloudDiagramDocument,
    context: CloudDocumentContext,
  ) => void;
  onRestoreVersion: (
    version: CloudDiagramVersion,
    document: CloudDiagramDocument,
    context: CloudDocumentContext,
  ) => Promise<boolean>;
  onDocumentUpdated: (document: CloudDiagramDocument) => void;
  onResetCurrent: (documentId: string) => void;
  onSaveCurrent: (options?: { force?: boolean }) => Promise<CloudDiagramDocument | null>;
  onReloadRemote: () => Promise<CloudDiagramDocument | null>;
  onSaveAsCopy: () => Promise<CloudDiagramDocument | null>;
  onSaveAsDetachedCopy: () => Promise<CloudDiagramDocument | null>;
  onCloudConflict: (
    documentId: string,
    error: Error,
    expectedRevision?: number,
    expectedEtag?: string,
  ) => void;
  onDiscardPendingSave: () => void;
  onCreateNew: () => Promise<boolean>;
  onBeforeShare: (
    document: CloudDiagramDocument,
    versions: CloudDiagramVersion[],
  ) => Promise<boolean>;
  currentUserEmail?: string;
  onLocateReviewAnchor: (anchor: CloudCommentAnchor) => void;
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
  hasLocalDraft,
  isCloseBlocked = false,
  onOpenDocument,
  onRestoreVersion,
  onDocumentUpdated,
  onResetCurrent,
  onSaveCurrent,
  onReloadRemote,
  onSaveAsCopy,
  onSaveAsDetachedCopy,
  onCloudConflict,
  onDiscardPendingSave,
  onCreateNew,
  onBeforeShare,
  currentUserEmail,
  onLocateReviewAnchor,
}) => {
  const { language, t } = useLanguage();
  const text = useCallback(
    (en: string, ja: string) => localize(language, { en, ja }),
    [language],
  );
  const [documents, setDocuments] = useState<CloudDiagramSummary[]>([]);
  const [selectedDocument, setSelectedDocument] = useState<CloudDiagramDocument | null>(null);
  const [versions, setVersions] = useState<CloudDiagramVersionSummary[]>([]);
  const [shares, setShares] = useState<CloudDiagramShare[]>([]);
  const [isLoading, setIsLoading] = useState(false);
  const [isDetailsLoading, setIsDetailsLoading] = useState(false);
  const [operation, setOperation] = useState('');
  const [error, setError] = useState('');
  const [comment, setComment] = useState('');
  const [commentAnchorKey, setCommentAnchorKey] = useState('canvas');
  const [showResolvedComments, setShowResolvedComments] = useState(false);
  const [reviewNote, setReviewNote] = useState('');
  const [shareRole, setShareRole] = useState<'viewer' | 'editor'>('viewer');
  const [newShareUrl, setNewShareUrl] = useState('');
  const newShareDocumentIdRef = useRef<string | null>(null);
  const isOpenRef = useRef(isOpen);
  const currentDocumentIdRef = useRef(currentDocument?.id ?? null);
  const selectedDocumentIdRef = useRef<string | null>(null);
  const replacementNoticeRef = useRef('');
  const loadGenerationRef = useRef(new OperationGeneration());
  const operationGenerationRef = useRef(new OperationGeneration());
  const useCompactDrawer = useMediaQuery(MEDIA_QUERIES.narrow);

  isOpenRef.current = isOpen;
  currentDocumentIdRef.current = currentDocument?.id ?? null;

  const closeModal = useCallback(() => {
    if (isCloseBlocked) return;
    loadGenerationRef.current.advance();
    operationGenerationRef.current.advance();
    selectedDocumentIdRef.current = null;
    replacementNoticeRef.current = '';
    setOperation('');
    setIsLoading(false);
    setIsDetailsLoading(false);
    newShareDocumentIdRef.current = null;
    setNewShareUrl('');
    onClose();
  }, [isCloseBlocked, onClose]);

  const isCurrentLoad = useCallback((generation: number) => (
    isOpenRef.current && loadGenerationRef.current.isCurrent(generation)
  ), []);

  const isCurrentOperation = useCallback((
    generation: number,
    documentId?: string,
  ) => (
    isOpenRef.current
    && operationGenerationRef.current.isCurrent(generation)
    && (!documentId || selectedDocumentIdRef.current === documentId)
  ), []);

  const selectDocument = useCallback((document: CloudDiagramDocument | null) => {
    selectedDocumentIdRef.current = document?.id ?? null;
    setSelectedDocument(document);
  }, []);

  const beginOperation = useCallback((name: string) => {
    const generation = operationGenerationRef.current.advance();
    replacementNoticeRef.current = '';
    setOperation(name);
    setError('');
    return generation;
  }, []);

  const saveCurrentBeforeMetadata = useCallback(async (documentId: string) => {
    if (
      currentDocument?.id !== documentId
      || currentContext?.role === 'viewer'
    ) return null;
    return onSaveCurrent();
  }, [currentContext?.role, currentDocument?.id, onSaveCurrent]);

  const reportCurrentConflict = useCallback((
    operationError: unknown,
    targetDocument: CloudDiagramDocument,
  ) => {
    if (
      operationError instanceof CloudDiagramApiError
      && (operationError.status === 404 || operationError.status === 412)
    ) {
      onCloudConflict(
        targetDocument.id,
        operationError,
        targetDocument.revision,
        targetDocument.etag,
      );
    }
  }, [onCloudConflict]);

  const preserveCurrentDraft = useCallback(async () => {
    const savedDocument = await (
      currentDocument && currentContext?.role === 'viewer'
      ? onSaveAsDetachedCopy()
      : onSaveCurrent({ force: true })
    );
    if (!savedDocument && hasLocalDraft) {
      throw new Error(text(
        'The current draft has not been saved to the cloud.',
        '現在の下書きはクラウドに保存されていません。',
      ));
    }
    return savedDocument;
  }, [
    currentContext?.role,
    currentDocument,
    hasLocalDraft,
    onSaveAsDetachedCopy,
    onSaveCurrent,
    text,
  ]);

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

  const selectedReview: CloudDiagramReview = selectedDocument?.review || { status: 'draft' };
  const commentAnchors = useMemo(() => {
    if (!selectedDocument) return [{ key: 'canvas', anchor: { type: 'canvas' as const }, label: text('Whole canvas', 'キャンバス全体') }];
    const nodeLabels = new Map<string, string>();
    const nodeAnchors = selectedDocument.payload.nodes
      .filter(node => node && typeof node.id === 'string')
      .map(node => {
        const label = String(node.data?.label || node.data?.serviceName || node.id);
        nodeLabels.set(node.id, label);
        return {
          key: `node\u0000${node.id}`,
          anchor: { type: 'node' as const, targetId: node.id, label },
          label,
        };
      });
    const edgeAnchors = selectedDocument.payload.edges
      .filter(edge => edge && typeof edge.id === 'string')
      .map(edge => {
        const source = nodeLabels.get(edge.source) || edge.source;
        const target = nodeLabels.get(edge.target) || edge.target;
        const label = String(edge.label || edge.data?.label || `${source} → ${target}`);
        return {
          key: `edge\u0000${edge.id}`,
          anchor: { type: 'edge' as const, targetId: edge.id, label },
          label,
        };
      });
    return [
      { key: 'canvas', anchor: { type: 'canvas' as const }, label: text('Whole canvas', 'キャンバス全体') },
      ...nodeAnchors,
      ...edgeAnchors,
    ];
  }, [selectedDocument, text]);

  const applyUpdatedMetadataDocument = useCallback((updated: CloudDiagramDocument) => {
    onDocumentUpdated(updated);
    if (!isOpenRef.current || selectedDocumentIdRef.current !== updated.id) return;
    setDocuments(items => items.map(item => (
      item.id === updated.id
        ? {
            ...item,
            updatedAt: updated.updatedAt,
            revision: updated.revision,
            commentCount: updated.comments.length,
            openCommentCount: updated.comments.filter(item => !item.resolved).length,
            reviewStatus: updated.review?.status || 'draft',
          }
        : item
    )));
    selectDocument(updated);
  }, [onDocumentUpdated, selectDocument]);

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
    generation: number,
  ) => {
    if (!isCurrentLoad(generation)) return;
    selectDocument(document);
    setIsDetailsLoading(true);
    setError(replacementNoticeRef.current);
    setVersions([]);
    setShares([]);
    setComment('');
    setCommentAnchorKey('canvas');
    setShowResolvedComments(false);
    setReviewNote('');
    if (newShareDocumentIdRef.current !== document.id) {
      newShareDocumentIdRef.current = null;
      setNewShareUrl('');
    }
    try {
      const [nextVersions, nextShares] = await Promise.all([
        listCloudVersions(context),
        context.access === 'owner' ? listCloudShares(document.id) : Promise.resolve([]),
      ]);
      if (!isCurrentLoad(generation)) return;
      setVersions(nextVersions);
      setShares(nextShares);
    } catch (loadError) {
      if (!isCurrentLoad(generation)) return;
      reportCurrentConflict(loadError, document);
      setVersions([]);
      setShares([]);
      setError(loadError instanceof Error ? loadError.message : text(
        'Failed to load cloud details.',
        'クラウドの詳細を読み込めませんでした。',
      ));
    } finally {
      if (isCurrentLoad(generation)) setIsDetailsLoading(false);
    }
  }, [isCurrentLoad, reportCurrentConflict, selectDocument, text]);

  const refreshDocuments = useCallback(async (
    preferredId?: string,
    preferredDocument?: CloudDiagramDocument,
  ) => {
    const generation = loadGenerationRef.current.advance();
    let documentListLoaded = false;
    setIsLoading(true);
    setError(replacementNoticeRef.current);
    try {
      const owned = await listCloudDiagrams();
      if (!isCurrentLoad(generation)) return;
      const merged = [...owned];
      const mergeLocalDocument = (
        document: CloudDiagramDocument | null | undefined,
        context?: CloudDocumentContext | null,
      ) => {
        if (!document || merged.some((item) => item.id === document.id)) return;
        merged.unshift({
          id: document.id,
          diagramName: document.diagramName,
          createdAt: document.createdAt,
          updatedAt: document.updatedAt,
          revision: document.revision,
          serviceCount: document.payload.nodes.length,
          connectionCount: document.payload.edges.length,
          commentCount: document.comments.length,
          openCommentCount: document.comments.filter(item => !item.resolved).length,
          reviewStatus: document.review?.status || 'draft',
          shareCount: document.shares?.length || 0,
          access: context?.access || document.access,
          role: context?.role || document.role,
          etag: document.etag,
        });
      };
      mergeLocalDocument(currentDocument, currentContext);
      mergeLocalDocument(
        preferredDocument,
        preferredDocument ? ownerContext(preferredDocument.id) : null,
      );
      if (preferredDocument) {
        const preferredIndex = merged.findIndex((item) => item.id === preferredDocument.id);
        if (preferredIndex > 0) {
          const [preferredSummary] = merged.splice(preferredIndex, 1);
          merged.unshift(preferredSummary);
        }
      }
      setDocuments(merged);
      documentListLoaded = true;

      const requestedTargetId = preferredId
        || selectedDocumentIdRef.current
        || currentDocumentIdRef.current;
      const targetId = requestedTargetId
        && merged.some((item) => item.id === requestedTargetId)
        ? requestedTargetId
        : merged[0]?.id;
      if (!targetId) {
        selectDocument(null);
        setVersions([]);
        setShares([]);
        return;
      }

      if (preferredDocument?.id === targetId) {
        await loadDetails(
          preferredDocument,
          ownerContext(preferredDocument.id),
          generation,
        );
      } else if (currentDocument?.id === targetId) {
        await loadDetails(
          currentDocument,
          currentContext || ownerContext(currentDocument.id),
          generation,
        );
      } else {
        const nextDocument = await getCloudDiagram(targetId);
        if (!isCurrentLoad(generation)) return;
        await loadDetails(nextDocument, ownerContext(targetId), generation);
      }
    } catch (loadError) {
      if (!isCurrentLoad(generation)) return;
      if (!documentListLoaded) {
        setDocuments([]);
        selectDocument(currentDocument);
      } else {
        selectDocument(null);
        setVersions([]);
        setShares([]);
      }
      setError(loadError instanceof Error ? loadError.message : text(
        'Failed to load cloud diagrams.',
        'クラウド図面を読み込めませんでした。',
      ));
    } finally {
      if (isCurrentLoad(generation)) setIsLoading(false);
    }
  }, [
    currentContext,
    currentDocument,
    isCurrentLoad,
    loadDetails,
    selectDocument,
    text,
  ]);

  useEffect(() => {
    if (isOpen) void refreshDocuments();
  }, [isOpen, refreshDocuments]);

  useEffect(() => {
    if (isOpen) return;
    loadGenerationRef.current.advance();
    operationGenerationRef.current.advance();
    selectedDocumentIdRef.current = null;
    replacementNoticeRef.current = '';
    setOperation('');
    setIsLoading(false);
    setIsDetailsLoading(false);
    newShareDocumentIdRef.current = null;
    setNewShareUrl('');
  }, [isOpen]);

  const handleSelect = async (summary: CloudDiagramSummary) => {
    if (selectedDocumentIdRef.current === summary.id) return;
    const generation = loadGenerationRef.current.advance();
    operationGenerationRef.current.advance();
    replacementNoticeRef.current = '';
    setOperation('');
    selectedDocumentIdRef.current = summary.id;
    setSelectedDocument(null);
    setIsDetailsLoading(true);
    setError('');
    setVersions([]);
    setShares([]);
    setComment('');
    newShareDocumentIdRef.current = null;
    setNewShareUrl('');
    try {
      if (currentDocument?.id === summary.id) {
        await loadDetails(
          currentDocument,
          currentContext || ownerContext(currentDocument.id),
          generation,
        );
      } else {
        const nextDocument = await getCloudDiagram(summary.id);
        if (!isCurrentLoad(generation)) return;
        await loadDetails(nextDocument, ownerContext(summary.id), generation);
      }
    } catch (loadError) {
      if (!isCurrentLoad(generation)) return;
      setError(loadError instanceof Error ? loadError.message : text(
        'Failed to open the cloud diagram.',
        'クラウド図面を開けませんでした。',
      ));
    } finally {
      if (isCurrentLoad(generation)) setIsDetailsLoading(false);
    }
  };

  const handleReplacementDocument = useCallback(async (
    savedDocument: CloudDiagramDocument | null,
    targetDocumentId: string,
    generation: number,
  ) => {
    if (!savedDocument || savedDocument.id === targetDocumentId) return false;
    const notice = text(
      'The original cloud diagram was deleted. Your work was saved as a replacement; retry the action on the selected replacement.',
      '元のクラウド図面は削除されました。作業は置換図面として保存されています。選択された置換図面で操作を再実行してください。',
    );
    replacementNoticeRef.current = notice;
    selectedDocumentIdRef.current = savedDocument.id;
    selectDocument(savedDocument);
    setVersions([]);
    setShares(savedDocument.shares || []);
    setError(notice);
    await refreshDocuments(savedDocument.id, savedDocument);
    if (isCurrentOperation(generation)) {
      setError(notice);
    }
    return true;
  }, [isCurrentOperation, refreshDocuments, selectDocument, text]);

  const handleOpenSelected = async () => {
    if (!selectedDocument || !selectedContext || syncStatus === 'conflict') return;
    const targetDocument = selectedDocument;
    const targetContext = selectedContext;
    const targetWasCurrent = currentDocument?.id === targetDocument.id;
    const currentWasViewer = currentContext?.role === 'viewer';
    const generation = beginOperation('open');
    let discardUnsavedChanges = false;

    try {
      try {
        await preserveCurrentDraft();
      } catch (saveError) {
        if (!isCurrentOperation(generation, targetDocument.id)) return;
        discardUnsavedChanges = window.confirm(text(
          'The current work could not be saved. Open the selected cloud diagram anyway and discard the unsaved changes?',
          '現在の作業を保存できませんでした。未保存の変更を破棄して、選択したクラウド図面を開きますか？',
        ));
        if (!discardUnsavedChanges) {
          setError(saveError instanceof Error ? saveError.message : text(
            'The current work could not be saved.',
            '現在の作業を保存できませんでした。',
          ));
          return;
        }
        onDiscardPendingSave();
      }

      if (!isCurrentOperation(generation, targetDocument.id)) return;
      if (targetWasCurrent && discardUnsavedChanges) {
        const reloaded = await onReloadRemote();
        if (!isCurrentOperation(generation, targetDocument.id)) return;
        if (!reloaded) {
          setError(text(
            'The latest cloud copy could not be loaded. Your local work is still preserved.',
            '最新のクラウド版を読み込めませんでした。ローカルの作業は保持されています。',
          ));
          return;
        }
        closeModal();
        return;
      }
      if (targetWasCurrent && !currentWasViewer && !discardUnsavedChanges) {
        closeModal();
        return;
      }
      onOpenDocument(targetDocument, targetContext);
      closeModal();
    } finally {
      if (isCurrentOperation(generation)) setOperation('');
    }
  };

  const handleDelete = async () => {
    if (!selectedDocument || selectedContext?.access !== 'owner') return;
    const targetDocument = selectedDocument;
    if (!window.confirm(text(
      `Delete "${targetDocument.diagramName}" and all of its cloud snapshots?`,
      `「${targetDocument.diagramName}」とクラウドスナップショットをすべて削除しますか？`,
    ))) return;

    const generation = beginOperation('delete');
    let conflictBaseline = targetDocument;
    try {
      const savedDocument = await saveCurrentBeforeMetadata(targetDocument.id);
      if (!isCurrentOperation(generation, targetDocument.id)) return;
      if (await handleReplacementDocument(savedDocument, targetDocument.id, generation)) return;
      const documentToDelete = savedDocument || targetDocument;
      conflictBaseline = documentToDelete;
      await deleteCloudDiagram(documentToDelete.id, documentToDelete.etag);
      onResetCurrent(documentToDelete.id);
      if (isOpenRef.current) {
        setDocuments((items) => items.filter((item) => item.id !== targetDocument.id));
      }
      if (isCurrentOperation(generation, targetDocument.id)) {
        selectDocument(null);
        setVersions([]);
        setShares([]);
        await refreshDocuments();
      }
    } catch (deleteError) {
      reportCurrentConflict(deleteError, conflictBaseline);
      if (isCurrentOperation(generation, targetDocument.id)) {
        setError(deleteError instanceof Error ? deleteError.message : text(
          'Failed to delete the cloud diagram.',
          'クラウド図面を削除できませんでした。',
        ));
      }
    } finally {
      if (isCurrentOperation(generation)) setOperation('');
    }
  };

  const handleAddComment = async () => {
    const message = comment.trim();
    if (!selectedDocument || !selectedContext || !message) return;
    const anchor = commentAnchors.find(item => item.key === commentAnchorKey)?.anchor;
    const targetDocument = selectedDocument;
    const targetContext = selectedContext;
    const generation = beginOperation('comment');
    let conflictBaseline = targetDocument;
    try {
      const savedDocument = await saveCurrentBeforeMetadata(targetDocument.id);
      if (!isCurrentOperation(generation, targetDocument.id)) return;
      if (await handleReplacementDocument(savedDocument, targetDocument.id, generation)) return;
      if (savedDocument) conflictBaseline = savedDocument;
      const updated = await addCloudComment(targetContext, message, anchor);
      applyUpdatedMetadataDocument(updated);
      if (isCurrentOperation(generation, targetDocument.id)) {
        setComment('');
        setCommentAnchorKey('canvas');
      }
    } catch (commentError) {
      reportCurrentConflict(commentError, conflictBaseline);
      if (isCurrentOperation(generation, targetDocument.id)) {
        setError(commentError instanceof Error ? commentError.message : text(
          'Failed to add the comment.',
          'コメントを追加できませんでした。',
        ));
      }
    } finally {
      if (isCurrentOperation(generation)) setOperation('');
    }
  };

  const handleCommentResolution = async (commentId: string, resolved: boolean) => {
    if (!selectedDocument || !selectedContext) return;
    const targetDocument = selectedDocument;
    const targetContext = selectedContext;
    const generation = beginOperation(`comment-resolution:${commentId}`);
    let conflictBaseline = targetDocument;
    try {
      const savedDocument = await saveCurrentBeforeMetadata(targetDocument.id);
      if (!isCurrentOperation(generation, targetDocument.id)) return;
      if (await handleReplacementDocument(savedDocument, targetDocument.id, generation)) return;
      if (savedDocument) conflictBaseline = savedDocument;
      const updated = await setCloudCommentResolved(targetContext, commentId, resolved);
      applyUpdatedMetadataDocument(updated);
    } catch (resolutionError) {
      reportCurrentConflict(resolutionError, conflictBaseline);
      if (isCurrentOperation(generation, targetDocument.id)) {
        setError(resolutionError instanceof Error ? resolutionError.message : text(
          'Failed to update the comment.',
          'コメントの状態を更新できませんでした。',
        ));
      }
    } finally {
      if (isCurrentOperation(generation)) setOperation('');
    }
  };

  const handleReviewAction = async (action: CloudReviewAction) => {
    if (!selectedDocument || !selectedContext) return;
    const targetDocument = selectedDocument;
    const targetContext = selectedContext;
    const generation = beginOperation(`review:${action}`);
    let conflictBaseline = targetDocument;
    try {
      const savedDocument = await saveCurrentBeforeMetadata(targetDocument.id);
      if (!isCurrentOperation(generation, targetDocument.id)) return;
      if (await handleReplacementDocument(savedDocument, targetDocument.id, generation)) return;
      if (savedDocument) conflictBaseline = savedDocument;
      const updated = await updateCloudReview(targetContext, action, reviewNote.trim());
      applyUpdatedMetadataDocument(updated);
      if (isCurrentOperation(generation, targetDocument.id)) setReviewNote('');
    } catch (reviewError) {
      reportCurrentConflict(reviewError, conflictBaseline);
      if (isCurrentOperation(generation, targetDocument.id)) {
        setError(reviewError instanceof Error ? reviewError.message : text(
          'Failed to update the review.',
          'レビュー状態を更新できませんでした。',
        ));
      }
    } finally {
      if (isCurrentOperation(generation)) setOperation('');
    }
  };

  const downloadReviewReport = () => {
    if (!selectedDocument) return;
    const report = buildCloudReviewReport(selectedDocument, language);
    const blob = new Blob([report], { type: 'text/markdown;charset=utf-8' });
    const url = URL.createObjectURL(blob);
    const link = document.createElement('a');
    link.href = url;
    link.download = `${toFileNameSegment(selectedDocument.diagramName) || 'diagram'}-review.md`;
    link.click();
    URL.revokeObjectURL(url);
  };

  const locateCommentAnchor = (anchor: CloudCommentAnchor) => {
    if (!selectedDocument || !selectedContext) return;
    if (currentDocument?.id !== selectedDocument.id) {
      onOpenDocument(selectedDocument, selectedContext);
    }
    closeModal();
    window.setTimeout(() => onLocateReviewAnchor(anchor), 80);
  };

  const handleCreateShare = async () => {
    if (!selectedDocument || selectedContext?.access !== 'owner') return;
    const targetDocument = selectedDocument;
    const targetContext = selectedContext;
    const targetRole = shareRole;
    const generation = beginOperation('share');
    let conflictBaseline = targetDocument;
    newShareDocumentIdRef.current = null;
    setNewShareUrl('');
    try {
      const savedDocument = await saveCurrentBeforeMetadata(targetDocument.id);
      if (!isCurrentOperation(generation, targetDocument.id)) return;
      if (await handleReplacementDocument(savedDocument, targetDocument.id, generation)) return;
      if (savedDocument) conflictBaseline = savedDocument;
      const versionSummaries = await listCloudVersions(targetContext);
      if (!isCurrentOperation(generation, targetDocument.id)) return;
      setVersions(versionSummaries);
      const versionDetails: CloudDiagramVersion[] = [];
      const batchSize = 8;
      for (let index = 0; index < versionSummaries.length; index += batchSize) {
        const batch = await Promise.all(
          versionSummaries
            .slice(index, index + batchSize)
            .map(version => getCloudVersion(targetContext, version.versionId)),
        );
        if (!isCurrentOperation(generation, targetDocument.id)) return;
        versionDetails.push(...batch);
      }
      const latestDocument = await getCloudDiagram(targetDocument.id);
      if (!isCurrentOperation(generation, targetDocument.id)) return;
      const preflightBaseline = savedDocument || targetDocument;
      if (
        latestDocument.revision < preflightBaseline.revision
        || (
          latestDocument.revision === preflightBaseline.revision
          && Boolean(preflightBaseline.etag)
          && latestDocument.etag !== preflightBaseline.etag
        )
        || latestDocument.diagramName !== preflightBaseline.diagramName
        || JSON.stringify(latestDocument.payload) !== JSON.stringify(preflightBaseline.payload)
      ) {
        throw new CloudDiagramApiError(
          'The cloud diagram changed before sharing. Reload it and review the latest content.',
          412,
          'SHARE_PREFLIGHT_CONFLICT',
        );
      }
      conflictBaseline = latestDocument;
      onDocumentUpdated(latestDocument);
      selectDocument(latestDocument);
      if (!await onBeforeShare(latestDocument, versionDetails)) {
        if (operationGenerationRef.current.isCurrent(generation)) setOperation('');
        return;
      }
      if (!isCurrentOperation(generation, targetDocument.id)) return;
      const result = await createCloudShare(targetDocument.id, targetRole);
      if (isCurrentOperation(generation, targetDocument.id)) {
        setShares((items) => [result.share, ...items]);
      }
      try {
        await navigator.clipboard.writeText(result.url);
      } catch {
        // The URL remains visible for manual copying when clipboard access is denied.
      }

      try {
        const refreshed = await getCloudDiagram(targetDocument.id);
        onDocumentUpdated(refreshed);
        if (!isCurrentOperation(generation, targetDocument.id)) return;
        selectDocument(refreshed);
        if (isOpenRef.current) {
          setDocuments((items) => items.map((item) => (
            item.id === refreshed.id
              ? {
                  ...item,
                  updatedAt: refreshed.updatedAt,
                  revision: refreshed.revision,
                  shareCount: refreshed.shares?.length ?? item.shareCount + 1,
                  etag: refreshed.etag,
                }
              : item
          )));
        }
      } catch (refreshError) {
        console.error('[cloud] share created but document refresh failed:', refreshError);
        onCloudConflict(
          targetDocument.id,
          refreshError instanceof Error
            ? refreshError
            : new Error('The updated cloud revision could not be loaded.'),
          conflictBaseline.revision,
          conflictBaseline.etag,
        );
        if (isCurrentOperation(generation, targetDocument.id)) {
          setError(text(
            'The share link was created, but the diagram metadata could not be refreshed.',
            '共有リンクは作成されましたが、図面のメタデータを更新できませんでした。',
          ));
        }
      }
      if (isCurrentOperation(generation, targetDocument.id)) {
        newShareDocumentIdRef.current = targetDocument.id;
        setNewShareUrl(result.url);
      }
    } catch (shareError) {
      reportCurrentConflict(shareError, conflictBaseline);
      if (isCurrentOperation(generation, targetDocument.id)) {
        setError(shareError instanceof Error ? shareError.message : text(
          'Failed to create the share link.',
          '共有リンクを作成できませんでした。',
        ));
      }
    } finally {
      if (isCurrentOperation(generation)) setOperation('');
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
    const targetDocument = selectedDocument;
    const generation = beginOperation(`revoke-${shareId}`);
    let conflictBaseline = targetDocument;
    try {
      const savedDocument = await saveCurrentBeforeMetadata(targetDocument.id);
      if (!isCurrentOperation(generation, targetDocument.id)) return;
      if (await handleReplacementDocument(savedDocument, targetDocument.id, generation)) return;
      if (savedDocument) conflictBaseline = savedDocument;
      await revokeCloudShare(targetDocument.id, shareId);
      if (isCurrentOperation(generation, targetDocument.id)) {
        setShares((items) => items.filter((share) => share.shareId !== shareId));
      }
      if (isOpenRef.current) {
        setDocuments((items) => items.map((item) => (
          item.id === targetDocument.id
            ? { ...item, shareCount: Math.max(0, item.shareCount - 1) }
            : item
        )));
      }

      try {
        const refreshed = await getCloudDiagram(targetDocument.id);
        onDocumentUpdated(refreshed);
        if (!isCurrentOperation(generation, targetDocument.id)) return;
        selectDocument(refreshed);
        if (isOpenRef.current) {
          setDocuments((items) => items.map((item) => (
            item.id === refreshed.id
              ? {
                  ...item,
                  updatedAt: refreshed.updatedAt,
                  revision: refreshed.revision,
                  shareCount: refreshed.shares?.length ?? item.shareCount,
                  etag: refreshed.etag,
                }
              : item
          )));
        }
      } catch (refreshError) {
        console.error('[cloud] share revoked but document refresh failed:', refreshError);
        onCloudConflict(
          targetDocument.id,
          refreshError instanceof Error
            ? refreshError
            : new Error('The updated cloud revision could not be loaded.'),
          conflictBaseline.revision,
          conflictBaseline.etag,
        );
        if (isCurrentOperation(generation, targetDocument.id)) {
          setError(text(
            'The share link was revoked, but the diagram metadata could not be refreshed.',
            '共有リンクは無効化されましたが、図面のメタデータを更新できませんでした。',
          ));
        }
      }
    } catch (revokeError) {
      reportCurrentConflict(revokeError, conflictBaseline);
      if (isCurrentOperation(generation, targetDocument.id)) {
        setError(revokeError instanceof Error ? revokeError.message : text(
          'Failed to revoke the share link.',
          '共有リンクを無効化できませんでした。',
        ));
      }
    } finally {
      if (isCurrentOperation(generation)) setOperation('');
    }
  };

  const handleRestoreVersion = async (versionId: string) => {
    if (!selectedDocument || !selectedContext || syncStatus === 'conflict') return;
    const targetDocument = selectedDocument;
    const targetContext = selectedContext;
    const currentDocumentIdBeforePreserve = currentDocument?.id;
    if (!window.confirm(text(
      'Restore this cloud snapshot to the canvas?',
      'このクラウドスナップショットをキャンバスへ復元しますか？',
    ))) return;
    const generation = beginOperation(`version-${versionId}`);
    let conflictBaseline = targetDocument;
    try {
      let savedCurrentDocument: CloudDiagramDocument | null = null;
      try {
        savedCurrentDocument = await preserveCurrentDraft();
        if (
          currentContext?.role !== 'viewer'
          && currentDocumentIdBeforePreserve
          && await handleReplacementDocument(
            savedCurrentDocument,
            currentDocumentIdBeforePreserve,
            generation,
          )
        ) return;
        if (savedCurrentDocument?.id === targetDocument.id) {
          conflictBaseline = savedCurrentDocument;
        }
      } catch (saveError) {
        if (!isCurrentOperation(generation, targetDocument.id)) return;
        const discardUnsavedChanges = window.confirm(text(
          'The current work could not be saved. Restore the snapshot anyway and discard the unsaved changes?',
          '現在の作業を保存できませんでした。未保存の変更を破棄して、スナップショットを復元しますか？',
        ));
        if (!discardUnsavedChanges) {
          setError(saveError instanceof Error ? saveError.message : text(
            'The current work could not be saved.',
            '現在の作業を保存できませんでした。',
          ));
          return;
        }
        onDiscardPendingSave();
      }
      if (!isCurrentOperation(generation, targetDocument.id)) return;
      const version = await getCloudVersion(targetContext, versionId);
      if (!isCurrentOperation(generation, targetDocument.id)) return;
      if (version.diagramId !== targetDocument.id) {
        throw new Error(text(
          'The selected snapshot does not belong to this diagram.',
          '選択したスナップショットはこの図面のものではありません。',
        ));
      }
      if (targetContext.role !== 'viewer') {
        await createCloudVersion(targetContext, text(
          'Automatic backup before restoring a snapshot',
          'スナップショット復元前の自動バックアップ',
        ));
        if (!isCurrentOperation(generation, targetDocument.id)) return;
      }
      const baseDocument = (
        savedCurrentDocument?.id === targetDocument.id
          ? savedCurrentDocument
          : targetDocument
      );
      const restored = await onRestoreVersion(version, baseDocument, targetContext);
      if (!isCurrentOperation(generation, targetDocument.id)) return;
      if (!restored) {
        setError(text(
          'The cloud diagram changed while the snapshot was loading. Resolve the synchronization conflict first.',
          'スナップショットの読み込み中にクラウド図面が変更されました。先に同期競合を解決してください。',
        ));
        return;
      }
      closeModal();
    } catch (versionError) {
      if (isCurrentOperation(generation, targetDocument.id)) {
        if (versionError instanceof CloudDiagramApiError && versionError.status === 404) {
          try {
            if (targetContext.access === 'shared' && targetContext.shareToken) {
              await getSharedCloudDiagram(targetContext.shareToken);
            } else {
              await getCloudDiagram(targetDocument.id);
            }
          } catch (verificationError) {
            if (isCurrentOperation(generation, targetDocument.id)) {
              reportCurrentConflict(verificationError, conflictBaseline);
            }
          }
        } else {
          reportCurrentConflict(versionError, conflictBaseline);
        }
        if (!isCurrentOperation(generation, targetDocument.id)) return;
        setError(versionError instanceof Error ? versionError.message : text(
          'Failed to restore the cloud snapshot.',
          'クラウドスナップショットを復元できませんでした。',
        ));
      }
    } finally {
      if (isCurrentOperation(generation)) setOperation('');
    }
  };

  const handleConflictAction = async (action: 'reload' | 'copy') => {
    const generation = beginOperation(action);
    try {
      const updated = action === 'reload'
        ? await onReloadRemote()
        : await onSaveAsCopy();
      if (updated && isCurrentOperation(generation)) {
        await refreshDocuments(updated.id);
      }
    } catch (actionError) {
      if (isCurrentOperation(generation)) {
        setError(actionError instanceof Error ? actionError.message : text(
          'The synchronization action failed.',
          '同期処理に失敗しました。',
        ));
      }
    } finally {
      if (isCurrentOperation(generation)) setOperation('');
    }
  };

  const handleCreateNew = async () => {
    const generation = beginOperation('new');
    try {
      if (await onCreateNew() && isCurrentOperation(generation)) closeModal();
    } catch (createError) {
      if (isCurrentOperation(generation)) {
        setError(createError instanceof Error ? createError.message : text(
          'The new diagram could not be created.',
          '新しい図面を作成できませんでした。',
        ));
      }
    } finally {
      if (isCurrentOperation(generation)) setOperation('');
    }
  };

  const reviewStatusLabel = {
    draft: text('Draft', '下書き'),
    in_review: text('In review', 'レビュー中'),
    changes_requested: text('Changes requested', '変更依頼'),
    approved: text('Approved', '承認済み'),
  }[selectedReview.status];
  const visibleComments = selectedDocument?.comments.filter(
    item => showResolvedComments || !item.resolved,
  ) || [];
  const openCommentCount = selectedDocument?.comments.filter(item => !item.resolved).length || 0;
  const normalizedCurrentUserEmail = currentUserEmail?.trim().toLowerCase() || '';

  return (
    <ResponsiveDrawer
        isOpen={isOpen}
        modal
        placement={useCompactDrawer ? 'bottom' : 'center'}
        className="modal-content cloud-workspace-modal"
        backdropClassName="cloud-workspace-overlay"
        ariaLabel={text('Cloud workspace', 'クラウド ワークスペース')}
        onClose={closeModal}
        backgroundSelectors={[
          '.app > .app-header',
          '.app > .workspace',
          '.app > .arch-chat-panel',
        ]}
      >
        <div className="modal-header">
          <h2>
            <Cloud size={24} />
            {text('Cloud workspace', 'クラウド ワークスペース')}
          </h2>
          <button className="modal-close" onClick={closeModal} title={t('Close')} aria-label={t('Close')}>
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

        {error && (
          <div className="cloud-workspace-error azd-callout azd-callout--danger" role="alert">
            {error}
          </div>
        )}

        <div className="cloud-workspace-body">
          <aside className="cloud-document-sidebar">
            <div className="cloud-section-heading">
              <span>{text('Cloud diagrams', 'クラウド図面')}</span>
              <button
                onClick={() => void refreshDocuments(selectedDocumentIdRef.current ?? undefined)}
                title={text('Refresh cloud diagrams', 'クラウド図面を更新')}
                aria-label={text('Refresh cloud diagrams', 'クラウド図面を更新')}
                disabled={isLoading || Boolean(operation)}
              >
                <RefreshCw size={16} className={isLoading ? 'spin' : ''} />
              </button>
            </div>

            <button
              type="button"
              className="cloud-new-diagram-button"
              onClick={() => void handleCreateNew()}
              disabled={Boolean(operation)}
            >
              <Plus size={16} />
              {text('New diagram', '新しい図面')}
            </button>

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
              isDetailsLoading ? (
                <div className="cloud-empty-state large"><Loader size={28} className="spin" /></div>
              ) : (
              <div className="cloud-empty-state large">
                <FolderOpen size={42} />
                <p>{text('Select a cloud diagram to view its details.', '詳細を表示するクラウド図面を選択してください。')}</p>
              </div>
              )
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
                      disabled={Boolean(operation) || syncStatus === 'conflict'}
                      onClick={() => void handleOpenSelected()}
                      title={syncStatus === 'conflict'
                        ? text(
                            'Resolve the synchronization conflict before opening a cloud diagram.',
                            'クラウド図面を開く前に同期競合を解決してください。',
                          )
                        : undefined}
                    >
                      <FolderOpen size={16} />
                      {text('Open', '開く')}
                    </button>
                    {selectedContext?.access === 'owner' && (
                      <button
                        className="btn-secondary danger"
                        onClick={() => void handleDelete()}
                        disabled={Boolean(operation)}
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
                                disabled={Boolean(operation) || syncStatus === 'conflict'}
                                title={syncStatus === 'conflict'
                                  ? text(
                                      'Resolve the synchronization conflict before restoring a snapshot.',
                                      'スナップショットを復元する前に同期競合を解決してください。',
                                    )
                                  : undefined}
                              >
                                {text('Restore', '復元')}
                              </button>
                            </div>
                          ))}
                        </div>
                      )}
                    </section>

                    <section className="cloud-detail-card cloud-review-card">
                      <div className="cloud-card-title">
                        <CheckCircle2 size={18} />
                        <h4>{text('Review and approval', 'レビューと承認')}</h4>
                        <span className={`cloud-review-status ${selectedReview.status}`}>
                          {reviewStatusLabel}
                        </span>
                      </div>
                      <div className="cloud-review-summary">
                        <div>
                          <strong>{text(
                            `${openCommentCount} open review comment${openCommentCount === 1 ? '' : 's'}`,
                            `未解決のレビューコメント ${openCommentCount} 件`,
                          )}</strong>
                          <p>{text(
                            'Approval applies to the current cloud revision and returns to Draft when approved content changes.',
                            '承認は現在のクラウド版に適用され、承認後に内容が変わると下書きへ戻ります。',
                          )}</p>
                        </div>
                        {(selectedReview.requestedAt || selectedReview.decidedAt) && (
                          <dl>
                            {selectedReview.requestedAt && (
                              <>
                                <dt>{text('Requested', '依頼')}</dt>
                                <dd>
                                  {formatDate(selectedReview.requestedAt)}
                                  {selectedReview.requestedByEmail
                                    ? ` · ${selectedReview.requestedByEmail}`
                                    : ''}
                                </dd>
                              </>
                            )}
                            {selectedReview.decidedAt && (
                              <>
                                <dt>{text('Decision', '判定')}</dt>
                                <dd>
                                  {formatDate(selectedReview.decidedAt)}
                                  {selectedReview.decidedByEmail
                                    ? ` · ${selectedReview.decidedByEmail}`
                                    : ''}
                                </dd>
                              </>
                            )}
                          </dl>
                        )}
                        {selectedReview.decisionNote && (
                          <blockquote>{selectedReview.decisionNote}</blockquote>
                        )}
                      </div>

                      {selectedContext?.access === 'shared'
                        && selectedReview.status === 'in_review' && (
                          <label className="cloud-review-note">
                            <span>{text('Decision note (optional)', '判定メモ（任意）')}</span>
                            <textarea
                              value={reviewNote}
                              onChange={event => setReviewNote(event.target.value)}
                              maxLength={1000}
                              rows={2}
                              placeholder={text(
                                'Summarize approval evidence or required changes...',
                                '承認根拠または必要な変更を要約...',
                              )}
                            />
                          </label>
                        )}

                      <div className="cloud-review-actions">
                        {selectedContext?.access === 'owner' && (
                          selectedReview.status === 'in_review' ? (
                            <button
                              type="button"
                              className="cloud-review-secondary"
                              onClick={() => void handleReviewAction('cancel')}
                              disabled={Boolean(operation)}
                            >
                              <RotateCcw size={15} />
                              {text('Cancel review', 'レビューを取り消す')}
                            </button>
                          ) : (
                            <button
                              type="button"
                              className="cloud-review-primary"
                              onClick={() => void handleReviewAction('request')}
                              disabled={Boolean(operation)}
                            >
                              <Send size={15} />
                              {selectedReview.status === 'changes_requested'
                                ? text('Request review again', '再レビューを依頼')
                                : text('Request review', 'レビューを依頼')}
                            </button>
                          )
                        )}
                        {selectedContext?.access === 'shared'
                          && selectedReview.status === 'in_review' && (
                            <>
                              <button
                                type="button"
                                className="cloud-review-secondary danger"
                                onClick={() => void handleReviewAction('request_changes')}
                                disabled={Boolean(operation)}
                              >
                                <Edit3 size={15} />
                                {text('Request changes', '変更を依頼')}
                              </button>
                              <button
                                type="button"
                                className="cloud-review-primary approve"
                                onClick={() => void handleReviewAction('approve')}
                                disabled={Boolean(operation)}
                              >
                                <ThumbsUp size={15} />
                                {text('Approve revision', 'この版を承認')}
                              </button>
                            </>
                          )}
                        <button
                          type="button"
                          className="cloud-review-secondary report"
                          onClick={downloadReviewReport}
                        >
                          <Download size={15} />
                          {text('Download report', 'レポートをダウンロード')}
                        </button>
                      </div>
                    </section>

                    <section className="cloud-detail-card cloud-comments-card">
                      <div className="cloud-card-title">
                        <MessageSquare size={18} />
                        <h4>{text('Anchored comments', 'アンカー付きコメント')}</h4>
                        <span>{openCommentCount}/{selectedDocument.comments.length}</span>
                      </div>
                      {selectedDocument.comments.some(item => item.resolved) && (
                        <button
                          type="button"
                          className="cloud-resolved-toggle"
                          onClick={() => setShowResolvedComments(value => !value)}
                        >
                          {showResolvedComments
                            ? text('Hide resolved comments', '解決済みを非表示')
                            : text('Show resolved comments', '解決済みを表示')}
                        </button>
                      )}
                      <div className="cloud-comment-list">
                        {visibleComments.length === 0 ? (
                          <p className="cloud-card-empty">{text('No comments yet.', 'コメントはまだありません。')}</p>
                        ) : visibleComments.map((item) => {
                          const canResolve = selectedContext?.role !== 'viewer'
                            || (
                              normalizedCurrentUserEmail
                              && item.authorEmail.toLowerCase() === normalizedCurrentUserEmail
                            );
                          return (
                          <article key={item.commentId} className={item.resolved ? 'resolved' : ''}>
                            <header>
                              <div>
                                <strong>{item.authorEmail}</strong>
                                <time>{formatDate(item.createdAt)}</time>
                              </div>
                              <span className={item.resolved ? 'resolved' : 'open'}>
                                {item.resolved ? text('Resolved', '解決済み') : text('Open', '未解決')}
                              </span>
                            </header>
                            <p>{item.message}</p>
                            <footer>
                              {item.anchor && (
                                <button
                                  type="button"
                                  onClick={() => locateCommentAnchor(item.anchor!)}
                                >
                                  <MapPin size={13} />
                                  {item.anchor.type === 'canvas'
                                    ? text('Whole canvas', 'キャンバス全体')
                                    : item.anchor.label || text('Locate item', '対象を表示')}
                                </button>
                              )}
                              {canResolve && (
                                <button
                                  type="button"
                                  onClick={() => void handleCommentResolution(
                                    item.commentId,
                                    !item.resolved,
                                  )}
                                  disabled={Boolean(operation)}
                                >
                                  {item.resolved
                                    ? text('Reopen', '再開')
                                    : text('Resolve', '解決')}
                                </button>
                              )}
                            </footer>
                          </article>
                        )})}
                      </div>
                      <div className="cloud-comment-composer">
                        <label>
                          <span>{text('Attach to', '対象')}</span>
                          <select
                            value={commentAnchorKey}
                            onChange={event => setCommentAnchorKey(event.target.value)}
                          >
                            {commentAnchors.map(item => (
                              <option key={item.key} value={item.key}>{item.label}</option>
                            ))}
                          </select>
                        </label>
                        <div>
                          <textarea
                            value={comment}
                            onChange={(event) => setComment(event.target.value)}
                            maxLength={2000}
                            rows={2}
                            placeholder={text('Add a review comment...', 'レビューコメントを追加...')}
                          />
                          <button
                            onClick={() => void handleAddComment()}
                            disabled={!comment.trim() || Boolean(operation)}
                          >
                            {text('Comment', 'コメント')}
                          </button>
                        </div>
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
                              disabled={Boolean(operation)}
                            >
                              <option value="viewer">{text('Viewer', '閲覧者')}</option>
                              <option value="editor">{text('Editor', '編集者')}</option>
                            </select>
                          </label>
                          <button onClick={() => void handleCreateShare()} disabled={Boolean(operation)}>
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
                                disabled={Boolean(operation)}
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
          <button className="btn-secondary" onClick={closeModal}>{t('Close')}</button>
        </div>
    </ResponsiveDrawer>
  );
};

export default CloudWorkspaceModal;
