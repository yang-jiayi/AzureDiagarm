// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

import React, { useCallback, useEffect, useRef, useState } from 'react';
import { X, Clock, ExternalLink, Trash2, Copy } from 'lucide-react';
import { DiagramVersion, getAllVersions, deleteVersion, getVersion } from '../services/versionStorageService';
import './VersionHistoryModal.css';
import { useLanguage } from '../i18n/LanguageContext';
import { localize } from '../i18n/localization';
import { useEscapeKey } from '../hooks/useEscapeKey';
import { OperationGeneration } from '../utils/operationGeneration';

interface VersionHistoryModalProps {
  isOpen: boolean;
  onClose: () => void;
  onRestoreVersion: (version: DiagramVersion) => void;
  currentDiagramName?: string; // For future filtering feature
}

const VersionHistoryModal: React.FC<VersionHistoryModalProps> = ({
  isOpen,
  onClose,
  onRestoreVersion,
  currentDiagramName: _currentDiagramName
}) => {
  const { t, language } = useLanguage();
  const [versions, setVersions] = useState<DiagramVersion[]>([]);
  const [isLoading, setIsLoading] = useState(true);
  const [selectedVersion, setSelectedVersion] = useState<string | null>(null);
  const [operation, setOperation] = useState('');
  const isOpenRef = useRef(isOpen);
  const loadGenerationRef = useRef(new OperationGeneration());
  const operationGenerationRef = useRef(new OperationGeneration());

  isOpenRef.current = isOpen;

  const closeModal = useCallback(() => {
    loadGenerationRef.current.advance();
    operationGenerationRef.current.advance();
    setOperation('');
    onClose();
  }, [onClose]);

  const isCurrentOperation = useCallback((generation: number) => (
    isOpenRef.current && operationGenerationRef.current.isCurrent(generation)
  ), []);

  const loadVersions = useCallback(async () => {
    const generation = loadGenerationRef.current.advance();
    setIsLoading(true);
    try {
      const allVersions = await getAllVersions();
      if (!isOpenRef.current || !loadGenerationRef.current.isCurrent(generation)) return;
      setVersions(allVersions);
    } catch (error) {
      if (!isOpenRef.current || !loadGenerationRef.current.isCurrent(generation)) return;
      console.error('Failed to load versions:', error);
    } finally {
      if (isOpenRef.current && loadGenerationRef.current.isCurrent(generation)) {
        setIsLoading(false);
      }
    }
  }, []);

  useEffect(() => {
    if (isOpen) {
      void loadVersions();
      return;
    }
    loadGenerationRef.current.advance();
    operationGenerationRef.current.advance();
    setOperation('');
    setIsLoading(false);
  }, [isOpen, loadVersions]);

  const handleDelete = async (versionId: string, event: React.MouseEvent) => {
    event.stopPropagation();
    
    if (!confirm(t("Are you sure you want to delete this version? This cannot be undone."))) {
      return;
    }

    const generation = operationGenerationRef.current.advance();
    setOperation(`delete-${versionId}`);
    try {
      await deleteVersion(versionId);
      if (!isCurrentOperation(generation)) return;
      await loadVersions();
    } catch (error) {
      if (!isCurrentOperation(generation)) return;
      console.error('Failed to delete version:', error);
      alert(t("Failed to delete version"));
    } finally {
      if (isCurrentOperation(generation)) setOperation('');
    }
  };

  const handleOpenInNewTab = async (versionId: string, event: React.MouseEvent) => {
    event.stopPropagation();
    
    const generation = operationGenerationRef.current.advance();
    setOperation(`open-${versionId}`);
    try {
      const version = await getVersion(versionId);
      if (!isCurrentOperation(generation)) return;
      if (!version) {
        alert(t("Version not found"));
        return;
      }

      // The version payload remains in IndexedDB, which is shared across
      // same-origin tabs. Passing only the id avoids browser URL-size limits.
      const versionUrl = `${window.location.origin}${window.location.pathname}${window.location.search}#version-id-${encodeURIComponent(version.versionId)}`;
      const newTab = window.open(versionUrl, '_blank');
      
      if (!newTab) {
        alert(t("Please allow pop-ups to open versions in new tabs"));
      } else {
        newTab.opener = null;
      }
    } catch (error) {
      if (!isCurrentOperation(generation)) return;
      console.error('Failed to open version:', error);
      alert(t("Failed to open version in new tab"));
    } finally {
      if (isCurrentOperation(generation)) setOperation('');
    }
  };

  const handleRestore = async (versionId: string) => {
    const generation = operationGenerationRef.current.advance();
    setOperation(`restore-${versionId}`);
    try {
      const version = await getVersion(versionId);
      if (!isCurrentOperation(generation)) return;
      if (!version) {
        alert(t("Version not found"));
        return;
      }

      const confirmation = localize(language, {
        en: `Restore this version? Your current diagram will be replaced.\n\nVersion: ${version.diagramName}\nCreated: ${formatDate(version.timestamp)}`,
        ja: `このバージョンを復元しますか？ 現在の図は置き換えられます。\n\nバージョン: ${version.diagramName}\n作成日時: ${formatDate(version.timestamp)}`,
      });
      if (confirm(confirmation)) {
        onRestoreVersion(version);
        closeModal();
      }
    } catch (error) {
      if (!isCurrentOperation(generation)) return;
      console.error('Failed to restore version:', error);
      alert(t("Failed to restore version"));
    } finally {
      if (isCurrentOperation(generation)) setOperation('');
    }
  };

  const formatDate = (timestamp: number): string => {
    const date = new Date(timestamp);
    return date.toLocaleString(language === 'ja' ? 'ja-JP' : 'en-US', {
      month: 'short',
      day: 'numeric',
      year: 'numeric',
      hour: '2-digit',
      minute: '2-digit'
    });
  };

  const formatTimeAgo = (timestamp: number): string => {
    const seconds = Math.floor((Date.now() - timestamp) / 1000);
    
    if (seconds < 60) return localize(language, { en: 'just now', ja: 'たった今' });
    if (seconds < 3600) {
      const minutes = Math.floor(seconds / 60);
      return localize(language, { en: `${minutes}m ago`, ja: `${minutes}分前` });
    }
    if (seconds < 86400) {
      const hours = Math.floor(seconds / 3600);
      return localize(language, { en: `${hours}h ago`, ja: `${hours}時間前` });
    }
    if (seconds < 604800) {
      const days = Math.floor(seconds / 86400);
      return localize(language, { en: `${days}d ago`, ja: `${days}日前` });
    }
    return formatDate(timestamp);
  };

  useEscapeKey(isOpen, closeModal);
  if (!isOpen) return null;

  return (
    <div className="modal-overlay" onClick={closeModal}>
      <div
        className="modal-content version-history-modal"
        role="dialog"
        aria-modal="true"
        aria-label={t("Version History")}
        tabIndex={-1}
        autoFocus
        onClick={(e) => e.stopPropagation()}
        onKeyDown={(event) => {
          if (event.key === 'Escape') closeModal();
        }}
      >
        <div className="modal-header">
          <h2>
            <Clock size={24} />
            {' '}{t("Version History")}{' '}</h2>
          <button className="modal-close" onClick={closeModal} title={t("Close")} aria-label={t("Close")}>
            <X size={24} />
          </button>
        </div>

        <div className="modal-body">
          {isLoading ? (
            <div className="version-loading">
              <div className="spinner"></div>
              <p>{t("Loading versions...")}</p>
            </div>
          ) : versions.length === 0 ? (
            <div className="version-empty">
              <Clock size={48} style={{ opacity: 0.3 }} />
              <p>{t("No versions saved yet")}</p>
              <p className="version-empty-hint">
                {' '}{t("Versions are automatically created when you regenerate architecture with AI, or you can manually create snapshots.")}{' '}</p>
            </div>
          ) : (
            <div className="version-list">
              {versions.map((version, index) => (
                <div
                  key={version.versionId}
                  className={`version-item ${selectedVersion === version.versionId ? 'selected' : ''}`}
                  onClick={() => setSelectedVersion(version.versionId)}
                >
                  <div className="version-header">
                    <div className="version-title">
                      <h4>{version.diagramName || t("Untitled Diagram")}</h4>
                      {index === 0 && <span className="version-badge latest">{t("Latest")}</span>}
                      {version.validationScore !== undefined && (
                        <span className="version-badge score" title={t("Validation Score")}>
                          {version.validationScore}{t("/100")}{' '}</span>
                      )}
                    </div>
                    <div className="version-actions">
                      <button
                        className="version-action-btn"
                        onClick={(e) => handleOpenInNewTab(version.versionId, e)}
                        title={t("Open in new tab for comparison")}
                        disabled={Boolean(operation)}
                      >
                        <ExternalLink size={16} />
                      </button>
                      <button
                        className="version-action-btn delete"
                        onClick={(e) => handleDelete(version.versionId, e)}
                        title={t("Delete this version")}
                        disabled={Boolean(operation)}
                      >
                        <Trash2 size={16} />
                      </button>
                    </div>
                  </div>

                  <div className="version-meta">
                    <span className="version-time" title={formatDate(version.timestamp)}>
                      <Clock size={14} />
                      {formatTimeAgo(version.timestamp)}
                    </span>
                    {version.nodes && (
                      <span className="version-stat">
                        {version.nodes.length} {' '}{t("services")}{' '}</span>
                    )}
                    {version.edges && (
                      <span className="version-stat">
                        {version.edges.length} {' '}{t("connections")}{' '}</span>
                    )}
                  </div>

                  {version.architecturePrompt && (
                    <div className="version-prompt">
                      <strong>{t("Prompt:")}</strong> {version.architecturePrompt.substring(0, 100)}
                      {version.architecturePrompt.length > 100 && '...'}
                    </div>
                  )}

                  {version.improvementsApplied && version.improvementsApplied.length > 0 && (
                    <div className="version-improvements">
                      <strong>{t("Improvements:")}</strong>
                      <ul>
                        {version.improvementsApplied.slice(0, 3).map((improvement, i) => (
                          <li key={i}>{improvement}</li>
                        ))}
                        {version.improvementsApplied.length > 3 && (
                          <li>{t("+")}{' '}{version.improvementsApplied.length - 3} {' '}{t("more...")}</li>
                        )}
                      </ul>
                    </div>
                  )}

                  {version.notes && (
                    <div className="version-notes">
                      <strong>{t("Notes:")}</strong> {version.notes}
                    </div>
                  )}

                  <div className="version-footer">
                    <button
                      className="btn-restore"
                      onClick={() => handleRestore(version.versionId)}
                      disabled={Boolean(operation)}
                    >
                      <Copy size={16} />
                      {' '}{t("Restore This Version")}{' '}</button>
                  </div>
                </div>
              ))}
            </div>
          )}
        </div>

        <div className="modal-actions">
          <div className="version-count">
            {localize(language, {
              en: `${versions.length} ${versions.length === 1 ? 'version' : 'versions'} saved`,
              ja: `${versions.length}件のバージョンを保存済み`,
            })}
          </div>
          <button className="btn-secondary" onClick={closeModal}>
            {' '}{t("Close")}{' '}</button>
        </div>
      </div>
    </div>
  );
};

export default VersionHistoryModal;
