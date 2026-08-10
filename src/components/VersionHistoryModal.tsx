// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

import React, { useCallback, useEffect, useMemo, useRef, useState } from 'react';
import { X, Clock, ExternalLink, Trash2, Copy, GitCompare, CheckSquare, Square } from 'lucide-react';
import type { Edge, Node } from 'reactflow';
import { DiagramVersion, getAllVersions, deleteVersion, getVersion } from '../services/versionStorageService';
import './VersionHistoryModal.css';
import { useLanguage } from '../i18n/LanguageContext';
import { localize } from '../i18n/localization';
import { useEscapeKey } from '../hooks/useEscapeKey';
import { useModalFocus } from '../hooks/useModalFocus';
import { OperationGeneration } from '../utils/operationGeneration';
import {
  compareDiagramVersions,
  type VersionDiff,
  type VersionDiffItem,
} from '../utils/versionDiff';

interface VersionHistoryModalProps {
  isOpen: boolean;
  onClose: () => void;
  onRestoreVersion: (version: DiagramVersion, restoreAsCopy: boolean) => void;
  onRestoreSelection: (version: DiagramVersion, selectedKeys: string[]) => Promise<boolean>;
  currentLineageId: string;
  currentNodes: Node[];
  currentEdges: Edge[];
}

interface VersionDiffPreviewProps {
  currentNodes: Node[];
  currentEdges: Edge[];
  targetNodes: Node[];
  targetEdges: Edge[];
  diff: VersionDiff;
  ariaLabel: string;
}

function VersionDiffPreview({
  currentNodes,
  currentEdges,
  targetNodes,
  targetEdges,
  diff,
  ariaLabel,
}: VersionDiffPreviewProps) {
  const preview = useMemo(() => {
    const diffByNodeId = new Map(
      diff.nodes.map(item => [item.id, item.status]),
    );
    const diffByEdgeId = new Map(
      diff.edges.map(item => [item.id, item.status]),
    );
    const currentNodeMap = new Map(currentNodes.map(node => [node.id, node]));
    const targetNodeMap = new Map(targetNodes.map(node => [node.id, node]));
    const nodeIds = [...new Set([...currentNodeMap.keys(), ...targetNodeMap.keys()])];
    const nodes = nodeIds.map((id) => {
      const status = diffByNodeId.get(id);
      const node = status === 'removed'
        ? currentNodeMap.get(id)
        : targetNodeMap.get(id) || currentNodeMap.get(id);
      return node ? { node, status } : null;
    }).filter((entry): entry is {
      node: Node;
      status: VersionDiffItem['status'] | undefined;
    } => entry !== null);

    const minX = Math.min(...nodes.map(entry => entry.node.position.x), 0);
    const minY = Math.min(...nodes.map(entry => entry.node.position.y), 0);
    const maxX = Math.max(...nodes.map(entry => entry.node.position.x), minX + 1);
    const maxY = Math.max(...nodes.map(entry => entry.node.position.y), minY + 1);
    const scaleX = 430 / Math.max(1, maxX - minX + 120);
    const scaleY = 130 / Math.max(1, maxY - minY + 80);
    const scale = Math.min(scaleX, scaleY, 1.6);
    const positions = new Map(nodes.map(({ node }) => [
      node.id,
      {
        x: 35 + (node.position.x - minX) * scale,
        y: 25 + (node.position.y - minY) * scale,
      },
    ]));

    const currentEdgeMap = new Map(currentEdges.map(edge => [edge.id, edge]));
    const targetEdgeMap = new Map(targetEdges.map(edge => [edge.id, edge]));
    const edgeIds = [...new Set([...currentEdgeMap.keys(), ...targetEdgeMap.keys()])];
    const edges = edgeIds.map((id) => {
      const status = diffByEdgeId.get(id);
      const edge = status === 'removed'
        ? currentEdgeMap.get(id)
        : targetEdgeMap.get(id) || currentEdgeMap.get(id);
      return edge ? { edge, status } : null;
    }).filter((entry): entry is {
      edge: Edge;
      status: VersionDiffItem['status'] | undefined;
    } => entry !== null);

    return { nodes, edges, positions };
  }, [currentEdges, currentNodes, diff.edges, diff.nodes, targetEdges, targetNodes]);

  return (
    <svg
      className="version-diff-preview"
      viewBox="0 0 500 180"
      role="img"
      aria-label={ariaLabel}
    >
      {preview.edges.map(({ edge, status }) => {
        const source = preview.positions.get(edge.source);
        const target = preview.positions.get(edge.target);
        if (!source || !target) return null;
        return (
          <line
            key={edge.id}
            className={`version-preview-edge version-preview-edge--${status || 'unchanged'}`}
            x1={source.x + 34}
            y1={source.y + 16}
            x2={target.x + 34}
            y2={target.y + 16}
          />
        );
      })}
      {preview.nodes.map(({ node, status }) => {
        const position = preview.positions.get(node.id);
        if (!position) return null;
        const label = String(node.data?.label || node.data?.serviceName || node.id);
        return (
          <g
            key={node.id}
            className={`version-preview-node version-preview-node--${status || 'unchanged'}`}
            transform={`translate(${position.x} ${position.y})`}
          >
            <rect width="68" height="32" rx="7" />
            <text x="34" y="20" textAnchor="middle">
              {label.length > 12 ? `${label.slice(0, 11)}…` : label}
            </text>
          </g>
        );
      })}
    </svg>
  );
}

const VersionHistoryModal: React.FC<VersionHistoryModalProps> = ({
  isOpen,
  onClose,
  onRestoreVersion,
  onRestoreSelection,
  currentLineageId,
  currentNodes,
  currentEdges,
}) => {
  const { t, language } = useLanguage();
  const [versions, setVersions] = useState<DiagramVersion[]>([]);
  const [isLoading, setIsLoading] = useState(true);
  const [selectedVersion, setSelectedVersion] = useState<string | null>(null);
  const [selectedDiffKeys, setSelectedDiffKeys] = useState<Set<string>>(new Set());
  const [operation, setOperation] = useState('');
  const isOpenRef = useRef(isOpen);
  const loadGenerationRef = useRef(new OperationGeneration());
  const operationGenerationRef = useRef(new OperationGeneration());
  const dialogRef = useModalFocus<HTMLDivElement>(isOpen);

  isOpenRef.current = isOpen;

  const closeModal = useCallback(() => {
    loadGenerationRef.current.advance();
    operationGenerationRef.current.advance();
    setOperation('');
    setSelectedDiffKeys(new Set());
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
    setSelectedDiffKeys(new Set());
    setIsLoading(false);
  }, [isOpen, loadVersions]);

  const selectedVersionData = useMemo(
    () => versions.find(version => version.versionId === selectedVersion) || null,
    [selectedVersion, versions],
  );
  const selectedDiff = useMemo(() => (
    selectedVersionData
      ? compareDiagramVersions(
          currentNodes,
          currentEdges,
          selectedVersionData.nodes as Node[],
          selectedVersionData.edges as Edge[],
        )
      : null
  ), [currentEdges, currentNodes, selectedVersionData]);

  const selectVersion = (versionId: string) => {
    setSelectedVersion(versionId);
    setSelectedDiffKeys(new Set());
  };

  const toggleDiffKey = (key: string) => {
    setSelectedDiffKeys(current => {
      const next = new Set(current);
      if (next.has(key)) next.delete(key);
      else next.add(key);
      return next;
    });
  };

  const handleSelectiveRestore = async () => {
    if (!selectedVersionData || selectedDiffKeys.size === 0) return;
    const confirmed = window.confirm(localize(language, {
      en: `Apply ${selectedDiffKeys.size} selected historical change${selectedDiffKeys.size === 1 ? '' : 's'} to the current diagram? A backup snapshot will be created first.`,
      ja: `選択した${selectedDiffKeys.size}件の過去変更を現在の図へ適用しますか？ 先にバックアップ スナップショットを作成します。`,
    }));
    if (!confirmed) return;
    const generation = operationGenerationRef.current.advance();
    setOperation(`selective-${selectedVersionData.versionId}`);
    try {
      const restored = await onRestoreSelection(
        selectedVersionData,
        [...selectedDiffKeys],
      );
      if (!isCurrentOperation(generation)) return;
      if (restored) closeModal();
    } catch (error) {
      if (!isCurrentOperation(generation)) return;
      console.error('Failed to restore selected version changes:', error);
      alert(localize(language, {
        en: 'The selected version changes could not be restored.',
        ja: '選択したバージョン変更を復元できませんでした。',
      }));
    } finally {
      if (isCurrentOperation(generation)) setOperation('');
    }
  };

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

      const restoreAsCopy = version.lineageId !== currentLineageId;
      const confirmation = localize(language, {
        en: restoreAsCopy
          ? `Restore this version as a local copy? The current cloud diagram will remain saved.\n\nVersion: ${version.diagramName}\nCreated: ${formatDate(version.timestamp)}`
          : `Restore this version? Your current diagram will be replaced.\n\nVersion: ${version.diagramName}\nCreated: ${formatDate(version.timestamp)}`,
        ja: restoreAsCopy
          ? `このバージョンをローカルコピーとして復元しますか？ 現在のクラウド図面は保存されたまま残ります。\n\nバージョン: ${version.diagramName}\n作成日時: ${formatDate(version.timestamp)}`
          : `このバージョンを復元しますか？ 現在の図は置き換えられます。\n\nバージョン: ${version.diagramName}\n作成日時: ${formatDate(version.timestamp)}`,
      });
      if (confirm(confirmation)) {
        onRestoreVersion(version, restoreAsCopy);
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
        ref={dialogRef}
        className="modal-content version-history-modal"
        role="dialog"
        aria-modal="true"
        aria-label={t("Version History")}
        tabIndex={-1}
        onClick={(e) => e.stopPropagation()}
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
            <>
            {selectedVersionData && selectedDiff && (
              <section className="version-comparison" aria-label={localize(language, {
                en: 'Visual version comparison',
                ja: 'バージョンの視覚比較',
              })}>
                <div className="version-comparison-header">
                  <div>
                    <GitCompare size={20} aria-hidden="true" />
                    <span>
                      <h3>{localize(language, {
                        en: 'Compare with current diagram',
                        ja: '現在の図と比較',
                      })}</h3>
                      <p>{selectedVersionData.diagramName}</p>
                    </span>
                  </div>
                  <div className="version-diff-summary" role="group" aria-label={localize(language, {
                    en: 'Difference summary',
                    ja: '差分の概要',
                  })}>
                    <span className="added">+{selectedDiff.counts.added}</span>
                    <span className="removed">−{selectedDiff.counts.removed}</span>
                    <span className="changed">~{selectedDiff.counts.changed}</span>
                  </div>
                </div>

                <VersionDiffPreview
                  currentNodes={currentNodes}
                  currentEdges={currentEdges}
                  targetNodes={selectedVersionData.nodes as Node[]}
                  targetEdges={selectedVersionData.edges as Edge[]}
                  diff={selectedDiff}
                  ariaLabel={localize(language, {
                    en: 'Diagram preview showing added, removed, and changed elements',
                    ja: '追加、削除、変更されたノードを示す図プレビュー',
                  })}
                />

                {selectedDiff.items.length === 0 ? (
                  <p className="version-no-differences">{localize(language, {
                    en: 'This version has the same diagram elements as the current canvas.',
                    ja: 'このバージョンのノードは現在のキャンバスと同じです。',
                  })}</p>
                ) : (
                  <>
                    <div className="version-diff-controls">
                      <button
                        type="button"
                        onClick={() => setSelectedDiffKeys(new Set(selectedDiff.items.map(item => item.key)))}
                        disabled={Boolean(operation)}
                      >
                        <CheckSquare size={15} />
                        {localize(language, { en: 'Select all', ja: 'すべて選択' })}
                      </button>
                      <button
                        type="button"
                        onClick={() => setSelectedDiffKeys(new Set())}
                        disabled={Boolean(operation) || selectedDiffKeys.size === 0}
                      >
                        <Square size={15} />
                        {localize(language, { en: 'Clear', ja: '選択解除' })}
                      </button>
                      <button
                        type="button"
                        className="version-apply-selection"
                        onClick={() => void handleSelectiveRestore()}
                        disabled={Boolean(operation) || selectedDiffKeys.size === 0}
                      >
                        <Copy size={15} />
                        {localize(language, {
                          en: `Apply selected (${selectedDiffKeys.size})`,
                          ja: `選択項目を適用 (${selectedDiffKeys.size})`,
                        })}
                      </button>
                    </div>
                    <div className="version-diff-list">
                      {selectedDiff.items.map(item => (
                        <label key={item.key} className={`version-diff-item version-diff-item--${item.status}`}>
                          <input
                            type="checkbox"
                            checked={selectedDiffKeys.has(item.key)}
                            onChange={() => toggleDiffKey(item.key)}
                            disabled={Boolean(operation)}
                          />
                          <span className="version-diff-kind">
                            {item.kind === 'node'
                              ? localize(language, { en: 'Service', ja: 'サービス' })
                              : localize(language, { en: 'Connection', ja: '接続' })}
                          </span>
                          <strong>{item.label}</strong>
                          <span className="version-diff-status">
                            {item.status === 'added'
                              ? localize(language, { en: 'Add from version', ja: 'バージョンから追加' })
                              : item.status === 'removed'
                                ? localize(language, { en: 'Remove from current', ja: '現在から削除' })
                                : localize(language, { en: 'Restore historical value', ja: '過去の値へ戻す' })}
                          </span>
                        </label>
                      ))}
                    </div>
                  </>
                )}
              </section>
            )}
            <div className="version-list">
              {versions.map((version, index) => {
                const currentLineage = version.lineageId === currentLineageId;
                return (
                  <div
                  key={version.versionId}
                  className={`version-item ${selectedVersion === version.versionId ? 'selected' : ''}`}
                  onClick={() => selectVersion(version.versionId)}
                >
                  <div className="version-header">
                    <div className="version-title">
                      <h4>{version.diagramName || t("Untitled Diagram")}</h4>
                      {index === 0 && <span className="version-badge latest">{t("Latest")}</span>}
                      <span className={`version-badge ${currentLineage ? 'current' : 'copy'}`}>
                        {localize(language, {
                          en: currentLineage ? 'Current diagram' : 'Restore as copy',
                          ja: currentLineage ? '現在の図' : 'コピーとして復元',
                        })}
                      </span>
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
                      {' '}
                      {currentLineage
                        ? t("Restore This Version")
                        : localize(language, { en: 'Restore as Copy', ja: 'コピーとして復元' })}
                      {' '}
                    </button>
                  </div>
                  </div>
                );
              })}
            </div>
            </>
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
