// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

import { useEffect, useId, useMemo, useRef, useState } from 'react';
import { createPortal } from 'react-dom';
import { AlertCircle, Check, Loader2, X } from 'lucide-react';
import { useLanguage } from '../i18n/LanguageContext';
import { localize } from '../i18n/localization';
import { useModalFocus } from '../hooks/useModalFocus';
import {
  applyDiagramChanges, type DiagramChange, type DiagramChangeSet, type DiagramGraph,
} from '../services/diagramChanges';
import './AIChangeReview.css';

export interface AIChangeReviewProps {
  changeSet: DiagramChangeSet;
  onApply: (graph: DiagramGraph) => void | Promise<void>;
  onCancel: () => void;
  isApplying?: boolean;
  error?: string;
}

export default function AIChangeReview({ changeSet, onApply, onCancel, isApplying = false, error }: AIChangeReviewProps) {
  const { language } = useLanguage();
  const titleId = useId();
  const descriptionId = useId();
  const applyingRef = useRef(false);
  const mountedRef = useRef(true);
  const [selectedIds, setSelectedIds] = useState(() => new Set(changeSet.changes.map(change => change.id)));
  const [submitting, setSubmitting] = useState(false);
  const [applyError, setApplyError] = useState('');
  const busy = isApplying || submitting;
  const dialogRef = useModalFocus(true, onCancel, { closeOnEscape: !busy });
  useEffect(() => {
    mountedRef.current = true;
    return () => { mountedRef.current = false; };
  }, []);
  useEffect(() => {
    setSelectedIds(new Set(changeSet.changes.map(change => change.id)));
    setApplyError('');
  }, [changeSet]);

  const selection = useMemo(() => {
    try {
      return { graph: applyDiagramChanges(changeSet, selectedIds), error: '' };
    } catch (cause) {
      return { graph: null, error: cause instanceof Error ? cause.message : String(cause) };
    }
  }, [changeSet, selectedIds]);
  const selected = changeSet.changes.filter(change => selectedIds.has(change.id));
  const priced = selected.filter(change => change.costDelta !== undefined);
  const delta = priced.reduce((sum, change) => sum + change.costDelta!, 0);
  const money = (value: number) => `${value >= 0 ? '+' : '−'}${new Intl.NumberFormat(
    language === 'ja' ? 'ja-JP' : 'en-US', { style: 'currency', currency: 'USD', maximumFractionDigits: 2 },
  ).format(Math.abs(value))}`;
  const kindLabel = (kind: DiagramChange['kind']) => localize(language, {
    en: kind === 'add' ? 'Add' : kind === 'delete' ? 'Delete' : 'Change',
    ja: kind === 'add' ? '追加' : kind === 'delete' ? '削除' : '変更',
  });
  const fieldLabel = (field: string) => {
    const fields: Record<string, { en: string; ja: string }> = {
      data: { en: 'Properties', ja: 'プロパティ' }, label: { en: 'Label', ja: 'ラベル' },
      position: { en: 'Position', ja: '位置' }, style: { en: 'Appearance', ja: '表示スタイル' },
      source: { en: 'Source', ja: '接続元' }, target: { en: 'Target', ja: '接続先' },
      parentNode: { en: 'Parent group', ja: '親グループ' }, extent: { en: 'Group bounds', ja: 'グループ範囲' },
      type: { en: 'Type', ja: '種類' }, sourceHandle: { en: 'Source handle', ja: '接続元ハンドル' },
      targetHandle: { en: 'Target handle', ja: '接続先ハンドル' },
    };
    return fields[field] ? localize(language, fields[field]) : field;
  };
  const selectionError = selection.error.split('\n').map(message => {
    let match = /^Connection "(.*)" requires node "(.*)"\.$/.exec(message);
    if (match) return localize(language, { en: message, ja: `接続「${match[1]}」にはノード「${match[2]}」が必要です。` });
    match = /^Node "(.*)" requires parent "(.*)"\.$/.exec(message);
    if (match) return localize(language, { en: message, ja: `ノード「${match[1]}」には親グループ「${match[2]}」が必要です。` });
    match = /^Parent cycle involving "(.*)"\.$/.exec(message);
    if (match) return localize(language, { en: message, ja: `「${match[1]}」の親グループ参照が循環しています。` });
    return localize(language, { en: message, ja: `図の構造を確認してください: ${message}` });
  }).join('\n');
  const toggle = (id: string) => {
    setApplyError('');
    setSelectedIds(previous => {
      const next = new Set(previous);
      if (next.has(id)) next.delete(id);
      else next.add(id);
      return next;
    });
  };
  const apply = async () => {
    if (busy || applyingRef.current || !selection.graph || !selectedIds.size) return;
    applyingRef.current = true;
    setSubmitting(true);
    setApplyError('');
    try {
      await onApply(selection.graph);
    } catch (cause) {
      if (mountedRef.current) setApplyError(cause instanceof Error ? cause.message : String(cause));
    } finally {
      applyingRef.current = false;
      if (mountedRef.current) setSubmitting(false);
    }
  };

  return createPortal(
    <div className="ai-review-overlay">
      <div ref={dialogRef} className="ai-review-dialog" role="dialog" aria-modal="true"
        aria-labelledby={titleId} aria-describedby={descriptionId} aria-busy={busy} tabIndex={-1}>
        <header className="ai-review-header">
          <h2 id={titleId}>{localize(language, { en: 'Review AI changes', ja: 'AIによる変更を確認' })}</h2>
          <button type="button" className="ai-review-close" onClick={onCancel} disabled={busy}
            aria-label={localize(language, { en: 'Cancel review', ja: '変更の確認をキャンセル' })}><X size={20} /></button>
        </header>
        <p id={descriptionId} className="ai-review-description">
          {localize(language, {
            en: 'Nothing has changed yet. Select exactly what to apply. Connected nodes and parent groups must remain valid.',
            ja: '図はまだ変更されていません。適用する変更を選択してください。接続先のノードと親グループが必要です。',
          })}
        </p>
        <div className="ai-review-selection-tools">
          <span aria-live="polite">{localize(language, {
            en: `${selectedIds.size} of ${changeSet.changes.length} selected`,
            ja: `${changeSet.changes.length}件中${selectedIds.size}件を選択`,
          })}</span>
          <button type="button" disabled={busy} onClick={() => setSelectedIds(new Set(changeSet.changes.map(change => change.id)))}>
            {localize(language, { en: 'Select all', ja: 'すべて選択' })}
          </button>
          <button type="button" disabled={busy} onClick={() => setSelectedIds(new Set())}>
            {localize(language, { en: 'Clear selection', ja: '選択を解除' })}
          </button>
        </div>
        <div className="ai-review-list">
          {!changeSet.changes.length && <p>{localize(language, {
            en: 'The proposal contains no diagram changes.', ja: '提案に図の変更はありません。',
          })}</p>}
          {(['node', 'edge'] as const).map(entity => {
            const changes = changeSet.changes.filter(change => change.entity === entity);
            if (!changes.length) return null;
            return <fieldset key={entity} disabled={busy}>
              <legend>{localize(language, {
                en: entity === 'node' ? 'Services and groups' : 'Connections',
                ja: entity === 'node' ? 'サービスとグループ' : '接続',
              })}</legend>
              {changes.map(change => <label key={change.id} className={`ai-review-item ai-review-item-${change.kind}`}>
                <input type="checkbox" checked={selectedIds.has(change.id)} onChange={() => toggle(change.id)} />
                <span className="ai-review-kind">{kindLabel(change.kind)}</span>
                <span className="ai-review-item-content">
                  <strong>{change.label}</strong>
                  {change.kind === 'change' && <span className="ai-review-fields">
                    {localize(language, { en: 'Updated fields: ', ja: '変更する項目: ' })}{change.fields.map(fieldLabel).join(', ')}
                  </span>}
                  {change.kind === 'change' && change.before && change.after &&
                    (change.entity === 'node' ? change.before.data?.label !== change.after.data?.label
                      : (change.before as any).label !== (change.after as any).label) &&
                    <span className="ai-review-fields">{String(change.entity === 'node'
                      ? change.before.data?.label ?? '' : (change.before as any).label ?? '')}
                      {' → '}{String(change.entity === 'node' ? change.after.data?.label ?? '' : (change.after as any).label ?? '')}</span>}
                </span>
                {change.costDelta !== undefined && <span className="ai-review-cost">
                  {money(change.costDelta)}{localize(language, { en: '/mo', ja: '/月' })}
                </span>}
              </label>)}
            </fieldset>;
          })}
        </div>
        {!!priced.length && <p className="ai-review-cost-summary">{localize(language, {
          en: `Selected estimated cost delta: ${money(delta)}/month (known estimates only; not a quote)`,
          ja: `選択した変更の概算コスト差額: ${money(delta)}/月（既知の概算のみ、見積書ではありません）`,
        })}</p>}
        {selection.error && <div className="ai-review-error" role="alert"><AlertCircle size={18} />
          <div><strong>{localize(language, {
            en: 'This selection cannot be applied. Select the required related changes or keep the affected resources.',
            ja: 'この選択は適用できません。必要な関連変更も選択するか、影響を受けるリソースを維持してください。',
          })}</strong><pre>{selectionError}</pre></div>
        </div>}
        {(error || applyError) && <div className="ai-review-error" role="alert"><AlertCircle size={18} />{error || applyError}</div>}
        <footer className="ai-review-footer">
          <button type="button" className="btn btn-secondary" onClick={onCancel} disabled={busy}>
            {localize(language, { en: 'Keep current diagram', ja: '現在の図を維持' })}
          </button>
          <button type="button" className="btn btn-primary" onClick={() => void apply()}
            disabled={busy || !!selection.error || !selectedIds.size}>
            {busy ? <Loader2 size={16} className="spin" /> : <Check size={16} />}
            {localize(language, { en: busy ? 'Applying…' : 'Apply selected changes', ja: busy ? '適用中…' : '選択した変更を適用' })}
          </button>
        </footer>
      </div>
    </div>, document.body,
  );
}
