// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

import { useEffect, useId, useMemo, useRef, useState } from 'react';
import { createPortal } from 'react-dom';
import { AlertCircle, Check, Loader2, X } from 'lucide-react';
import { useLanguage } from '../i18n/LanguageContext';
import { localize, type LocalizedText } from '../i18n/localization';
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

interface ReviewField {
  path: string;
  label: LocalizedText;
  resource?: boolean;
}

// Explicit display fields keep callbacks, caches, and provider metadata out of the review.
const REVIEW_FIELDS: ReviewField[] = [
  { path: 'data.label', label: { en: 'Label', ja: 'ラベル' } },
  { path: 'label', label: { en: 'Connection label', ja: '接続ラベル' } },
  { path: 'data.serviceName', label: { en: 'Service', ja: 'サービス' } },
  { path: 'data.description', label: { en: 'Description', ja: '説明' } },
  { path: 'data.category', label: { en: 'Category', ja: 'カテゴリ' } },
  { path: 'data.tags', label: { en: 'Tags', ja: 'タグ' } },
  { path: 'data.customNote', label: { en: 'Note', ja: 'メモ' } },
  { path: 'source', label: { en: 'Source', ja: '接続元' }, resource: true },
  { path: 'target', label: { en: 'Target', ja: '接続先' }, resource: true },
  { path: 'parentNode', label: { en: 'Parent group', ja: '親グループ' }, resource: true },
  { path: 'parentId', label: { en: 'Parent group', ja: '親グループ' }, resource: true },
  { path: 'extent', label: { en: 'Group bounds', ja: 'グループ範囲' } },
  { path: 'type', label: { en: 'Type', ja: '種類' } },
  { path: 'position.x', label: { en: 'Horizontal position', ja: '横位置' } },
  { path: 'position.y', label: { en: 'Vertical position', ja: '縦位置' } },
  { path: 'sourceHandle', label: { en: 'Source handle', ja: '接続元ハンドル' } },
  { path: 'targetHandle', label: { en: 'Target handle', ja: '接続先ハンドル' } },
  { path: 'data.stylePreset', label: { en: 'Display style', ja: '表示スタイル' } },
  { path: 'data.pathStyle', label: { en: 'Connection path', ja: '接続線の形状' } },
  { path: 'data.direction', label: { en: 'Connection direction', ja: '接続方向' } },
  { path: 'data.flowAnimated', label: { en: 'Flow animation', ja: 'フロー アニメーション' } },
  { path: 'data.labelOffsetX', label: { en: 'Label horizontal offset', ja: 'ラベルの横オフセット' } },
  { path: 'data.labelOffsetY', label: { en: 'Label vertical offset', ja: 'ラベルの縦オフセット' } },
  { path: 'data.labelMaxWidth', label: { en: 'Label width', ja: 'ラベル幅' } },
  { path: 'data.customColor.bg', label: { en: 'Group background', ja: 'グループ背景色' } },
  { path: 'data.customColor.border', label: { en: 'Group border', ja: 'グループ枠線色' } },
  { path: 'data.customColor.header', label: { en: 'Group heading color', ja: 'グループ見出しの色' } },
  { path: 'style.color', label: { en: 'Text color', ja: '文字色' } },
  { path: 'style.background', label: { en: 'Background', ja: '背景' } },
  { path: 'style.backgroundColor', label: { en: 'Background color', ja: '背景色' } },
  { path: 'style.borderColor', label: { en: 'Border color', ja: '枠線色' } },
  { path: 'style.border', label: { en: 'Border', ja: '枠線' } },
  { path: 'style.borderWidth', label: { en: 'Border width', ja: '枠線幅' } },
  { path: 'style.width', label: { en: 'Width', ja: '幅' } },
  { path: 'style.height', label: { en: 'Height', ja: '高さ' } },
  { path: 'style.stroke', label: { en: 'Connection color', ja: '接続線の色' } },
  { path: 'style.strokeWidth', label: { en: 'Connection width', ja: '接続線の幅' } },
  { path: 'style.strokeDasharray', label: { en: 'Connection dash pattern', ja: '接続線の破線パターン' } },
  { path: 'style.opacity', label: { en: 'Opacity', ja: '不透明度' } },
  { path: 'markerStart.type', label: { en: 'Start arrow', ja: '始点の矢印' } },
  { path: 'markerEnd.type', label: { en: 'End arrow', ja: '終点の矢印' } },
  { path: 'labelStyle.color', label: { en: 'Label color', ja: 'ラベルの色' } },
  { path: 'labelStyle.fontSize', label: { en: 'Label font size', ja: 'ラベルの文字サイズ' } },
  { path: 'data.pricing.tier', label: { en: 'Pricing tier', ja: '料金プラン' } },
  { path: 'data.pricing.skuName', label: { en: 'SKU', ja: 'SKU' } },
  { path: 'data.pricing.quantity', label: { en: 'Quantity', ja: '数量' } },
  { path: 'data.pricing.estimatedCost', label: { en: 'Estimated unit cost (USD/month)', ja: '概算単価（USD/月）' } },
  { path: 'data.pricing.region', label: { en: 'Pricing region', ja: '料金リージョン' } },
  { path: 'data.pricing.unit', label: { en: 'Pricing unit', ja: '料金単位' } },
  { path: 'data.pricing.isCustom', label: { en: 'Custom pricing', ja: 'カスタム料金' } },
  { path: 'data.pricing.customPrice', label: { en: 'Custom unit cost (USD/month)', ja: 'カスタム単価（USD/月）' } },
  { path: 'data.pricing.usage.amount', label: { en: 'Usage amount', ja: '使用量' } },
  { path: 'data.pricing.usage.unit', label: { en: 'Usage unit', ja: '使用量の単位' } },
  { path: 'data.pricing.usage.unitPrice', label: { en: 'Usage unit price (USD)', ja: '使用量単価（USD）' } },
];

function isRecord(value: unknown): value is Record<string, unknown> {
  return value !== null && typeof value === 'object' && !Array.isArray(value);
}

function fieldValue(entity: unknown, path: string): unknown {
  return path.split('.').reduce<unknown>((value, key) => isRecord(value) ? value[key] : undefined, entity);
}

function displayValue(value: unknown): string | number | boolean | string[] | null | undefined {
  if (value === null) return null;
  if (typeof value === 'string' || typeof value === 'boolean') return value;
  if (typeof value === 'number' && Number.isFinite(value)) return value;
  if (Array.isArray(value) && value.every(item => typeof item === 'string')) return value;
  return undefined;
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
    return localize(language, fields[field] ?? { en: 'Other properties', ja: 'その他のプロパティ' });
  };
  const resourceName = (id: string, graph: DiagramGraph) => {
    const resource = graph.nodes.find(node => node.id === id);
    const label = fieldValue(resource, 'data.label') ?? fieldValue(resource, 'data.serviceName');
    if (typeof label !== 'string' || !label.trim()) return id;
    const duplicate = graph.nodes.some(node => node.id !== id
      && (fieldValue(node, 'data.label') ?? fieldValue(node, 'data.serviceName')) === label);
    return duplicate ? `${label} (${id})` : label;
  };
  const formatValue = (value: unknown, field: ReviewField, graph: DiagramGraph) => {
    const display = displayValue(value);
    if (display === undefined) return localize(language, { en: 'Not set', ja: '未設定' });
    if (display === null) return localize(language, field.path === 'data.pricing.estimatedCost'
      ? { en: 'Unpriced', ja: '料金未設定' } : { en: 'Not set', ja: '未設定' });
    if (display === '' || (Array.isArray(display) && !display.length)) return localize(language, { en: 'Empty', ja: '空' });
    if (typeof display === 'boolean') return localize(language, display
      ? { en: 'Yes', ja: 'はい' } : { en: 'No', ja: 'いいえ' });
    if (field.resource && typeof display === 'string') return resourceName(display, graph);
    return Array.isArray(display) ? display.join(', ') : String(display);
  };
  const renderLabelPreview = (change: DiagramChange) => {
    if (change.kind !== 'change') return null;
    const path = change.entity === 'node' ? 'data.label' : 'label';
    const before = fieldValue(change.before, path);
    const after = fieldValue(change.after, path);
    if (before === after || (typeof before !== 'string' && typeof after !== 'string')) return null;
    return <span className="ai-review-fields">
      {typeof before === 'string' ? before : ''}{' → '}{typeof after === 'string' ? after : ''}
    </span>;
  };
  const renderDetails = (change: DiagramChange) => {
    const rows = REVIEW_FIELDS.flatMap(field => {
      const before = displayValue(fieldValue(change.before, field.path));
      const after = displayValue(fieldValue(change.after, field.path));
      if (JSON.stringify(before) === JSON.stringify(after)) return [];
      return [{ field, before, after }];
    });
    return <details className="ai-review-details">
      <summary>{localize(language, {
        en: `View change details for ${change.label}`,
        ja: `${change.label} の変更内容を表示`,
      })}</summary>
      {rows.length ? <dl className="ai-review-diff">
        {rows.map(({ field, before, after }) => <div className="ai-review-detail-row" key={field.path}>
          <dt>{localize(language, field.label)}</dt>
          <dd><strong>{localize(language, { en: 'Before', ja: '変更前' })}</strong>
            <span>{formatValue(before, field, changeSet.before)}</span></dd>
          <dd><strong>{localize(language, { en: 'After', ja: '変更後' })}</strong>
            <span>{formatValue(after, field, changeSet.proposed)}</span></dd>
        </div>)}
      </dl> : <p>{localize(language, {
        en: 'This change has no supported field preview.',
        ja: 'この変更にはプレビューに対応した項目がありません。',
      })}</p>}
      <p className="ai-review-fields">{localize(language, {
        en: 'Supported diagram fields are shown. Internal metadata is omitted.',
        ja: '対応する図の項目を表示しています。内部メタデータは省略しています。',
      })}</p>
    </details>;
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
              {changes.map(change => <div key={change.id} className={`ai-review-item ai-review-item-${change.kind}`}>
                <label className="ai-review-choice">
                  <input type="checkbox" checked={selectedIds.has(change.id)} onChange={() => toggle(change.id)} />
                  <span className="ai-review-kind">{kindLabel(change.kind)}</span>
                  <span className="ai-review-item-content">
                    <strong>{change.label}</strong>
                    {change.kind === 'change' && <span className="ai-review-fields">
                      {localize(language, { en: 'Updated fields: ', ja: '変更する項目: ' })}{[...new Set(change.fields.map(fieldLabel))].join(', ')}
                    </span>}
                    {renderLabelPreview(change)}
                  </span>
                  {change.costDelta !== undefined && <span className="ai-review-cost">
                    {money(change.costDelta)}{localize(language, { en: '/mo', ja: '/月' })}
                  </span>}
                </label>
                {renderDetails(change)}
              </div>)}
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
