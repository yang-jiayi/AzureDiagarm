// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

import { useEffect, useMemo, useRef, useState } from 'react';
import {
  AlignCenterHorizontal,
  AlignCenterVertical,
  AlignEndHorizontal,
  AlignEndVertical,
  AlignHorizontalDistributeCenter,
  AlignStartHorizontal,
  AlignStartVertical,
  AlignVerticalDistributeCenter,
  ChevronDown,
  SlidersHorizontal,
  X,
} from 'lucide-react';
import type { Node } from 'reactflow';
import { AVAILABLE_REGIONS, type AzureRegion } from '../services/regionalPricingService';
import {
  BULK_GROUP_COLORS,
  type BulkAlignmentType,
  type BulkTagMode,
} from '../utils/bulkNodeEditing';
import { useLanguage } from '../i18n/LanguageContext';
import { localize } from '../i18n/localization';
import './AlignmentToolbar.css';

export interface BulkEditRequest {
  targetGroupId?: string | null;
  region?: AzureRegion;
  stylePreset?: 'detailed' | 'presentation';
  groupColorName?: string | null;
  tags?: { mode: BulkTagMode; values: string[] };
  quantity?: number;
  customPrice?: number;
}

export interface BulkEditResult {
  updatedCount: number;
  pricingFailureCount: number;
}

interface AlignmentToolbarProps {
  selectedNodes: Node[];
  groups: Array<{ id: string; label: string }>;
  onAlign: (type: BulkAlignmentType) => void;
  onApplyBulkEdit: (request: BulkEditRequest) => Promise<BulkEditResult>;
}

export default function AlignmentToolbar({
  selectedNodes,
  groups,
  onAlign,
  onApplyBulkEdit,
}: AlignmentToolbarProps) {
  const { language } = useLanguage();
  const text = (en: string, ja: string) => localize(language, { en, ja });
  const rootRef = useRef<HTMLDivElement>(null);
  const [isEditorOpen, setIsEditorOpen] = useState(false);
  const [targetGroup, setTargetGroup] = useState('keep');
  const [region, setRegion] = useState('keep');
  const [stylePreset, setStylePreset] = useState('keep');
  const [groupColor, setGroupColor] = useState('keep');
  const [tagMode, setTagMode] = useState<BulkTagMode>('add');
  const [tags, setTags] = useState('');
  const [quantity, setQuantity] = useState('');
  const [customPrice, setCustomPrice] = useState('');
  const [isApplying, setIsApplying] = useState(false);
  const [status, setStatus] = useState('');

  const selectedServiceCount = selectedNodes.filter(node => node.type === 'azureNode').length;
  const selectedGroupCount = selectedNodes.filter(node => node.type === 'groupNode').length;
  const parsedTags = useMemo(
    () => tags.split(',').map(tag => tag.trim()).filter(Boolean),
    [tags],
  );
  const hasChanges = (
    targetGroup !== 'keep'
    || region !== 'keep'
    || stylePreset !== 'keep'
    || groupColor !== 'keep'
    || parsedTags.length > 0
    || quantity.trim() !== ''
    || customPrice.trim() !== ''
  );

  useEffect(() => {
    if (!isEditorOpen) return;
    const closeOnOutsideClick = (event: PointerEvent) => {
      if (!rootRef.current?.contains(event.target as globalThis.Node)) setIsEditorOpen(false);
    };
    const closeOnEscape = (event: KeyboardEvent) => {
      if (event.key === 'Escape') setIsEditorOpen(false);
    };
    document.addEventListener('pointerdown', closeOnOutsideClick);
    document.addEventListener('keydown', closeOnEscape);
    return () => {
      document.removeEventListener('pointerdown', closeOnOutsideClick);
      document.removeEventListener('keydown', closeOnEscape);
    };
  }, [isEditorOpen]);

  if (selectedNodes.length < 2) return null;

  const resetForm = () => {
    setTargetGroup('keep');
    setRegion('keep');
    setStylePreset('keep');
    setGroupColor('keep');
    setTags('');
    setQuantity('');
    setCustomPrice('');
  };

  const applyBulkEdit = async () => {
    if (!hasChanges || isApplying) return;
    setIsApplying(true);
    setStatus('');
    try {
      const request: BulkEditRequest = {};
      if (targetGroup === 'ungroup') request.targetGroupId = null;
      else if (targetGroup.startsWith('group:')) request.targetGroupId = targetGroup.slice(6);
      if (region !== 'keep') request.region = region as AzureRegion;
      if (stylePreset !== 'keep') {
        request.stylePreset = stylePreset as 'detailed' | 'presentation';
      }
      if (groupColor === 'reset') request.groupColorName = null;
      else if (groupColor !== 'keep') request.groupColorName = groupColor;
      if (parsedTags.length > 0) request.tags = { mode: tagMode, values: parsedTags };
      if (quantity.trim() !== '') request.quantity = Number(quantity);
      if (customPrice.trim() !== '') request.customPrice = Number(customPrice);

      const result = await onApplyBulkEdit(request);
      setStatus(result.pricingFailureCount > 0
        ? text(
            `Updated ${result.updatedCount} items; pricing refresh failed for ${result.pricingFailureCount}.`,
            `${result.updatedCount}件を更新しました。${result.pricingFailureCount}件の価格更新に失敗しました。`,
          )
        : text(
            `Updated ${result.updatedCount} selected items.`,
            `選択した${result.updatedCount}件を更新しました。`,
          ));
      resetForm();
      setIsEditorOpen(false);
    } catch (error) {
      setStatus(error instanceof Error
        ? error.message
        : text('Bulk edit failed.', '一括編集に失敗しました。'));
    } finally {
      setIsApplying(false);
    }
  };

  const alignmentButtons: Array<{
    type: BulkAlignmentType;
    label: string;
    icon: React.ReactNode;
  }> = [
    { type: 'left', label: text('Align left', '左揃え'), icon: <AlignStartHorizontal size={18} /> },
    { type: 'center-h', label: text('Align horizontal centers', '左右中央揃え'), icon: <AlignCenterHorizontal size={18} /> },
    { type: 'right', label: text('Align right', '右揃え'), icon: <AlignEndHorizontal size={18} /> },
    { type: 'top', label: text('Align top', '上揃え'), icon: <AlignStartVertical size={18} /> },
    { type: 'center-v', label: text('Align vertical centers', '上下中央揃え'), icon: <AlignCenterVertical size={18} /> },
    { type: 'bottom', label: text('Align bottom', '下揃え'), icon: <AlignEndVertical size={18} /> },
    { type: 'distribute-h', label: text('Distribute horizontally', '水平方向に均等配置'), icon: <AlignHorizontalDistributeCenter size={18} /> },
    { type: 'distribute-v', label: text('Distribute vertically', '垂直方向に均等配置'), icon: <AlignVerticalDistributeCenter size={18} /> },
  ];

  return (
    <div className="alignment-toolbar" ref={rootRef}>
      <span className="toolbar-label">{text('Align', '整列')}</span>
      <div className="alignment-toolbar-buttons" aria-label={text('Multi-selection alignment', '複数選択の整列')}>
        {alignmentButtons.map(button => (
          <button
            key={button.type}
            type="button"
            onClick={() => onAlign(button.type)}
            title={button.label}
            aria-label={button.label}
            className="alignment-toolbar-btn"
          >
            {button.icon}
          </button>
        ))}
      </div>

      <div className="alignment-toolbar-separator" />

      <button
        type="button"
        className="alignment-toolbar-edit"
        onClick={() => setIsEditorOpen(current => !current)}
        aria-expanded={isEditorOpen}
        aria-haspopup="dialog"
      >
        <SlidersHorizontal size={17} />
        {text('Bulk edit', '一括編集')}
        <ChevronDown size={14} />
      </button>

      <span className="alignment-toolbar-info toolbar-info">
        {text(`${selectedNodes.length} selected`, `${selectedNodes.length}件選択`)}
      </span>
      <span className="alignment-toolbar-status" aria-live="polite">{status}</span>

      {isEditorOpen && (
        <div
          className="bulk-edit-popover"
          role="dialog"
          aria-label={text('Bulk edit selected items', '選択項目を一括編集')}
        >
          <header>
            <div>
              <strong>{text('Bulk edit', '一括編集')}</strong>
              <span>
                {text(
                  `${selectedServiceCount} services · ${selectedGroupCount} groups`,
                  `サービス ${selectedServiceCount}件・グループ ${selectedGroupCount}件`,
                )}
              </span>
            </div>
            <button
              type="button"
              onClick={() => setIsEditorOpen(false)}
              aria-label={text('Close bulk editor', '一括編集を閉じる')}
            >
              <X size={17} />
            </button>
          </header>

          <div className="bulk-edit-grid">
            <label>
              <span>{text('Group', 'グループ')}</span>
              <select
                value={targetGroup}
                onChange={event => setTargetGroup(event.target.value)}
                aria-label={text('Group', 'グループ')}
              >
                <option value="keep">{text('Keep current groups', '現在のグループを維持')}</option>
                <option value="ungroup">{text('Remove from groups', 'グループから外す')}</option>
                {groups.map(group => (
                  <option key={group.id} value={`group:${group.id}`}>{group.label}</option>
                ))}
              </select>
            </label>

            <label>
              <span>{text('Azure region', 'Azure リージョン')}</span>
              <select
                value={region}
                onChange={event => setRegion(event.target.value)}
                aria-label={text('Azure region', 'Azure リージョン')}
              >
                <option value="keep">{text('Keep current regions', '現在のリージョンを維持')}</option>
                {AVAILABLE_REGIONS.map(item => (
                  <option key={item.id} value={item.id}>{item.flag} {item.displayName}</option>
                ))}
              </select>
            </label>

            <label>
              <span>{text('Service style', 'サービス スタイル')}</span>
              <select
                value={stylePreset}
                onChange={event => setStylePreset(event.target.value)}
                aria-label={text('Service style', 'サービス スタイル')}
              >
                <option value="keep">{text('Keep current styles', '現在のスタイルを維持')}</option>
                <option value="detailed">{text('Detailed', '詳細')}</option>
                <option value="presentation">{text('Presentation', 'プレゼンテーション')}</option>
              </select>
            </label>

            <label>
              <span>{text('Group color', 'グループ色')}</span>
              <select
                value={groupColor}
                onChange={event => setGroupColor(event.target.value)}
                aria-label={text('Group color', 'グループ色')}
              >
                <option value="keep">{text('Keep current colors', '現在の色を維持')}</option>
                <option value="reset">{text('Automatic color', '自動配色')}</option>
                {BULK_GROUP_COLORS.map(color => (
                  <option key={color.name} value={color.name}>{color.name}</option>
                ))}
              </select>
            </label>

            <div className="bulk-edit-tags">
              <span id="bulk-edit-tags-label">
                {text('Tags (comma-separated)', 'タグ（カンマ区切り）')}
              </span>
              <div>
                <select
                  value={tagMode}
                  onChange={event => setTagMode(event.target.value as BulkTagMode)}
                  aria-label={text('Tag operation', 'タグ操作')}
                >
                  <option value="add">{text('Add', '追加')}</option>
                  <option value="replace">{text('Replace', '置換')}</option>
                  <option value="remove">{text('Remove', '削除')}</option>
                </select>
                <input
                  value={tags}
                  onChange={event => setTags(event.target.value)}
                  maxLength={480}
                  aria-labelledby="bulk-edit-tags-label"
                  placeholder={text('production, critical', 'production, critical')}
                />
              </div>
            </div>

            <label>
              <span>{text('Quantity', '数量')}</span>
              <input
                type="number"
                min="1"
                max="100000"
                step="1"
                value={quantity}
                onChange={event => setQuantity(event.target.value)}
                aria-label={text('Quantity', '数量')}
                placeholder={text('Keep current', '現在値を維持')}
              />
            </label>

            <label>
              <span>{text('Custom monthly unit price (USD)', '独自の月額単価（USD）')}</span>
              <input
                type="number"
                min="0"
                max="1000000000"
                step="0.01"
                value={customPrice}
                onChange={event => setCustomPrice(event.target.value)}
                aria-label={text('Custom monthly unit price (USD)', '独自の月額単価（USD）')}
                placeholder={text('Keep catalog price', 'カタログ価格を維持')}
              />
            </label>
          </div>

          <footer>
            <button type="button" onClick={resetForm} className="bulk-edit-reset">
              {text('Reset', 'リセット')}
            </button>
            <button
              type="button"
              onClick={() => void applyBulkEdit()}
              className="bulk-edit-apply"
              disabled={!hasChanges || isApplying}
            >
              {isApplying ? text('Applying...', '適用中...') : text('Apply to selection', '選択項目に適用')}
            </button>
          </footer>
        </div>
      )}
    </div>
  );
}
