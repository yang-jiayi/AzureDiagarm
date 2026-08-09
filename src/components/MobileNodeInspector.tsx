// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

import { useEffect, useState } from 'react';
import type { Node } from 'reactflow';
import { DollarSign, Edit3, Save, Trash2, X } from 'lucide-react';
import { useMediaQuery } from '../hooks/useMediaQuery';
import { useLanguage } from '../i18n/LanguageContext';
import { localize } from '../i18n/localization';
import { MEDIA_QUERIES } from '../styles/breakpoints';
import ResponsiveDrawer from './ResponsiveDrawer';
import './MobileNodeInspector.css';

interface MobileNodeInspectorProps {
  node: Node | null;
  onUpdate: (nodeId: string, patch: Record<string, unknown>) => void;
  onDelete: (nodeId: string) => void;
  onOpenPricing: (nodeId: string) => void;
}

interface GroupColorScheme {
  name: string;
  bg: string;
  border: string;
  header: string;
}

function createGroupColorScheme(color: string): GroupColorScheme {
  const normalized = /^#[0-9a-f]{6}$/i.test(color) ? color.toLowerCase() : '#0f6cbd';
  const red = Number.parseInt(normalized.slice(1, 3), 16);
  const green = Number.parseInt(normalized.slice(3, 5), 16);
  const blue = Number.parseInt(normalized.slice(5, 7), 16);
  return {
    name: 'Custom',
    bg: `rgba(${red}, ${green}, ${blue}, 0.10)`,
    border: normalized,
    header: normalized,
  };
}

export default function MobileNodeInspector({
  node,
  onUpdate,
  onDelete,
  onOpenPricing,
}: MobileNodeInspectorProps) {
  const { language } = useLanguage();
  const isMobile = useMediaQuery(MEDIA_QUERIES.compact);
  const [isOpen, setIsOpen] = useState(false);
  const [label, setLabel] = useState('');
  const [description, setDescription] = useState('');
  const [color, setColor] = useState('#0f6cbd');
  const selectedLabel = String(node?.data?.label || node?.data?.serviceName || '');
  const selectedDescription = String(node?.data?.description || '');
  const selectedCustomColor = node?.data?.customColor as Partial<GroupColorScheme> | undefined;
  const selectedColor = String(
    selectedCustomColor?.header
    || selectedCustomColor?.border
    || node?.data?.color
    || '#0f6cbd',
  );

  useEffect(() => {
    setIsOpen(false);
    setLabel(selectedLabel);
    setDescription(selectedDescription);
    setColor(selectedColor);
  }, [node?.id, selectedColor, selectedDescription, selectedLabel]);

  if (!isMobile || !node) return null;
  const isGroup = node.type === 'groupNode';
  const nodeLabel = selectedLabel;

  const save = () => {
    onUpdate(node.id, {
      label: label.trim() || nodeLabel,
      description: description.trim(),
      ...(isGroup ? { customColor: createGroupColorScheme(color) } : {}),
    });
    setIsOpen(false);
  };

  return (
    <>
      <button
        type="button"
        className="mobile-node-inspector-launcher"
        onClick={() => setIsOpen(true)}
        aria-label={localize(language, {
          en: `Edit selected node: ${nodeLabel}`,
          ja: `選択したノードを編集: ${nodeLabel}`,
        })}
      >
        <Edit3 size={16} />
        <span>{nodeLabel}</span>
        <strong>{localize(language, { en: 'Edit', ja: '編集' })}</strong>
      </button>

      <ResponsiveDrawer
        isOpen={isOpen}
        modal
        placement="bottom"
        className="mobile-node-inspector"
        backdropClassName="mobile-node-inspector-backdrop"
        ariaLabel={localize(language, { en: 'Edit selected node', ja: '選択したノードを編集' })}
        onClose={() => setIsOpen(false)}
        backgroundSelectors={[
          '.app-header',
          '.app > .workspace',
          '.mobile-node-inspector-launcher',
        ]}
      >
        <header>
          <div>
            <span>{isGroup
              ? localize(language, { en: 'Group', ja: 'グループ' })
              : localize(language, { en: 'Azure service', ja: 'Azure サービス' })}</span>
            <h2>{nodeLabel}</h2>
          </div>
          <button type="button" onClick={() => setIsOpen(false)} aria-label={localize(language, { en: 'Close node editor', ja: 'ノード エディターを閉じる' })}>
            <X size={20} />
          </button>
        </header>

        <div className="mobile-node-inspector-body">
          <label>
            <span>{localize(language, { en: 'Display name', ja: '表示名' })}</span>
            <input value={label} onChange={event => setLabel(event.target.value)} maxLength={120} />
          </label>
          <label>
            <span>{localize(language, { en: 'Description', ja: '説明' })}</span>
            <textarea value={description} onChange={event => setDescription(event.target.value)} rows={3} maxLength={500} />
          </label>
          {isGroup && (
            <label className="mobile-node-color">
              <span>{localize(language, { en: 'Group color', ja: 'グループ色' })}</span>
              <input type="color" value={color} onChange={event => setColor(event.target.value)} />
            </label>
          )}
          {!isGroup && (
            <div className="mobile-node-service-meta">
              <span>{localize(language, { en: 'Service type', ja: 'サービス種類' })}</span>
              <strong>{String(node.data?.serviceName || nodeLabel)}</strong>
            </div>
          )}
        </div>

        <footer>
          <button type="button" className="btn btn-danger" onClick={() => { onDelete(node.id); setIsOpen(false); }}>
            <Trash2 size={17} />
            {localize(language, { en: 'Delete', ja: '削除' })}
          </button>
          {!isGroup && (
            <button type="button" className="btn btn-secondary" onClick={() => { setIsOpen(false); onOpenPricing(node.id); }}>
              <DollarSign size={17} />
              {localize(language, { en: 'Cost', ja: 'コスト' })}
            </button>
          )}
          <button type="button" className="btn btn-primary" onClick={save}>
            <Save size={17} />
            {localize(language, { en: 'Save changes', ja: '変更を保存' })}
          </button>
        </footer>
      </ResponsiveDrawer>
    </>
  );
}
