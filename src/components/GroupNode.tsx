// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

import React, { memo, useEffect, useRef, useState } from 'react';
import { NodeProps, NodeResizer, useReactFlow } from 'reactflow';
import { Palette, Minimize2 } from 'lucide-react';
import { fitGroupToContent } from '../utils/groupUtils';
import './GroupNode.css';
import { useLanguage } from '../i18n/LanguageContext';
import { useNodeKeyboardInteraction } from '../hooks/useNodeKeyboardInteraction';
import { DEFAULT_ACCENT, matchZoneAccent, rgba } from '../utils/canvasPalette';

// Predefined color palette for groups
const COLOR_PALETTE = [
  { name: 'Gray', bg: 'rgba(107, 114, 128, 0.10)', border: '#6b7280', header: '#6b7280' },
  { name: 'Blue', bg: 'rgba(0, 120, 212, 0.10)', border: '#0078d4', header: '#0078d4' },
  { name: 'Green', bg: 'rgba(16, 185, 129, 0.10)', border: '#10b981', header: '#10b981' },
  { name: 'Orange', bg: 'rgba(245, 158, 11, 0.10)', border: '#f59e0b', header: '#f59e0b' },
  { name: 'Red', bg: 'rgba(239, 68, 68, 0.10)', border: '#ef4444', header: '#ef4444' },
  { name: 'Purple', bg: 'rgba(139, 92, 246, 0.10)', border: '#8b5cf6', header: '#8b5cf6' },
  { name: 'Cyan', bg: 'rgba(6, 182, 212, 0.10)', border: '#06b6d4', header: '#06b6d4' },
  { name: 'Pink', bg: 'rgba(236, 72, 153, 0.10)', border: '#ec4899', header: '#ec4899' },
  { name: 'Yellow', bg: 'rgba(234, 179, 8, 0.10)', border: '#eab308', header: '#eab308' },
  { name: 'Teal', bg: 'rgba(20, 184, 166, 0.10)', border: '#14b8a6', header: '#14b8a6' },
];

// Detect group category from label and return appropriate colors. The keyword
// table lives in canvasPalette so the exporters tint a zone the same way.
const getGroupColors = (label: string): { bg: string; border: string; header: string } => {
  const matched = matchZoneAccent(label);
  const accent = matched ?? DEFAULT_ACCENT;
  return { bg: rgba(accent, matched ? 0.10 : 0.08), border: accent, header: accent };
};

const GroupNode: React.FC<NodeProps> = memo(({ id, data, selected }) => {
  const { t } = useLanguage();
  const [isEditingLabel, setIsEditingLabel] = useState(false);
  const [label, setLabel] = useState(data.label || 'Group');
  const cancelLabelEditRef = useRef(false);
  const labelRef = useRef<HTMLDivElement>(null);
  const [showColorPicker, setShowColorPicker] = useState(false);
  const [customColor, setCustomColor] = useState(data.customColor || null);
  const { getNodes, setNodes } = useReactFlow();

  useEffect(() => {
    if (!isEditingLabel) {
      setLabel(data.label || 'Group');
    }
  }, [data.label, isEditingLabel]);

  useEffect(() => {
    setCustomColor(data.customColor || null);
  }, [data.customColor]);

  const handleFitToContent = () => {
    const allNodes = getNodes();
    const updated = fitGroupToContent(allNodes, id);
    if (updated) setNodes(updated);
  };

  const handleLabelDoubleClick = () => {
    cancelLabelEditRef.current = false;
    setIsEditingLabel(true);
  };
  const {
    handleFocus: handleNodeFocus,
    handleKeyDown: handleNodeKeyDown,
  } = useNodeKeyboardInteraction(id, handleLabelDoubleClick, data.label);

  const handleLabelChange = (e: React.ChangeEvent<HTMLInputElement>) => {
    setLabel(e.target.value);
  };

  const restoreLabelFocus = () => {
    window.requestAnimationFrame(() => labelRef.current?.focus());
  };

  const commitLabel = (restoreFocus = false) => {
    if (cancelLabelEditRef.current) {
      cancelLabelEditRef.current = false;
      return;
    }
    setNodes(nodes => nodes.map(node => (
      node.id === id
        ? { ...node, data: { ...node.data, label } }
        : node
    )));
    setIsEditingLabel(false);
    if (restoreFocus) restoreLabelFocus();
  };

  const handleLabelKeyDown = (e: React.KeyboardEvent<HTMLInputElement>) => {
    if (e.key === 'Enter') {
      commitLabel(true);
    } else if (e.key === 'Escape') {
      cancelLabelEditRef.current = true;
      setLabel(data.label || 'Group');
      setIsEditingLabel(false);
      restoreLabelFocus();
    }
  };

  const handleColorSelect = (colorScheme: typeof COLOR_PALETTE[0]) => {
    setCustomColor(colorScheme);
    setNodes(nodes => nodes.map(node => (
      node.id === id
        ? { ...node, data: { ...node.data, customColor: { ...colorScheme } } }
        : node
    )));
    setShowColorPicker(false);
  };

  const colors = customColor || getGroupColors(label);
  const groupStyle = {
    backgroundColor: colors.bg,
    borderColor: colors.border,
  };
  const headerStyle = {
    backgroundColor: `${colors.header}15`,
    borderBottomColor: colors.border,
  };
  const labelStyle = {
    color: colors.header,
  };

  return (
    <div
      className={`group-node ${selected ? 'selected' : ''}`}
      style={groupStyle}
      onFocus={handleNodeFocus}
    >
      <NodeResizer
        color="#0078d4"
        isVisible={selected}
        minWidth={200}
        minHeight={150}
      />
      <div className="group-node-header" style={headerStyle}>
        <div className="group-header-content">
          {isEditingLabel ? (
            <input
              type="text"
              value={label}
              onChange={handleLabelChange}
              onBlur={() => commitLabel()}
              onKeyDown={handleLabelKeyDown}
              autoFocus
              className="group-label-input"
            />
          ) : (
            <div
              ref={labelRef}
              data-node-keyboard-target
              className="group-label"
              onDoubleClick={handleLabelDoubleClick}
              onKeyDown={handleNodeKeyDown}
              role="button"
              tabIndex={0}
              aria-keyshortcuts="F2"
              title={t("Double-click to edit")}
              style={labelStyle}
            >
              {label}
            </div>
          )}
          <button
            className="fit-to-content-button"
            onClick={handleFitToContent}
            title={t("Fit to content")}
            style={{ color: colors.header }}
          >
            <Minimize2 size={16} />
          </button>
          <button
            className="color-picker-button"
            onClick={() => setShowColorPicker(!showColorPicker)}
            title={t("Change color")}
            style={{ color: colors.header }}
          >
            <Palette size={18} />
          </button>
        </div>
        
        {showColorPicker && (
          <div className="color-picker-panel">
            <div className="color-picker-title">{t("Choose Color")}</div>
            <div className="color-picker-grid">
              {COLOR_PALETTE.map((colorScheme) => (
                <button
                  key={colorScheme.name}
                  className={`color-option ${customColor?.name === colorScheme.name ? 'active' : ''}`}
                  onClick={() => handleColorSelect(colorScheme)}
                  style={{
                    backgroundColor: colorScheme.bg,
                    borderColor: colorScheme.border,
                  }}
                  title={colorScheme.name}
                >
                  <div
                    className="color-option-inner"
                    style={{ backgroundColor: colorScheme.border }}
                  />
                </button>
              ))}
            </div>
          </div>
        )}
      </div>
      <div className="group-node-content">
        {/* This area will contain other nodes visually */}
      </div>
    </div>
  );
});

GroupNode.displayName = 'GroupNode';

export default GroupNode;
