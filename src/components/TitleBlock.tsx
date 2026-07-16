// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

import React, { useState, useRef, useCallback, useEffect, useId } from 'react';
import { ChevronDown, ChevronUp } from 'lucide-react';
import './TitleBlock.css';
import { useLanguage } from '../i18n/LanguageContext';

interface TitleBlockProps {
  architectureName?: string;
  author?: string;
  version?: string;
  date?: string;
  onUpdate?: (data: { architectureName?: string; author?: string; version?: string }) => void;
}

const TitleBlock: React.FC<TitleBlockProps> = ({
  architectureName = 'Untitled Architecture',
  author = 'Unknown',
  version = '1.0',
  date = new Date().toLocaleDateString(),
  onUpdate,
}) => {
  const { t } = useLanguage();
  const [isEditing, setIsEditing] = useState(false);
  const [isCollapsed, setIsCollapsed] = useState(
    () => typeof window !== 'undefined' && window.matchMedia('(max-width: 1024px)').matches,
  );
  const [editData, setEditData] = useState({
    architectureName,
    author,
    version,
  });
  const [dragPosition, setDragPosition] = useState<{ x: number; y: number } | null>(null);
  const [isDragging, setIsDragging] = useState(false);
  const dragOffsetRef = useRef({ x: 0, y: 0 });
  const activePointerRef = useRef<number | null>(null);
  const blockRef = useRef<HTMLDivElement>(null);
  const nameInputId = useId();
  const authorInputId = useId();
  const versionInputId = useId();

  useEffect(() => {
    if (!isEditing) {
      setEditData({ architectureName, author, version });
    }
  }, [architectureName, author, version, isEditing]);

  useEffect(() => {
    const compactViewport = window.matchMedia('(max-width: 1024px)');
    const collapseForCompactViewport = (event: MediaQueryListEvent) => {
      if (event.matches) setIsCollapsed(true);
    };
    compactViewport.addEventListener('change', collapseForCompactViewport);
    return () => compactViewport.removeEventListener('change', collapseForCompactViewport);
  }, []);

  const handleEdit = () => {
    setEditData({ architectureName, author, version });
    setIsEditing(true);
  };

  const handleSave = () => {
    setIsEditing(false);
    onUpdate?.(editData);
  };

  const handleCancel = () => {
    setIsEditing(false);
    setEditData({ architectureName, author, version });
  };

  const handlePointerDown = useCallback((event: React.PointerEvent<HTMLDivElement>) => {
    if ((event.target as HTMLElement).closest('button')) return;
    const el = blockRef.current;
    if (!el) return;
    const rect = el.getBoundingClientRect();
    const parentRect = el.offsetParent?.getBoundingClientRect() ?? { left: 0, top: 0 };
    const currentX = rect.left - parentRect.left;
    const currentY = rect.top - parentRect.top;
    dragOffsetRef.current = {
      x: event.clientX - parentRect.left - currentX,
      y: event.clientY - parentRect.top - currentY,
    };
    activePointerRef.current = event.pointerId;
    setDragPosition({ x: currentX, y: currentY });
    setIsDragging(true);
    event.currentTarget.setPointerCapture(event.pointerId);
    event.preventDefault();
  }, []);

  useEffect(() => {
    if (!isDragging) return;
    const handlePointerMove = (event: PointerEvent) => {
      if (event.pointerId !== activePointerRef.current) return;
      const el = blockRef.current;
      const parent = el?.offsetParent;
      if (!el || !parent) return;
      const parentRect = parent.getBoundingClientRect();
      const maxX = Math.max(8, parentRect.width - el.offsetWidth - 8);
      const maxY = Math.max(8, parentRect.height - el.offsetHeight - 8);
      setDragPosition({
        x: Math.min(maxX, Math.max(8, event.clientX - parentRect.left - dragOffsetRef.current.x)),
        y: Math.min(maxY, Math.max(8, event.clientY - parentRect.top - dragOffsetRef.current.y)),
      });
    };
    const handlePointerUp = (event: PointerEvent) => {
      if (event.pointerId !== activePointerRef.current) return;
      activePointerRef.current = null;
      setIsDragging(false);
    };
    window.addEventListener('pointermove', handlePointerMove);
    window.addEventListener('pointerup', handlePointerUp);
    window.addEventListener('pointercancel', handlePointerUp);
    return () => {
      window.removeEventListener('pointermove', handlePointerMove);
      window.removeEventListener('pointerup', handlePointerUp);
      window.removeEventListener('pointercancel', handlePointerUp);
    };
  }, [isDragging]);

  const style: React.CSSProperties = dragPosition
    ? { left: dragPosition.x, right: 'auto', top: dragPosition.y, bottom: 'auto' }
    : {};

  return (
    <div
      ref={blockRef}
      className={`title-block${isCollapsed ? ' collapsed' : ''}${isDragging ? ' dragging' : ''}`}
      style={style}
    >
      <div className="title-block-header" onPointerDown={handlePointerDown}>
        <span className="title-block-label">{t("ARCHITECTURE DIAGRAM")}</span>
        <button
          type="button"
          className="title-block-toggle"
          onClick={() => setIsCollapsed((current) => !current)}
          aria-expanded={!isCollapsed}
          aria-controls="diagram-title-details"
          title={isCollapsed ? t('title.showDetails') : t('title.hideDetails')}
        >
          {isCollapsed ? <ChevronDown size={15} /> : <ChevronUp size={15} />}
        </button>
      </div>
      {!isCollapsed && (isEditing ? (
        <div className="title-block-edit">
          <div className="title-block-row">
            <label htmlFor={nameInputId}>{t("Name:")}</label>
            <input
              id={nameInputId}
              type="text"
              value={editData.architectureName}
              onChange={(e) => setEditData({ ...editData, architectureName: e.target.value })}
              placeholder={t("Architecture name")}
            />
          </div>
          <div className="title-block-row">
            <label htmlFor={authorInputId}>{t("Author:")}</label>
            <input
              id={authorInputId}
              type="text"
              value={editData.author}
              onChange={(e) => setEditData({ ...editData, author: e.target.value })}
              placeholder={t("Your name")}
            />
          </div>
          <div className="title-block-row">
            <label htmlFor={versionInputId}>{t("Version:")}</label>
            <input
              id={versionInputId}
              type="text"
              value={editData.version}
              onChange={(e) => setEditData({ ...editData, version: e.target.value })}
              placeholder={t("1.0")}
            />
          </div>
          <div className="title-block-actions">
            <button type="button" onClick={handleSave} className="btn-save">{t("Save")}</button>
            <button type="button" onClick={handleCancel} className="btn-cancel">{t("Cancel")}</button>
          </div>
        </div>
      ) : (
        <div
          className="title-block-display"
          id="diagram-title-details"
          onDoubleClick={handleEdit}
          onKeyDown={(event) => {
            if (event.key === 'Enter' || event.key === ' ' || event.key === 'F2') {
              event.preventDefault();
              handleEdit();
            }
          }}
          role="button"
          tabIndex={0}
          aria-label={t('title.editDetails')}
        >
          <div className="title-block-content">
            <div className="title-block-row">
              <span className="title-block-field">{t("Name:")}</span>
              <span className="title-block-value">{architectureName}</span>
            </div>
            <div className="title-block-row">
              <span className="title-block-field">{t("Author:")}</span>
              <span className="title-block-value">{author}</span>
            </div>
            <div className="title-block-row">
              <span className="title-block-field">{t("Date:")}</span>
              <span className="title-block-value">{date}</span>
            </div>
            <div className="title-block-row">
              <span className="title-block-field">{t("Version:")}</span>
              <span className="title-block-value">{version}</span>
            </div>
          </div>
          <div className="title-block-hint">{t("Double-click to edit")}</div>
        </div>
      ))}
    </div>
  );
};

export default TitleBlock;
