// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

import React, { useState, useRef, useCallback, useEffect } from 'react';
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
  const [editData, setEditData] = useState({
    architectureName,
    author,
    version,
  });
  const [dragPosition, setDragPosition] = useState<{ x: number; y: number } | null>(null);
  const [isDragging, setIsDragging] = useState(false);
  const dragOffsetRef = useRef({ x: 0, y: 0 });
  const blockRef = useRef<HTMLDivElement>(null);

  useEffect(() => {
    if (!isEditing) {
      setEditData({ architectureName, author, version });
    }
  }, [architectureName, author, version, isEditing]);

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

  const handleMouseDown = useCallback((e: React.MouseEvent) => {
    if (isEditing || (e.target as HTMLElement).tagName === 'INPUT' || (e.target as HTMLElement).tagName === 'BUTTON') {
      return;
    }
    const el = blockRef.current;
    if (!el) return;
    const rect = el.getBoundingClientRect();
    const parentRect = el.offsetParent?.getBoundingClientRect() ?? { left: 0, top: 0 };
    const currentX = rect.left - parentRect.left;
    const currentY = rect.top - parentRect.top;
    dragOffsetRef.current = { x: e.clientX - currentX, y: e.clientY - currentY };
    setDragPosition({ x: currentX, y: currentY });
    setIsDragging(true);
  }, [isEditing]);

  React.useEffect(() => {
    if (!isDragging) return;
    const handleMouseMove = (e: MouseEvent) => {
      setDragPosition({
        x: e.clientX - dragOffsetRef.current.x,
        y: e.clientY - dragOffsetRef.current.y,
      });
    };
    const handleMouseUp = () => setIsDragging(false);
    window.addEventListener('mousemove', handleMouseMove);
    window.addEventListener('mouseup', handleMouseUp);
    return () => {
      window.removeEventListener('mousemove', handleMouseMove);
      window.removeEventListener('mouseup', handleMouseUp);
    };
  }, [isDragging]);

  const style: React.CSSProperties = dragPosition
    ? { left: dragPosition.x, top: dragPosition.y, bottom: 'auto' }
    : {};
  if (isDragging) style.cursor = 'grabbing';
  else if (!isEditing) style.cursor = 'grab';

  return (
    <div
      ref={blockRef}
      className={`title-block ${isDragging ? 'dragging' : ''}`}
      style={style}
      onMouseDown={handleMouseDown}
    >
      {isEditing ? (
        <div className="title-block-edit">
          <div className="title-block-row">
            <label>{t("Name:")}</label>
            <input
              type="text"
              value={editData.architectureName}
              onChange={(e) => setEditData({ ...editData, architectureName: e.target.value })}
              placeholder={t("Architecture name")}
            />
          </div>
          <div className="title-block-row">
            <label>{t("Author:")}</label>
            <input
              type="text"
              value={editData.author}
              onChange={(e) => setEditData({ ...editData, author: e.target.value })}
              placeholder={t("Your name")}
            />
          </div>
          <div className="title-block-row">
            <label>{t("Version:")}</label>
            <input
              type="text"
              value={editData.version}
              onChange={(e) => setEditData({ ...editData, version: e.target.value })}
              placeholder={t("1.0")}
            />
          </div>
          <div className="title-block-actions">
            <button onClick={handleSave} className="btn-save">{t("Save")}</button>
            <button onClick={handleCancel} className="btn-cancel">{t("Cancel")}</button>
          </div>
        </div>
      ) : (
        <div
          className="title-block-display"
          onDoubleClick={handleEdit}
          onKeyDown={(event) => {
            if (event.key === 'Enter' || event.key === ' ' || event.key === 'F2') {
              event.preventDefault();
              handleEdit();
            }
          }}
          role="button"
          tabIndex={0}
        >
          <div className="title-block-header">
            <span className="title-block-label">{t("ARCHITECTURE DIAGRAM")}</span>
          </div>
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
      )}
    </div>
  );
};

export default TitleBlock;
