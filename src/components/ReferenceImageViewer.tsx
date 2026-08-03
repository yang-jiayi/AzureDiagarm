// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

import React, { useCallback, useState, useRef } from 'react';
import { Image, X, Maximize2, Minimize2 } from 'lucide-react';
import './ReferenceImageViewer.css';
import { useLanguage } from '../i18n/LanguageContext';
import { useEscapeKey } from '../hooks/useEscapeKey';
import { useModalFocus } from '../hooks/useModalFocus';
import { localize } from '../i18n/localization';

interface ReferenceImageViewerProps {
  imageUrl: string;
  onDismiss: () => void;
}

const MIN_W = 160;
const MIN_H = 110;
const MAX_W = 700;
const MAX_H = 700;

const ReferenceImageViewer: React.FC<ReferenceImageViewerProps> = ({ imageUrl, onDismiss }) => {
  const { t, language } = useLanguage();
  const [isExpanded, setIsExpanded] = useState(false);
  const [isCollapsed, setIsCollapsed] = useState(false);
  const [pos, setPos] = useState(() => ({
    x: Math.max(0, window.innerWidth - 220),
    y: Math.max(0, window.innerHeight - 200),
  }));
  const [size, setSize] = useState({ w: 200, h: 170 });

  const dragState = useRef<{ startX: number; startY: number; origX: number; origY: number } | null>(null);
  const resizeState = useRef<{ startX: number; startY: number; origW: number; origH: number } | null>(null);
  const closeExpanded = useCallback(() => setIsExpanded(false), []);
  const expandedDialogRef = useModalFocus<HTMLDivElement>(isExpanded);
  useEscapeKey(isExpanded, closeExpanded);

  const onHeaderPointerDown = (e: React.PointerEvent<HTMLDivElement>) => {
    e.preventDefault();
    dragState.current = { startX: e.clientX, startY: e.clientY, origX: pos.x, origY: pos.y };
    (e.currentTarget as HTMLElement).setPointerCapture(e.pointerId);
  };

  const onHeaderPointerMove = (e: React.PointerEvent<HTMLDivElement>) => {
    if (!dragState.current) return;
    const dx = e.clientX - dragState.current.startX;
    const dy = e.clientY - dragState.current.startY;
    setPos({
      x: Math.max(0, Math.min(window.innerWidth - size.w, dragState.current.origX + dx)),
      y: Math.max(0, Math.min(window.innerHeight - 40, dragState.current.origY + dy)),
    });
  };

  const onHeaderPointerUp = () => { dragState.current = null; };

  const onResizePointerDown = (e: React.PointerEvent<HTMLDivElement>) => {
    e.preventDefault();
    e.stopPropagation();
    resizeState.current = { startX: e.clientX, startY: e.clientY, origW: size.w, origH: size.h };
    (e.currentTarget as HTMLElement).setPointerCapture(e.pointerId);
  };

  const onResizePointerMove = (e: React.PointerEvent<HTMLDivElement>) => {
    if (!resizeState.current) return;
    const dx = e.clientX - resizeState.current.startX;
    const dy = e.clientY - resizeState.current.startY;
    setSize({
      w: Math.max(MIN_W, Math.min(MAX_W, resizeState.current.origW + dx)),
      h: Math.max(MIN_H, Math.min(MAX_H, resizeState.current.origH + dy)),
    });
  };

  const onResizePointerUp = () => { resizeState.current = null; };

  const imageHeight = Math.max(60, size.h - 32);

  if (isCollapsed) {
    return (
      <button
        type="button"
        className="ref-image-viewer ref-image-collapsed"
        style={{ position: 'fixed', left: pos.x, top: pos.y }}
        onClick={() => setIsCollapsed(false)}
        aria-label={localize(language, {
          en: 'Show reference image',
          ja: '参照画像を表示',
        })}
      >
        <Image size={16} />
        <span>{t("Reference")}</span>
      </button>
    );
  }

  return (
    <>
      {/* Expanded overlay */}
      {isExpanded && (
        <div className="ref-image-overlay" onClick={closeExpanded}>
          <div
            ref={expandedDialogRef}
            className="ref-image-overlay-content"
            role="dialog"
            aria-modal="true"
            aria-labelledby="reference-image-dialog-title"
            tabIndex={-1}
            onClick={(e) => e.stopPropagation()}
          >
            <div className="ref-image-overlay-header">
              <span id="reference-image-dialog-title">
                <Image size={16} /> {' '}{t("Reference Diagram")}
              </span>
              <button
                type="button"
                onClick={closeExpanded}
                className="ref-image-close-btn"
                title={t("Close")}
                aria-label={t("Close")}
              >
                <X size={18} />
              </button>
            </div>
            <img src={imageUrl} alt={t("Reference architecture diagram")} className="ref-image-full" />
          </div>
        </div>
      )}

      {/* Floating thumbnail — draggable + resizable */}
      <div
        className="ref-image-viewer ref-image-thumbnail"
        style={{ position: 'fixed', left: pos.x, top: pos.y, width: size.w }}
        aria-hidden={isExpanded}
      >
        <div
          className="ref-image-header ref-image-drag-handle"
          onPointerDown={onHeaderPointerDown}
          onPointerMove={onHeaderPointerMove}
          onPointerUp={onHeaderPointerUp}
        >
          <span className="ref-image-label">
            <Image size={12} />
            {' '}{t("Reference")}{' '}</span>
          <div className="ref-image-actions">
            <button
              type="button"
              onPointerDown={e => e.stopPropagation()}
              onClick={() => setIsExpanded(true)}
              title={t("Expand")}
              aria-label={t("Expand")}
              className="ref-image-btn"
              tabIndex={isExpanded ? -1 : 0}
            >
              <Maximize2 size={12} />
            </button>
            <button
              type="button"
              onPointerDown={e => e.stopPropagation()}
              onClick={() => setIsCollapsed(true)}
              title={t("Minimize")}
              aria-label={t("Minimize")}
              className="ref-image-btn"
              tabIndex={isExpanded ? -1 : 0}
            >
              <Minimize2 size={12} />
            </button>
            <button
              type="button"
              onPointerDown={e => e.stopPropagation()}
              onClick={onDismiss}
              title={t("Dismiss")}
              aria-label={t("Dismiss")}
              className="ref-image-btn ref-image-btn-dismiss"
              tabIndex={isExpanded ? -1 : 0}
            >
              <X size={12} />
            </button>
          </div>
        </div>
        <button
          type="button"
          className="ref-image-preview-button"
          onClick={() => setIsExpanded(true)}
          aria-label={localize(language, {
            en: 'Expand reference image',
            ja: '参照画像を拡大',
          })}
          tabIndex={isExpanded ? -1 : 0}
        >
          <img
            src={imageUrl}
            alt={t("Reference architecture diagram")}
            className="ref-image-thumb"
            style={{ height: imageHeight }}
          />
        </button>
        {/* Resize handle — bottom-right corner */}
        <div
          className="ref-image-resize-handle"
          onPointerDown={onResizePointerDown}
          onPointerMove={onResizePointerMove}
          onPointerUp={onResizePointerUp}
        />
      </div>
    </>
  );
};

export default ReferenceImageViewer;
