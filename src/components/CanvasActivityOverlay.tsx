// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

import { useEffect, useRef, useState } from 'react';
import { useLanguage } from '../i18n/LanguageContext';
import './CanvasActivityOverlay.css';

interface CanvasActivityOverlayProps {
  isApplyingRecommendations: boolean;
  isImportingTemplate: boolean;
  importFormatLabel: string;
  architecturePrompt: string;
  showArchitecturePrompt: boolean;
}

interface DragState {
  pointerId: number;
  offsetX: number;
  offsetY: number;
  canvas: DOMRect;
}

export default function CanvasActivityOverlay({
  isApplyingRecommendations,
  isImportingTemplate,
  importFormatLabel,
  architecturePrompt,
  showArchitecturePrompt,
}: CanvasActivityOverlayProps) {
  const { t } = useLanguage();
  const [promptPosition, setPromptPosition] = useState<{ x: number; y: number } | null>(null);
  const [isDragging, setIsDragging] = useState(false);
  const dragStateRef = useRef<DragState | null>(null);
  const promptRef = useRef<HTMLDivElement>(null);

  useEffect(() => {
    setPromptPosition(null);
    setIsDragging(false);
    dragStateRef.current = null;
  }, [architecturePrompt]);

  useEffect(() => {
    if (!isDragging) return;

    const handlePointerMove = (event: PointerEvent) => {
      const dragState = dragStateRef.current;
      const prompt = promptRef.current;
      if (!dragState || !prompt || dragState.pointerId !== event.pointerId) return;
      const margin = 8;
      setPromptPosition({
        x: Math.min(
          dragState.canvas.right - prompt.offsetWidth - margin,
          Math.max(dragState.canvas.left + margin, event.clientX - dragState.offsetX),
        ),
        y: Math.min(
          dragState.canvas.bottom - prompt.offsetHeight - margin,
          Math.max(dragState.canvas.top + margin, event.clientY - dragState.offsetY),
        ),
      });
    };
    const handlePointerUp = (event: PointerEvent) => {
      if (dragStateRef.current?.pointerId !== event.pointerId) return;
      dragStateRef.current = null;
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

  const handlePointerDown = (event: React.PointerEvent<HTMLDivElement>) => {
    const canvas = event.currentTarget.closest('.react-flow')?.getBoundingClientRect();
    if (!canvas) return;
    const rect = event.currentTarget.getBoundingClientRect();
    dragStateRef.current = {
      pointerId: event.pointerId,
      offsetX: event.clientX - rect.left,
      offsetY: event.clientY - rect.top,
      canvas,
    };
    setPromptPosition({ x: rect.left, y: rect.top });
    setIsDragging(true);
    event.preventDefault();
  };

  const showPrompt = showArchitecturePrompt && architecturePrompt.length > 0;
  if (!isApplyingRecommendations && !isImportingTemplate && !showPrompt) return null;

  return (
    <div className="canvas-activity-stack" aria-live="polite">
      {isApplyingRecommendations && (
        <div
          className="canvas-activity-banner canvas-activity-banner--recommendations"
          role="status"
        >
          <strong>{t('⏳ Applying recommendations...')}</strong>
          <span>{t('Regenerating architecture with improvements')}</span>
        </div>
      )}
      {isImportingTemplate && (
        <div className="canvas-activity-banner canvas-activity-banner--import" role="status">
          <strong>{t('📄 Parsing')} {importFormatLabel} {t('Template...')}</strong>
          <span>{t('Analyzing resources and generating architecture diagram')}</span>
        </div>
      )}
      {showPrompt && (
        <div
          ref={promptRef}
          className={`canvas-prompt-banner${isDragging ? ' is-dragging' : ''}`}
          style={promptPosition
            ? {
                position: 'fixed',
                left: promptPosition.x,
                top: promptPosition.y,
                justifySelf: 'auto',
                margin: 0,
                transform: 'none',
              }
            : undefined}
          onPointerDown={handlePointerDown}
        >
          <strong>{t('Generated from:')}</strong> {architecturePrompt}
        </div>
      )}
    </div>
  );
}
