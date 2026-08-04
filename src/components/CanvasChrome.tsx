// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

import { Frame, Hand, MessagesSquare, X, ZoomIn } from 'lucide-react';
import { MiniMap } from 'reactflow';
import { useLanguage } from '../i18n/LanguageContext';
import FeedbackFab from './FeedbackFab';
import Legend from './Legend';
import ModelBadge from './ModelBadge';
import TitleBlock from './TitleBlock';
import './CanvasChrome.css';

interface CanvasTitleBlockData {
  architectureName: string;
  author: string;
  version: string;
  date: string;
}

interface CanvasChromeProps {
  hasNodes: boolean;
  showNavigationHint: boolean;
  showEmptyState: boolean;
  showFeedback: boolean;
  titleBlockData: CanvasTitleBlockData;
  generatedWithModel: { name: string; timeMs?: number } | null;
  forceCollapsed?: number;
  feedbackPulse?: boolean;
  onDismissNavigationHint: () => void;
  onFitView: () => void;
  onOpenChat: () => void;
  onTitleBlockUpdate: (
    data: { architectureName?: string; author?: string; version?: string },
  ) => void;
  onFeedback: () => void;
}

export default function CanvasChrome({
  hasNodes,
  showNavigationHint,
  showEmptyState,
  showFeedback,
  titleBlockData,
  generatedWithModel,
  forceCollapsed,
  feedbackPulse = false,
  onDismissNavigationHint,
  onFitView,
  onOpenChat,
  onTitleBlockUpdate,
  onFeedback,
}: CanvasChromeProps) {
  const { t } = useLanguage();

  return (
    <>
      {hasNodes && (
        <>
          <div className="nav-minimap-caption">{t('canvas.miniMapCaption')}</div>
          <MiniMap
            pannable
            zoomable
            position="bottom-right"
            className="nav-minimap"
            ariaLabel={t('canvas.miniMap')}
            nodeColor="#60a5fa"
            nodeStrokeColor="#3b82f6"
            maskColor="rgba(30, 41, 59, 0.45)"
          />
        </>
      )}

      {showNavigationHint && hasNodes && (
        <div className="canvas-nav-hint" role="note" aria-label={t('Canvas navigation tips')}>
          <div className="canvas-nav-hint-tips">
            <span className="canvas-nav-hint-tip canvas-nav-hint-desktop">
              <ZoomIn size={15} /> {t('Scroll to zoom in / out')}
            </span>
            <span className="canvas-nav-hint-sep canvas-nav-hint-desktop" aria-hidden="true">
              {t('·')}
            </span>
            <span className="canvas-nav-hint-tip canvas-nav-hint-desktop">
              <Hand size={15} /> {t('Right-click + drag to pan')}
            </span>
            <span className="canvas-nav-hint-sep canvas-nav-hint-desktop" aria-hidden="true">
              {t('·')}
            </span>
            <span className="canvas-nav-hint-tip canvas-nav-hint-mobile">
              <Hand size={15} /> {t('canvas.touchNavigation')}
            </span>
            <button
              type="button"
              className="canvas-nav-hint-fit"
              onClick={onFitView}
              title={t('Zoom to fit the whole diagram in view')}
            >
              <Frame size={15} /> {t('Fit to view')}
            </button>
          </div>
          <button
            type="button"
            className="canvas-nav-hint-close"
            onClick={onDismissNavigationHint}
            title={t("Dismiss (won't show again)")}
            aria-label={t('Dismiss navigation tips')}
          >
            <X size={14} />
          </button>
        </div>
      )}

      {showEmptyState && !hasNodes && (
        <div className="canvas-empty-cta" role="note" aria-label={t('Get started')}>
          <div className="canvas-empty-cta-inner">
            <MessagesSquare size={34} className="canvas-empty-cta-icon" />
            <h2 className="canvas-empty-cta-title">{t('Start with a conversation')}</h2>
            <p className="canvas-empty-cta-desc">
              {t(
                'Describe what you want to build in plain English — I’ll draw the first version, then you refine it step by step.',
              )}
            </p>
            <button type="button" className="canvas-empty-cta-btn" onClick={onOpenChat}>
              <MessagesSquare size={18} /> {t('Start with a conversation')}
            </button>
            <span className="canvas-empty-cta-alt">
              {t('or use')} <strong>{t('Generate with AI')}</strong>
              {' '}{t('· or add services from the left panel')}
            </span>
          </div>
        </div>
      )}

      {hasNodes && (
        <TitleBlock
          architectureName={titleBlockData.architectureName}
          author={titleBlockData.author}
          version={titleBlockData.version}
          date={titleBlockData.date}
          onUpdate={onTitleBlockUpdate}
        />
      )}
      {generatedWithModel && (
        <ModelBadge
          modelName={generatedWithModel.name}
          elapsedTimeMs={generatedWithModel.timeMs}
        />
      )}
      <Legend forceCollapsed={forceCollapsed} />
      {showFeedback && <FeedbackFab pulse={feedbackPulse} onClick={onFeedback} />}
    </>
  );
}
