// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

import { Frame, Hand, Maximize2, X, ZoomIn } from 'lucide-react';
import { MiniMap } from 'reactflow';
import { useLanguage } from '../i18n/LanguageContext';
import FeedbackFab from './FeedbackFab';
import { StartChooser } from './GuidedJourney';
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
  focusMode: boolean;
  showNavigationHint: boolean;
  showEmptyState: boolean;
  showFeedback: boolean;
  titleBlockData: CanvasTitleBlockData;
  generatedWithModel: { name: string; timeMs?: number } | null;
  forceCollapsed?: number;
  feedbackPulse?: boolean;
  onDismissNavigationHint: () => void;
  onFitView: () => void;
  onGuidedChat: () => void;
  onGenerateDiagram: () => void;
  onImportTemplate: () => void;
  onImportAzure: () => void;
  onTitleBlockUpdate: (
    data: { architectureName?: string; author?: string; version?: string },
  ) => void;
  onFeedback: () => void;
  onExitFocus: () => void;
}

export default function CanvasChrome({
  hasNodes,
  focusMode,
  showNavigationHint,
  showEmptyState,
  showFeedback,
  titleBlockData,
  generatedWithModel,
  forceCollapsed,
  feedbackPulse = false,
  onDismissNavigationHint,
  onFitView,
  onGuidedChat,
  onGenerateDiagram,
  onImportTemplate,
  onImportAzure,
  onTitleBlockUpdate,
  onFeedback,
  onExitFocus,
}: CanvasChromeProps) {
  const { t } = useLanguage();

  return (
    <>
      {focusMode && (
        <button
          type="button"
          className="canvas-focus-exit"
          onClick={onExitFocus}
          aria-keyshortcuts="Escape"
          autoFocus
        >
          <Maximize2 size={17} aria-hidden="true" />
          {t('Exit Focus')}
        </button>
      )}

      {!focusMode && hasNodes && (
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

      {!focusMode && showNavigationHint && hasNodes && (
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

      {!focusMode && showEmptyState && !hasNodes && (
        <StartChooser
          onGuidedChat={onGuidedChat}
          onGenerate={onGenerateDiagram}
          onImportTemplate={onImportTemplate}
          onImportAzure={onImportAzure}
        />
      )}

      {!focusMode && hasNodes && (
        <TitleBlock
          architectureName={titleBlockData.architectureName}
          author={titleBlockData.author}
          version={titleBlockData.version}
          date={titleBlockData.date}
          onUpdate={onTitleBlockUpdate}
        />
      )}
      {!focusMode && generatedWithModel && (
        <ModelBadge
          modelName={generatedWithModel.name}
          elapsedTimeMs={generatedWithModel.timeMs}
        />
      )}
      {!focusMode && hasNodes && <Legend forceCollapsed={forceCollapsed} />}
      {!focusMode && hasNodes && showFeedback && (
        <FeedbackFab pulse={feedbackPulse} onClick={onFeedback} />
      )}
    </>
  );
}
