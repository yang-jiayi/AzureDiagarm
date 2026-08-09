// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

import React from 'react';
import {
  Check,
  FileCode2,
  Import,
  LayoutTemplate,
  MessageSquare,
  Presentation,
  ShieldCheck,
  Sparkles,
  Wrench,
  type LucideIcon,
} from 'lucide-react';
import { useLanguage } from '../i18n/LanguageContext';
import { localize, type LocalizedText } from '../i18n/localization';
import './GuidedJourney.css';

export type JourneyStep = 'create' | 'refine' | 'validate' | 'deliver';

interface JourneyStripProps {
  hasDiagram: boolean;
  hasValidation: boolean;
  hasDeploymentGuide: boolean;
  isValidating?: boolean;
  isGeneratingGuide?: boolean;
  onStep: (step: JourneyStep) => void;
}

const STEPS: Array<{
  id: JourneyStep;
  number: number;
  title: LocalizedText;
  detail: LocalizedText;
  Icon: LucideIcon;
}> = [
  {
    id: 'create',
    number: 1,
    title: { en: 'Create', ja: '作成' },
    detail: { en: 'Chat, brief, image, template, or import', ja: 'チャット、要件、画像、テンプレート、取り込み' },
    Icon: Sparkles,
  },
  {
    id: 'refine',
    number: 2,
    title: { en: 'Refine', ja: '改善' },
    detail: { en: 'Chat or edit the canvas', ja: 'チャットまたはキャンバスで編集' },
    Icon: Wrench,
  },
  {
    id: 'validate',
    number: 3,
    title: { en: 'Validate & Improve', ja: '検証と改善' },
    detail: { en: 'Review, apply, and revalidate', ja: '確認、適用、再検証' },
    Icon: ShieldCheck,
  },
  {
    id: 'deliver',
    number: 4,
    title: { en: 'Share or Build', ja: '共有または構築' },
    detail: { en: 'Export or create deployment artifacts', ja: 'エクスポートまたはデプロイ成果物を作成' },
    Icon: Presentation,
  },
];

export const JourneyStrip: React.FC<JourneyStripProps> = ({
  hasDiagram,
  hasValidation,
  hasDeploymentGuide,
  isValidating,
  isGeneratingGuide,
  onStep,
}) => {
  const { language } = useLanguage();
  const active: JourneyStep = !hasDiagram
    ? 'create'
    : isValidating
      ? 'validate'
      : isGeneratingGuide || hasDeploymentGuide
        ? 'deliver'
        : hasValidation
          ? 'deliver'
          : 'refine';

  const complete = (step: JourneyStep) => {
    if (step === 'create') return hasDiagram;
    if (step === 'refine') return hasDiagram && (hasValidation || hasDeploymentGuide);
    if (step === 'validate') return hasValidation;
    return hasDeploymentGuide;
  };

  return (
    <nav
      className="journey-strip"
      aria-label={localize(language, {
        en: 'Recommended architecture workflow',
        ja: '推奨アーキテクチャ ワークフロー',
      })}
    >
      <span className="journey-strip-label">
        {localize(language, { en: 'Recommended workflow', ja: '推奨ワークフロー' })}
      </span>
      <div className="journey-steps">
        {STEPS.map(({ id, number, title, detail, Icon }, index) => {
          const isActive = active === id;
          const isComplete = complete(id);
          const localizedTitle = localize(language, title);
          const localizedDetail = localize(language, detail);
          return (
            <React.Fragment key={id}>
              {index > 0 && <span className={`journey-connector${isComplete ? ' complete' : ''}`} aria-hidden="true" />}
              <button
                type="button"
                className={`journey-step${isActive ? ' active' : ''}${isComplete ? ' complete' : ''}`}
                onClick={() => onStep(id)}
                aria-current={isActive ? 'step' : undefined}
                title={`${number}. ${localizedTitle}: ${localizedDetail}`}
              >
                <span className="journey-step-number" aria-hidden="true">
                  {isComplete ? <Check size={13} /> : number}
                </span>
                <Icon size={16} />
                <span className="journey-step-copy">
                  <strong>{localizedTitle}</strong>
                  <small>{localizedDetail}</small>
                </span>
              </button>
            </React.Fragment>
          );
        })}
      </div>
    </nav>
  );
};

interface StartChooserProps {
  onGuidedChat: () => void;
  onGenerate: () => void;
  onBrowseTemplates: () => void;
  onImportTemplate: () => void;
  onImportAzure: () => void;
}

export const StartChooser: React.FC<StartChooserProps> = ({
  onGuidedChat,
  onGenerate,
  onBrowseTemplates,
  onImportTemplate,
  onImportAzure,
}) => {
  const { t, language } = useLanguage();
  return (
    <div className="start-chooser" role="region" aria-label={t('Choose how to start')}>
      <div className="start-chooser-heading">
        <h2>{t('How would you like to start?')}</h2>
        <p>{t('Choose a path now. All paths lead to the same editable canvas and Guided Chat refinement.')}</p>
      </div>
      <div className="start-choice-grid">
        <button type="button" className="start-choice recommended" onClick={onGuidedChat}>
          <span className="start-choice-icon"><MessageSquare size={24} /></span>
          <span className="start-choice-badge">{t('Recommended for first-time users')}</span>
          <strong>{t('Guided Chat')}</strong>
          <span>{t('Describe the outcome, create the first diagram, and keep refining in one conversation.')}</span>
        </button>
        <button type="button" className="start-choice" onClick={onGenerate}>
          <span className="start-choice-icon"><Sparkles size={24} /></span>
          <strong>{t('Generate Diagram')}</strong>
          <span>{t('Use a detailed prompt, upload a sketch, and choose Topology, Blueprint, or both.')}</span>
        </button>
        <button type="button" className="start-choice" onClick={onBrowseTemplates}>
          <span className="start-choice-icon"><LayoutTemplate size={24} /></span>
          <strong>{localize(language, { en: 'Browse Templates', ja: 'テンプレートを見る' })}</strong>
          <span>{localize(language, {
            en: 'Preview a proven starting architecture, then customize every service and connection.',
            ja: '実績あるスターター構成をプレビューし、すべてのサービスと接続を編集できます。',
          })}</span>
        </button>
        <div className="start-choice import-choice">
          <span className="start-choice-icon"><Import size={24} /></span>
          <strong>{t('Import Existing')}</strong>
          <span>{t('Start from infrastructure code or reverse-engineer a live Azure resource group.')}</span>
          <div className="start-choice-actions">
            <button type="button" onClick={onImportTemplate}><FileCode2 size={15} /> {t('Template')}</button>
            <button type="button" onClick={onImportAzure}><Import size={15} /> {t('Azure')}</button>
          </div>
        </div>
      </div>
    </div>
  );
};

interface DeliverChooserProps {
  onShare: () => void;
  onBuild: () => void;
  onClose: () => void;
  isBuilding?: boolean;
}

export const DeliverChooser: React.FC<DeliverChooserProps> = ({ onShare, onBuild, onClose, isBuilding }) => {
  const { t } = useLanguage();
  return (
    <div className="deliver-chooser-backdrop" role="presentation" onClick={onClose}>
      <div className="deliver-chooser" role="dialog" aria-modal="true" aria-label={t('Choose a delivery outcome')} onClick={event => event.stopPropagation()}>
        <span className="start-chooser-kicker">{t('Step 4 · Share or Build')}</span>
        <h2>{t('What do you need next?')}</h2>
        <p>{t('Share the architecture for review, or create deployment guidance when the design is ready.')}</p>
        <div className="deliver-choice-grid">
          <button type="button" onClick={onShare}>
            <Presentation size={22} />
            <strong>{t('Share')}</strong>
            <span>{t('PowerPoint, Visio, Draw.io, PNG, HTML, JSON, and other review formats.')}</span>
          </button>
          <button type="button" onClick={onBuild} disabled={isBuilding}>
            <FileCode2 size={22} />
            <strong>{isBuilding ? t('Creating guide…') : t('Build')}</strong>
            <span>{t('Generate a deployment guide and starter infrastructure artifacts for review.')}</span>
          </button>
        </div>
        <button type="button" className="deliver-chooser-close" onClick={onClose}>{t('Cancel')}</button>
      </div>
    </div>
  );
};
