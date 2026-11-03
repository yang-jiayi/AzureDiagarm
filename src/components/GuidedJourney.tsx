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
  onGuidedChat: (brief: string) => void;
  onGenerate: (brief: string) => void;
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
  const { language } = useLanguage();
  const [brief, setBrief] = React.useState('');
  const example = localize(language, {
    en: 'A customer-facing web app with a secure API, relational data, and monitoring',
    ja: '安全なAPI、リレーショナル データ、監視を備えた顧客向けWebアプリ',
  });
  return (
    <div
      className="start-chooser"
      role="region"
      aria-label={localize(language, { en: 'Start an architecture', ja: 'アーキテクチャを作成' })}
    >
      <div className="start-chooser-heading">
        <h2>{localize(language, { en: 'What would you like to build?', ja: '何を作成しますか？' })}</h2>
        <p>{localize(language, {
          en: 'Start with a short brief. You can review and edit it before sending.',
          ja: 'まず要件を簡潔に入力してください。送信前に確認・編集できます。',
        })}</p>
      </div>
      <div className="start-brief">
        <label htmlFor="start-architecture-brief">
          {localize(language, { en: 'Architecture brief', ja: 'アーキテクチャの要件' })}
        </label>
        <textarea
          id="start-architecture-brief"
          value={brief}
          onChange={(event) => setBrief(event.target.value)}
          placeholder={localize(language, {
            en: 'Describe the workload, who uses it, and what it needs to do…',
            ja: 'ワークロード、利用者、必要な機能を説明してください…',
          })}
          rows={3}
        />
        <button
          type="button"
          className="start-brief-example"
          onClick={() => setBrief(example)}
        >
          {localize(language, { en: 'Use an editable example', ja: '編集できる例を入力' })}
          <span>{example}</span>
        </button>
      </div>
      <div className="start-choice-grid">
        <button type="button" className="start-choice recommended" onClick={() => onGuidedChat(brief.trim())}>
          <span className="start-choice-icon"><MessageSquare size={21} /></span>
          <strong>{localize(language, { en: 'Continue in Guided Chat', ja: 'ガイド付きチャットで続ける' })}</strong>
          <span>{localize(language, {
            en: 'Your brief is placed in the input. Nothing is sent until you choose Send.',
            ja: '要件が入力欄に設定されます。送信を押すまで送信されません。',
          })}</span>
        </button>
      </div>
      <div className="start-choice-advanced">
        <span>{localize(language, { en: 'Need image input or more output control?', ja: '画像入力や出力の詳細設定が必要ですか？' })}</span>
        <button type="button" onClick={() => onGenerate(brief.trim())}>
          <span className="start-choice-icon"><Sparkles size={21} /></span>
          <strong>{localize(language, { en: 'Image or detailed generation', ja: '画像・詳細設定で生成' })}</strong>
          <span>{localize(language, {
            en: 'Keep this brief while adding an image or choosing Topology, Blueprint, or both.',
            ja: 'この要件を引き継ぎ、画像やTopology、Blueprint、両方の出力を選べます。',
          })}</span>
        </button>
      </div>
      <div className="start-chooser-secondary">
        <span>{localize(language, { en: 'Start from an existing design', ja: '既存の設計から開始' })}</span>
        <button type="button" onClick={onBrowseTemplates}>
          <LayoutTemplate size={15} aria-hidden="true" />
          {localize(language, { en: 'Browse templates', ja: 'テンプレートを見る' })}
        </button>
        <button type="button" onClick={onImportTemplate}>
          <FileCode2 size={15} aria-hidden="true" />
          {localize(language, { en: 'Import Bicep, Terraform, or ARM', ja: 'Bicep、Terraform、ARMを取り込む' })}
        </button>
        <button type="button" onClick={onImportAzure}>
          <Import size={15} aria-hidden="true" />
          {localize(language, { en: 'Import Azure resource group', ja: 'Azureリソース グループを取り込む' })}
        </button>
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
