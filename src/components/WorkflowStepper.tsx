// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

import {
  Check,
  ChevronDown,
  ChevronUp,
  CircleDollarSign,
  Download,
  Loader2,
  ShieldCheck,
  Sparkles,
  type LucideIcon,
} from 'lucide-react';
import { useLanguage } from '../i18n/LanguageContext';
import { localize } from '../i18n/localization';
import {
  getWorkflowStepStatuses,
  type WorkflowProgress,
  type WorkflowStepId,
  type WorkflowStepStatus,
} from './workflowStepperState';
import './WorkflowStepper.css';

interface WorkflowStepperProps extends WorkflowProgress {
  monthlyCostLabel: string;
  collapsed?: boolean;
  onToggleCollapsed?: () => void;
  onGenerate: () => void;
  onValidate: () => void;
  onReviewCost: () => void;
  onOutput: () => void;
}

export default function WorkflowStepper({
  serviceCount,
  validationScore,
  hasCostData,
  hasDeploymentGuide,
  isValidating,
  isGeneratingGuide,
  monthlyCostLabel,
  collapsed = false,
  onToggleCollapsed,
  onGenerate,
  onValidate,
  onReviewCost,
  onOutput,
}: WorkflowStepperProps) {
  const { language } = useLanguage();
  const statuses = getWorkflowStepStatuses({
    serviceCount,
    validationScore,
    hasCostData,
    hasDeploymentGuide,
    isValidating,
    isGeneratingGuide,
  });

  const statusLabel = (status: WorkflowStepStatus) => localize(language, {
    en: status === 'complete'
      ? 'Complete'
      : status === 'busy'
        ? 'In progress'
        : status === 'current'
          ? 'Next recommended step'
          : 'Waiting for an earlier step',
    ja: status === 'complete'
      ? '完了'
      : status === 'busy'
        ? '処理中'
        : status === 'current'
          ? '次の推奨ステップ'
          : '前のステップを待機',
  });

  const steps: Array<{
    id: Exclude<WorkflowStepId, 'cost'>;
    icon: LucideIcon;
    label: string;
    detail: string;
    value?: string;
    action: () => void;
    disabled: boolean;
    secondaryAction?: {
      label: string;
      detail: string;
      action: () => void;
      disabled: boolean;
    };
  }> = [
    {
      id: 'generate',
      icon: Sparkles,
      label: localize(language, { en: 'Generate', ja: '生成' }),
      value: serviceCount > 0
        ? localize(language, { en: `${serviceCount} services`, ja: `${serviceCount} サービス` })
        : undefined,
      detail: serviceCount > 0
        ? localize(language, {
            en: `${serviceCount} services on canvas`,
            ja: `キャンバス上に ${serviceCount} サービス`,
          })
        : localize(language, { en: 'Describe or add services', ja: '要件を入力またはサービスを追加' }),
      action: onGenerate,
      disabled: false,
    },
    {
      id: 'validate',
      icon: ShieldCheck,
      label: localize(language, { en: 'Validate', ja: '検証' }),
      value: validationScore !== null
        ? localize(language, {
            en: `Well-Architected ${validationScore}`,
            ja: `Well-Architected ${validationScore}`,
          })
        : undefined,
      detail: validationScore !== null
        ? localize(language, {
            en: `Azure Well-Architected design review score: ${validationScore}. Cost is also a review dimension.`,
            ja: `Azure Well-Architected 設計レビュー スコア: ${validationScore}。コストもレビュー項目です。`,
          })
        : localize(language, {
            en: 'Review architecture quality, risk, and estimated cost',
            ja: '設計の品質、リスク、概算コストを確認',
          }),
      action: onValidate,
      disabled: statuses.validate === 'pending' || statuses.validate === 'busy',
      secondaryAction: {
        label: localize(language, { en: 'Cost', ja: 'コスト' }),
        detail: localize(language, {
          en: 'Review cost estimates as part of architecture validation',
          ja: 'アーキテクチャ検証の一環としてコスト見積を確認',
        }),
        action: onReviewCost,
        disabled: statuses.cost === 'pending' || !hasCostData,
      },
    },
    {
      id: 'output',
      icon: Download,
      label: localize(language, { en: 'Output', ja: '出力' }),
      detail: hasDeploymentGuide
        ? localize(language, {
            en: 'Guide ready. This app does not deploy to Azure.',
            ja: 'ガイドの準備ができました。このアプリは Azure にデプロイしません。',
          })
        : localize(language, {
            en: 'Export or create deployment guidance and starter IaC. This app does not deploy to Azure.',
            ja: 'エクスポート、デプロイ手順、IaC のひな形を作成します。このアプリは Azure にデプロイしません。',
          }),
      action: onOutput,
      disabled: statuses.output === 'pending' || statuses.output === 'busy',
    },
  ];
  const stepStatus = (step: (typeof steps)[number]): WorkflowStepStatus => (
    step.id === 'validate'
      && statuses.validate === 'complete'
      && statuses.cost !== 'complete'
      ? statuses.cost
      : statuses[step.id]
  );
  const activeStep = steps.find(step => (
    stepStatus(step) === 'current' || stepStatus(step) === 'busy'
  )) || steps.find(step => stepStatus(step) !== 'complete') || steps[steps.length - 1];
  const activeStatus = stepStatus(activeStep);

  if (collapsed) {
    return (
      <nav
        className="workflow-stepper workflow-stepper--collapsed"
        aria-label={localize(language, {
          en: 'Architecture workflow: create, review, and output',
          ja: 'アーキテクチャ ワークフロー: 作成、検証、出力',
        })}
      >
        <button
          type="button"
          className="workflow-stepper-summary"
          onClick={onToggleCollapsed}
          aria-expanded="false"
        >
          <span className={`workflow-stepper-index is-${activeStatus}`} aria-hidden="true">
            {activeStatus === 'busy'
              ? <Loader2 size={15} className="workflow-stepper-spinner" />
              : steps.findIndex(step => step.id === activeStep.id) + 1}
          </span>
          <span className="workflow-stepper-summary-copy">
            <small>{localize(language, { en: 'Next step', ja: '次のステップ' })}</small>
            <strong>{activeStep.label}</strong>
          </span>
          <span className="workflow-stepper-summary-detail">{activeStep.detail}</span>
          <ChevronDown size={18} aria-hidden="true" />
        </button>
      </nav>
    );
  }

  return (
    <nav
      className="workflow-stepper"
      aria-label={localize(language, {
        en: 'Architecture workflow: create, review, and output',
        ja: 'アーキテクチャ ワークフロー: 作成、検証、出力',
      })}
    >
      {onToggleCollapsed && (
        <button
          type="button"
          className="workflow-stepper-collapse"
          onClick={onToggleCollapsed}
          aria-expanded="true"
          aria-label={localize(language, { en: 'Collapse workflow', ja: 'ワークフローを縮小' })}
          title={localize(language, { en: 'Collapse workflow', ja: 'ワークフローを縮小' })}
        >
          <ChevronUp size={16} aria-hidden="true" />
        </button>
      )}
      <ol>
        {steps.map((step, index) => {
          const status = stepStatus(step);
          const Icon = step.icon;
          const statusId = `workflow-step-${step.id}-status`;
          const costStatusId = 'workflow-step-cost-status';
          return (
            <li key={step.id} className={`workflow-stepper-item is-${status}`}>
              <div className={`workflow-stepper-actions${step.secondaryAction ? ' workflow-stepper-actions--review' : ''}`}>
                <button
                  type="button"
                  className="workflow-stepper-primary"
                  onClick={step.action}
                  disabled={step.disabled}
                  aria-label={step.label}
                  aria-current={status === 'current' || status === 'busy' ? 'step' : undefined}
                  aria-describedby={statusId}
                  title={`${step.label}: ${step.detail}`}
                >
                  <span className="workflow-stepper-index" aria-hidden="true">
                    {status === 'complete'
                      ? <Check size={15} />
                      : status === 'busy'
                        ? <Loader2 size={15} className="workflow-stepper-spinner" />
                        : index + 1}
                  </span>
                  <Icon className="workflow-stepper-icon" size={18} aria-hidden="true" />
                  <span className="workflow-stepper-copy">
                    <strong>{step.label}</strong>
                    {step.value && <span className="workflow-stepper-value" aria-hidden="true">{step.value}</span>}
                  </span>
                  <span id={statusId} className="sr-only">{statusLabel(status)}. {step.detail}</span>
                </button>
                {step.secondaryAction && (
                  <button
                    type="button"
                    className={`workflow-stepper-cost is-${statuses.cost}`}
                    onClick={step.secondaryAction.action}
                    disabled={step.secondaryAction.disabled}
                    aria-label={step.secondaryAction.label}
                    aria-describedby={costStatusId}
                    title={step.secondaryAction.detail}
                  >
                    <CircleDollarSign size={16} aria-hidden="true" />
                    <span>{step.secondaryAction.label}</span>
                    <span id={costStatusId} className="sr-only">
                      {statusLabel(statuses.cost)}. {step.secondaryAction.detail}
                      {hasCostData ? `. ${monthlyCostLabel}` : ''}
                    </span>
                  </button>
                )}
              </div>
            </li>
          );
        })}
      </ol>
    </nav>
  );
}
