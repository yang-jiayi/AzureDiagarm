// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

import {
  Check,
  ChevronDown,
  ChevronUp,
  CircleDollarSign,
  Loader2,
  Rocket,
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
  onDeploy: () => void;
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
  onDeploy,
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
    id: WorkflowStepId;
    icon: LucideIcon;
    label: string;
    detail: string;
    action: () => void;
    disabled: boolean;
  }> = [
    {
      id: 'generate',
      icon: Sparkles,
      label: localize(language, { en: 'Generate', ja: '生成' }),
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
      detail: validationScore !== null
        ? localize(language, {
            en: `WAF score ${validationScore}`,
            ja: `WAF スコア ${validationScore}`,
          })
        : localize(language, { en: 'Check architecture readiness', ja: '設計の準備状況を確認' }),
      action: onValidate,
      disabled: statuses.validate === 'pending' || statuses.validate === 'busy',
    },
    {
      id: 'cost',
      icon: CircleDollarSign,
      label: localize(language, { en: 'Cost', ja: 'コスト' }),
      detail: hasCostData
        ? monthlyCostLabel
        : localize(language, { en: 'Waiting for pricing data', ja: '料金データを待機中' }),
      action: onReviewCost,
      disabled: statuses.cost === 'pending' || !hasCostData,
    },
    {
      id: 'deploy',
      icon: Rocket,
      label: localize(language, { en: 'Deploy', ja: 'デプロイ' }),
      detail: hasDeploymentGuide
        ? localize(language, { en: 'Deployment guide ready', ja: 'デプロイ ガイド準備完了' })
        : localize(language, { en: 'Generate implementation guidance', ja: '実装ガイドを生成' }),
      action: onDeploy,
      disabled: statuses.deploy === 'pending' || statuses.deploy === 'busy',
    },
  ];
  const activeStep = steps.find(step => (
    statuses[step.id] === 'current' || statuses[step.id] === 'busy'
  )) || steps.find(step => statuses[step.id] !== 'complete') || steps[steps.length - 1];
  const activeStatus = statuses[activeStep.id];

  if (collapsed) {
    return (
      <nav
        className="workflow-stepper workflow-stepper--collapsed"
        aria-label={localize(language, {
          en: 'Architecture delivery workflow',
          ja: 'アーキテクチャ提供ワークフロー',
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
        en: 'Architecture delivery workflow',
        ja: 'アーキテクチャ提供ワークフロー',
      })}
    >
      {onToggleCollapsed && (
        <button
          type="button"
          className="workflow-stepper-collapse"
          onClick={onToggleCollapsed}
          aria-expanded="true"
        >
          <ChevronUp size={16} aria-hidden="true" />
          {localize(language, { en: 'Collapse workflow', ja: 'ワークフローを縮小' })}
        </button>
      )}
      <ol>
        {steps.map((step, index) => {
          const status = statuses[step.id];
          const Icon = step.icon;
          const statusId = `workflow-step-${step.id}-status`;
          return (
            <li key={step.id} className={`workflow-stepper-item is-${status}`}>
              <button
                type="button"
                onClick={step.action}
                disabled={step.disabled}
                aria-current={status === 'current' || status === 'busy' ? 'step' : undefined}
                aria-describedby={statusId}
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
                  <small>{step.detail}</small>
                </span>
                <span id={statusId} className="sr-only">{statusLabel(status)}</span>
              </button>
            </li>
          );
        })}
      </ol>
    </nav>
  );
}
