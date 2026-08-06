// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

export type WorkflowStepStatus = 'complete' | 'current' | 'pending' | 'busy';
export type WorkflowStepId = 'generate' | 'validate' | 'cost' | 'deploy';

export interface WorkflowProgress {
  serviceCount: number;
  validationScore: number | null;
  hasCostData: boolean;
  hasDeploymentGuide: boolean;
  isValidating: boolean;
  isGeneratingGuide: boolean;
}

export function getWorkflowStepStatuses({
  serviceCount,
  validationScore,
  hasCostData,
  hasDeploymentGuide,
  isValidating,
  isGeneratingGuide,
}: WorkflowProgress): Record<WorkflowStepId, WorkflowStepStatus> {
  const generated = serviceCount > 0;
  const validated = validationScore !== null;
  const costed = validated && hasCostData;

  return {
    generate: generated ? 'complete' : 'current',
    validate: isValidating
      ? 'busy'
      : validated
        ? 'complete'
        : generated
          ? 'current'
          : 'pending',
    cost: costed ? 'complete' : validated ? 'current' : 'pending',
    deploy: isGeneratingGuide
      ? 'busy'
      : hasDeploymentGuide
        ? 'complete'
        : costed
          ? 'current'
          : 'pending',
  };
}
