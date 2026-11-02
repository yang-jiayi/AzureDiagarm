// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

import assert from 'node:assert/strict';
import test from 'node:test';
import { getWorkflowStepStatuses } from '../src/components/workflowStepperState';

test('the workflow starts at generation and gates later recommendations', () => {
  assert.deepEqual(
    getWorkflowStepStatuses({
      serviceCount: 0,
      validationScore: null,
      hasCostData: false,
      hasDeploymentGuide: false,
      isValidating: false,
      isGeneratingGuide: false,
    }),
    {
      generate: 'current',
      validate: 'pending',
      cost: 'pending',
      output: 'pending',
    },
  );
});

test('a generated diagram advances the recommended action to validation', () => {
  assert.deepEqual(
    getWorkflowStepStatuses({
      serviceCount: 4,
      validationScore: null,
      hasCostData: true,
      hasDeploymentGuide: false,
      isValidating: false,
      isGeneratingGuide: false,
    }),
    {
      generate: 'complete',
      validate: 'current',
      cost: 'pending',
      output: 'pending',
    },
  );
});

test('validation and pricing completion advance the workflow to output', () => {
  assert.deepEqual(
    getWorkflowStepStatuses({
      serviceCount: 4,
      validationScore: 82,
      hasCostData: true,
      hasDeploymentGuide: false,
      isValidating: false,
      isGeneratingGuide: false,
    }),
    {
      generate: 'complete',
      validate: 'complete',
      cost: 'complete',
      output: 'current',
    },
  );
});

test('long-running validation and output actions expose busy states', () => {
  assert.equal(
    getWorkflowStepStatuses({
      serviceCount: 4,
      validationScore: null,
      hasCostData: false,
      hasDeploymentGuide: false,
      isValidating: true,
      isGeneratingGuide: false,
    }).validate,
    'busy',
  );
  assert.equal(
    getWorkflowStepStatuses({
      serviceCount: 4,
      validationScore: 82,
      hasCostData: true,
      hasDeploymentGuide: false,
      isValidating: false,
      isGeneratingGuide: true,
    }).output,
    'busy',
  );
});
