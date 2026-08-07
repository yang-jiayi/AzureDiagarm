// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

import assert from 'node:assert/strict';
import test from 'node:test';
import {
  getConnectionPresentation,
  inferConnectionType,
  normalizeConnectionType,
} from '../src/utils/edgePresentation';

test('connection presentation normalizes generated semantic aliases', () => {
  assert.equal(normalizeConnectionType('asynchronous'), 'async');
  assert.equal(normalizeConnectionType('identity'), 'security');
  assert.equal(normalizeConnectionType('observability'), 'telemetry');
  assert.equal(normalizeConnectionType('unexpected'), 'sync');
});

test('security and telemetry connections have distinct export-safe styling', () => {
  const security = getConnectionPresentation('security');
  const telemetry = getConnectionPresentation('telemetry');

  assert.notEqual(security.stroke, telemetry.stroke);
  assert.ok(security.strokeDasharray);
  assert.ok(telemetry.strokeDasharray);
  assert.equal(security.baseFlowAnimated, false);
  assert.equal(telemetry.baseFlowAnimated, true);
});

test('semantic connection inference recognizes identity and observability edges', () => {
  assert.equal(inferConnectionType({
    label: 'Validate token',
    toCategory: 'identity',
    toName: 'Microsoft Entra ID',
  }), 'security');
  assert.equal(inferConnectionType({
    label: 'Send telemetry',
    toCategory: 'monitor',
    toName: 'Log Analytics',
  }), 'telemetry');
});

test('semantic connection inference falls back to authored line styles', () => {
  assert.equal(inferConnectionType({ style: 'dotted' }), 'optional');
  assert.equal(inferConnectionType({ style: 'dashed' }), 'async');
  assert.equal(inferConnectionType({ label: 'HTTPS request' }), 'sync');
});
