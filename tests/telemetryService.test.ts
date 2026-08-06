// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

import assert from 'node:assert/strict';
import test from 'node:test';
import {
  redactSensitiveTelemetry,
  redactSensitiveTelemetryValue,
} from '../src/services/telemetryService';

test('shared diagram capability tokens are removed from telemetry URLs', () => {
  const token = 'abcdefghijklmnopqrstuvwxyz0123456789ABCDEFG';
  const value = `https://example.test/api/diagrams/shared/${token}/comments?limit=20`;

  const redacted = redactSensitiveTelemetryValue(value);

  assert.equal(
    redacted,
    'https://example.test/api/diagrams/shared/[redacted]/comments?limit=20',
  );
  assert.doesNotMatch(redacted, new RegExp(token));
});

test('shared link hashes are removed from page-view telemetry', () => {
  const item = {
    name: 'Microsoft.ApplicationInsights.Pageview',
    baseType: 'PageviewData',
    baseData: {
      name: 'Azure Architecture Diagram Builder',
      uri: 'https://example.test/#share-secret-capability-token',
    },
  };

  redactSensitiveTelemetry(item);

  assert.equal(item.baseData.uri, 'https://example.test/#share-[redacted]');
});

test('dependency fields are redacted without changing unrelated telemetry', () => {
  const item = {
    name: 'Microsoft.ApplicationInsights.RemoteDependency',
    baseType: 'RemoteDependencyData',
    baseData: {
      name: 'GET /api/diagrams/shared/token-value',
      data: 'https://example.test/api/diagrams/shared/token-value',
      target: 'example.test',
      resultCode: '200',
    },
  };

  redactSensitiveTelemetry(item);

  assert.equal(item.baseData.name, 'GET /api/diagrams/shared/[redacted]');
  assert.equal(item.baseData.data, 'https://example.test/api/diagrams/shared/[redacted]');
  assert.equal(item.baseData.target, 'example.test');
  assert.equal(item.baseData.resultCode, '200');
});
