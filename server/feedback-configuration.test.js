// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

const assert = require('node:assert/strict');
const { test } = require('node:test');
const {
  hasFeedbackArchiveConfiguration,
  hasFeedbackDeliveryConfiguration,
  hasFeedbackEmailConfiguration,
} = require('./feedback-configuration');

test('feedback email delivery requires every email setting', () => {
  assert.equal(hasFeedbackEmailConfiguration({
    emailEndpoint: 'https://example.communication.azure.com',
    emailSender: 'sender@example.com',
  }), false);
  assert.equal(hasFeedbackEmailConfiguration({
    emailEndpoint: 'https://example.communication.azure.com',
    emailSender: 'sender@example.com',
    emailRecipient: 'recipient@example.com',
  }), true);
});

test('feedback configuration checks do not initialize asynchronous storage clients', () => {
  assert.equal(hasFeedbackArchiveConfiguration({ tablesEndpoint: 'https://example.table.core.windows.net' }), true);
  assert.equal(hasFeedbackArchiveConfiguration({ cosmosEndpoint: 'https://example.documents.azure.com' }), true);
  assert.equal(hasFeedbackArchiveConfiguration({}), false);
  assert.equal(hasFeedbackDeliveryConfiguration({}), false);
});
