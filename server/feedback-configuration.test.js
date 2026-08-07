// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

const assert = require('node:assert/strict');
const { test } = require('node:test');
const {
  createArchivedFeedbackContact,
  hasFeedbackArchiveConfiguration,
  hasFeedbackContactConfiguration,
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

test('follow-up contact requires an explicit enablement and email delivery', () => {
  const emailConfiguration = {
    emailEndpoint: 'https://example.communication.azure.com',
    emailSender: 'sender@example.com',
    emailRecipient: 'recipient@example.com',
  };
  assert.equal(hasFeedbackContactConfiguration(emailConfiguration), false);
  assert.equal(hasFeedbackContactConfiguration({
    ...emailConfiguration,
    contactEnabled: true,
  }), true);
});

test('feedback archives retain consent metadata without the contact address', () => {
  assert.deepEqual(createArchivedFeedbackContact({ consent: false }), { consent: false });
  assert.deepEqual(createArchivedFeedbackContact({
    consent: true,
    email: 'person@example.com',
    consentAt: '2026-08-05T20:00:00.000Z',
    expiresAt: '2027-02-01T20:00:00.000Z',
    followUpStatus: 'new',
  }), {
    consent: true,
    consentAt: '2026-08-05T20:00:00.000Z',
    expiresAt: '2027-02-01T20:00:00.000Z',
    followUpStatus: 'new',
  });
});
