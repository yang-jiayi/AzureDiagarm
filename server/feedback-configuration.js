// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

function hasFeedbackEmailConfiguration(configuration) {
  return Boolean(
    configuration.emailEndpoint
    && configuration.emailSender
    && configuration.emailRecipient
  );
}

function hasFeedbackArchiveConfiguration(configuration) {
  return Boolean(configuration.tablesEndpoint || configuration.cosmosEndpoint);
}

function hasFeedbackDeliveryConfiguration(configuration) {
  return hasFeedbackEmailConfiguration(configuration)
    || hasFeedbackArchiveConfiguration(configuration);
}

function hasFeedbackContactConfiguration(configuration) {
  return configuration.contactEnabled === true
    && hasFeedbackEmailConfiguration(configuration);
}

function createArchivedFeedbackContact(contact) {
  if (!contact?.consent) return { consent: false };
  return {
    consent: true,
    consentAt: contact.consentAt,
    expiresAt: contact.expiresAt,
    followUpStatus: contact.followUpStatus,
  };
}

module.exports = {
  createArchivedFeedbackContact,
  hasFeedbackArchiveConfiguration,
  hasFeedbackContactConfiguration,
  hasFeedbackDeliveryConfiguration,
  hasFeedbackEmailConfiguration,
};
