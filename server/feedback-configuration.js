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

module.exports = {
  hasFeedbackArchiveConfiguration,
  hasFeedbackDeliveryConfiguration,
  hasFeedbackEmailConfiguration,
};
