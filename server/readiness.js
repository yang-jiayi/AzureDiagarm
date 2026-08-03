// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

function createReadinessHandler(options = {}) {
  const isShuttingDown = options.isShuttingDown || (() => false);
  const isConfigured = options.isConfigured || (() => true);

  return (_req, res) => {
    if (isShuttingDown() || !isConfigured()) {
      return res.status(503).type('text/plain').send('not ready\n');
    }
    return res.type('text/plain').send('ready\n');
  };
}

module.exports = { createReadinessHandler };
