// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

function createGracefulShutdown(server, options = {}) {
  const logger = options.logger || console;
  const timeoutMs = Math.max(1_000, Number(options.timeoutMs) || 25_000);
  const exit = options.exit || ((code) => process.exit(code));
  let shuttingDown = false;

  return (signal) => {
    if (shuttingDown) return false;
    shuttingDown = true;
    logger.info(`[server] ${signal} received; draining active requests.`);

    const forceTimer = setTimeout(() => {
      logger.error(`[server] Graceful shutdown exceeded ${timeoutMs}ms; closing active connections.`);
      server.closeAllConnections?.();
      exit(1);
    }, timeoutMs);
    forceTimer.unref?.();

    const drained = Promise.resolve().then(() => options.drain?.());
    // Observe failures immediately, even while HTTP connections are draining.
    drained.catch(() => {});
    server.close(async (error) => {
      try { await drained; }
      catch (drainError) {
        logger.error('[server] Background work could not be drained safely.');
        error ||= drainError;
      }
      clearTimeout(forceTimer);
      if (error) {
        logger.error('[server] Failed to close cleanly:', error);
        exit(1);
        return;
      }
      logger.info('[server] Active requests drained.');
      exit(0);
    });
    server.closeIdleConnections?.();
    return true;
  };
}

module.exports = { createGracefulShutdown };
