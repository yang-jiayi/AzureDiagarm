// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

/**
 * Express 4 does not forward rejected promises returned by `async` route
 * handlers to the error middleware, so an unhandled rejection escapes to the
 * process. Node terminates the process on an unhandled rejection by default,
 * and start.sh stops the whole container when the token server exits — so a
 * single failed await inside a route would restart the container for every
 * user. Wrapping handlers routes those rejections into Express instead.
 */
function asyncHandler(handler) {
  return function wrapped(req, res, next) {
    try {
      Promise.resolve(handler(req, res, next)).catch(next);
    } catch (error) {
      next(error);
    }
  };
}

function createErrorHandler(logger = console) {
  return function handleError(error, _req, res, next) {
    // Preserve client-error statuses that Express's default finalhandler used
    // to surface (express.json() rejects malformed bodies with 400 and
    // oversized bodies with 413). Reporting those as 500 would mislead clients
    // that only retry on server errors and would bury real faults in the log.
    const declared = Number(error?.status ?? error?.statusCode);
    const status = Number.isInteger(declared) && declared >= 400 && declared <= 599
      ? declared
      : 500;

    if (status >= 500) {
      logger.error('[server] unhandled request error:', error?.stack || error?.message || error);
    } else {
      logger.warn('[server] rejected request:', status, error?.message || error);
    }

    if (res.headersSent) return next(error);
    res.status(status).json({
      error: status >= 500 ? 'Internal server error' : 'Invalid request',
    });
  };
}

module.exports = {
  asyncHandler,
  createErrorHandler,
};
