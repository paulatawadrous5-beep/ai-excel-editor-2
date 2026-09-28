"use strict";

/**
 * Minimal structured logger. Keeps technical details out of user-facing
 * responses while still giving developers something useful in stdout/stderr.
 */
function timestamp() {
  return new Date().toISOString();
}

const logger = {
  info(message, meta) {
    console.log(`[${timestamp()}] INFO  ${message}`, meta !== undefined ? meta : "");
  },
  warn(message, meta) {
    console.warn(`[${timestamp()}] WARN  ${message}`, meta !== undefined ? meta : "");
  },
  error(message, err) {
    if (err && err.stack) {
      console.error(`[${timestamp()}] ERROR ${message}: ${err.message}`);
      console.error(err.stack);
    } else {
      console.error(`[${timestamp()}] ERROR ${message}`, err !== undefined ? err : "");
    }
  },
};

module.exports = logger;
