"use strict";

const multer = require("multer");
const logger = require("../utils/logger");
const { AiServiceError } = require("../services/aiService");

// eslint-disable-next-line no-unused-vars
function errorHandler(err, req, res, next) {
  if (err instanceof multer.MulterError) {
    const message = err.userMessage || (err.code === "LIMIT_FILE_SIZE" ? "That file is too large." : "Upload failed.");
    return res.status(400).json({ error: message });
  }

  if (err instanceof AiServiceError) {
    logger.warn("AI service error", err.message);
    return res.status(err.userFacing ? 422 : 500).json({ error: err.message });
  }

  if (err && err.userFacing) {
    return res.status(400).json({ error: err.message });
  }

  logger.error("Unhandled error", err);
  return res.status(500).json({
    error: "Something went wrong while processing your request. Please try again.",
  });
}

function notFoundHandler(req, res) {
  res.status(404).json({ error: "Not found." });
}

module.exports = { errorHandler, notFoundHandler };
