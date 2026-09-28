"use strict";

require("dotenv").config();
const path = require("path");

function parseOrigins(value) {
  if (!value) return ["http://localhost:3000"];
  return value.split(",").map((s) => s.trim()).filter(Boolean);
}

const config = {
  port: parseInt(process.env.PORT, 10) || 3000,
  nodeEnv: process.env.NODE_ENV || "development",

  ai: {
    apiKey: process.env.ANTHROPIC_API_KEY || "",
    model: process.env.AI_MODEL || "claude-sonnet-5",
    apiBase: process.env.ANTHROPIC_API_BASE || "https://api.anthropic.com/v1/messages",
  },

  allowedOrigins: parseOrigins(process.env.ALLOWED_ORIGINS),

  upload: {
    maxSizeBytes: (parseInt(process.env.MAX_UPLOAD_MB, 10) || 15) * 1024 * 1024,
    allowedExtensions: [".xlsx"],
  },

  tempDir: path.resolve(process.cwd(), process.env.TEMP_DIR || "./tmp"),
  fileTtlMs: (parseInt(process.env.FILE_TTL_MINUTES, 10) || 30) * 60 * 1000,

  rateLimit: {
    windowMs: (parseInt(process.env.RATE_LIMIT_WINDOW_MINUTES, 10) || 15) * 60 * 1000,
    max: parseInt(process.env.RATE_LIMIT_MAX_REQUESTS, 10) || 100,
  },
};

module.exports = config;
