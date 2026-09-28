"use strict";

const path = require("path");
const dotenv = require("dotenv");

// Always load <project root>/.env
const envPath = path.resolve(__dirname, "..", ".env");
const envResult = dotenv.config({ path: envPath, override: true });

/** Trims whitespace/CR and surrounding quotes. */
function clean(value) {
  if (value === undefined || value === null) return "";

  let s = String(value).trim();

  if (
    (s.startsWith('"') && s.endsWith('"')) ||
    (s.startsWith("'") && s.endsWith("'"))
  ) {
    s = s.slice(1, -1).trim();
  }

  return s;
}

function parseOrigins(value) {
  if (!value) return ["http://localhost:3000"];

  return value
    .split(",")
    .map((s) => s.trim())
    .filter(Boolean);
}

const config = {
  port: parseInt(process.env.PORT, 10) || 3000,
  nodeEnv: process.env.NODE_ENV || "development",

  envPath,
  envFileLoaded: !envResult.error,

  ai: {
    provider: "gemini",

    apiKey: clean(process.env.GEMINI_API_KEY),

    // Current Gemini Flash model suitable for this application.
    model: clean(process.env.AI_MODEL) || "gemini-3.8-flash",

    apiBase:
      clean(process.env.GEMINI_API_BASE) ||
      "https://generativelanguage.googleapis.com/v1beta",
  },

  allowedOrigins: parseOrigins(process.env.ALLOWED_ORIGINS),

  upload: {
    maxSizeBytes:
      (parseInt(process.env.MAX_UPLOAD_MB, 10) || 15) * 1024 * 1024,

    allowedExtensions: [".xlsx"],
  },

  tempDir: path.resolve(
    process.cwd(),
    process.env.TEMP_DIR || "./tmp"
  ),

  fileTtlMs:
    (parseInt(process.env.FILE_TTL_MINUTES, 10) || 30) *
    60 *
    1000,

  rateLimit: {
    windowMs:
      (parseInt(process.env.RATE_LIMIT_WINDOW_MINUTES, 10) || 15) *
      60 *
      1000,

    max:
      parseInt(process.env.RATE_LIMIT_MAX_REQUESTS, 10) || 100,
  },
};

module.exports = config;