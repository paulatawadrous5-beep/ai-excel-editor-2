"use strict";

const express = require("express");
const path = require("path");
const cors = require("cors");
const helmet = require("helmet");

const config = require("./src/config");
const logger = require("./src/utils/logger");
const fileUtils = require("./src/utils/fileUtils");
const sessionStore = require("./src/services/sessionStore");
const { apiLimiter } = require("./src/middleware/rateLimiter");
const { errorHandler, notFoundHandler } = require("./src/middleware/errorHandler");

const editRoutes = require("./src/routes/editRoutes");
const createRoutes = require("./src/routes/createRoutes");
const downloadRoutes = require("./src/routes/downloadRoutes");

const app = express();

app.disable("x-powered-by");
app.use(
  helmet({
    contentSecurityPolicy: false, // the static frontend is same-origin and script-free of inline eval; keep simple for self-hosted deploys
  })
);
app.use(
  cors({
    origin: (origin, cb) => {
      if (!origin || config.allowedOrigins.includes("*") || config.allowedOrigins.includes(origin)) {
        return cb(null, true);
      }
      cb(new Error("Not allowed by CORS"));
    },
  })
);

app.use("/api", apiLimiter);
app.use("/api/edit", editRoutes);
app.use("/api/create", createRoutes);
app.use("/api/download", downloadRoutes);

app.get("/api/health", (req, res) => {
  res.json({ ok: true, aiConfigured: !!config.ai.apiKey });
});

app.use(express.static(path.join(__dirname, "public")));
app.get("*", (req, res, next) => {
  if (req.path.startsWith("/api/")) return next();
  res.sendFile(path.join(__dirname, "public", "index.html"));
});

app.use(notFoundHandler);
app.use(errorHandler);

fileUtils.ensureTempDir();
const sweepInterval = setInterval(() => {
  fileUtils.sweepExpired().catch((err) => logger.error("Temp file sweep failed", err));
  sessionStore.sweepExpired();
}, 5 * 60 * 1000);
sweepInterval.unref();

if (require.main === module) {
  app.listen(config.port, () => {
    logger.info(`AI Excel Editor listening on http://localhost:${config.port}`);
    logger.info(`.env ${config.envFileLoaded ? "loaded from" : "NOT found at"} ${config.envPath}`);
    logger.info(`AI model: ${config.ai.model} | API key ${config.ai.apiKey ? "loaded" : "MISSING"}`);
    if (!config.ai.apiKey) {
      logger.warn("ANTHROPIC_API_KEY is not set - AI-powered requests will fail until you configure .env and restart");
    }
  });
}

module.exports = app;

