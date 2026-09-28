"use strict";

const express = require("express");
const { upload, verifyMagicBytes } = require("../middleware/upload");
const { loadWorkbook, analyzeWorkbook } = require("../services/workbookAnalyzer");
const { validateOperations } = require("../operations/operationValidator");
const { planEdit, summarizeOperations } = require("../services/aiService");
const { ExcelEngine } = require("../services/excelEngine");
const { runPipeline } = require("../services/applyPipeline");
const sessionStore = require("../services/sessionStore");
const fileUtils = require("../utils/fileUtils");
const logger = require("../utils/logger");

const router = express.Router();

router.post("/preview", upload.single("file"), verifyMagicBytes, async (req, res, next) => {
  try {
    const instructions = (req.body.instructions || "").trim();
    if (!instructions) {
      return res.status(400).json({ error: "Please describe the changes you'd like to make." });
    }
    if (instructions.length > 4000) {
      return res.status(400).json({ error: "That instruction is too long (max 4000 characters)." });
    }

    const workbook = await loadWorkbook(req.file.buffer);
    const summary = analyzeWorkbook(workbook);

    let operations;
    try {
      operations = await planEdit(instructions, summary);
    } catch (err) {
      return next(err);
    }

    const { valid, issues } = validateOperations(operations);
    if (valid.length === 0) {
      return res.status(422).json({
        error: "The AI could not produce a valid set of changes for that request. Try rephrasing it, or be more specific.",
        details: issues.map((i) => i.message),
      });
    }

    const sessionId = sessionStore.create({
      mode: "edit",
      instructions,
      operations: valid,
      originalName: fileUtils.sanitizeDisplayName(req.file.originalname),
    });
    await fileUtils.writeBuffer(sessionId, ".xlsx", req.file.buffer);

    res.json({
      sessionId,
      summary: summarizeOperations(valid),
      operations: valid,
      warnings: issues.map((i) => i.message),
      workbookInfo: { sheetNames: summary.sheetNames },
    });
  } catch (err) {
    next(err);
  }
});

router.post("/apply", express.json(), async (req, res, next) => {
  try {
    const { sessionId } = req.body;
    const session = sessionStore.get(sessionId);
    if (!session || session.mode !== "edit") {
      return res.status(404).json({ error: "This editing session has expired. Please upload your file again." });
    }

    const requestedOps = Array.isArray(req.body.operations) ? req.body.operations : session.operations;
    const { valid: operations, issues } = validateOperations(requestedOps);
    if (operations.length === 0) {
      return res.status(422).json({ error: "No valid operations to apply.", details: issues.map((i) => i.message) });
    }

    const originalBuffer = await fileUtils.readBuffer(sessionId, ".xlsx");
    const engine = await ExcelEngine.fromBuffer(originalBuffer);
    const { finalBuffer, report } = await runPipeline(engine, operations);

    if (report.verifiedCount + report.unverifiedCount === 0) {
      return res.status(422).json({
        error: "None of the requested changes could be safely completed. Please review the instructions and try again.",
        report,
      });
    }

    const fileId = fileUtils.newInternalId();
    await fileUtils.writeBuffer(fileId, ".xlsx", finalBuffer);
    sessionStore.remove(sessionId);
    await fileUtils.removeIfExists(sessionId, ".xlsx");

    const baseName = (session.originalName || "workbook.xlsx").replace(/\.xlsx$/i, "");
    res.json({
      downloadUrl: `/api/download/${fileId}?name=${encodeURIComponent(baseName + "-edited.xlsx")}`,
      status: report.status,
      verifiedCount: report.verifiedCount,
      failedCount: report.failedCount,
      unverifiedCount: report.unverifiedCount,
      failed: report.failed,
      unverified: report.unverified,
    });
  } catch (err) {
    next(err);
  }
});

module.exports = router;
