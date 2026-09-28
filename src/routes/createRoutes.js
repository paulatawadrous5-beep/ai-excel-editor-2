"use strict";

const express = require("express");
const { validateOperations } = require("../operations/operationValidator");
const { planCreate, summarizeOperations } = require("../services/aiService");
const { ExcelEngine } = require("../services/excelEngine");
const { runPipeline } = require("../services/applyPipeline");
const sessionStore = require("../services/sessionStore");
const fileUtils = require("../utils/fileUtils");

const router = express.Router();

router.post("/preview", express.json(), async (req, res, next) => {
  try {
    const instructions = (req.body.instructions || "").trim();
    if (!instructions) {
      return res.status(400).json({ error: "Please describe the workbook you'd like to create." });
    }
    if (instructions.length > 4000) {
      return res.status(400).json({ error: "That instruction is too long (max 4000 characters)." });
    }

    let operations;
    try {
      operations = await planCreate(instructions);
    } catch (err) {
      return next(err);
    }

    const { valid, issues } = validateOperations(operations);
    if (valid.length === 0) {
      return res.status(422).json({
        error: "The AI could not produce a valid workbook plan for that request. Try adding more detail.",
        details: issues.map((i) => i.message),
      });
    }

    const sessionId = sessionStore.create({ mode: "create", instructions, operations: valid });

    res.json({
      sessionId,
      summary: summarizeOperations(valid),
      operations: valid,
      warnings: issues.map((i) => i.message),
    });
  } catch (err) {
    next(err);
  }
});

router.post("/apply", express.json(), async (req, res, next) => {
  try {
    const { sessionId } = req.body;
    const session = sessionStore.get(sessionId);
    if (!session || session.mode !== "create") {
      return res.status(404).json({ error: "This session has expired. Please describe your workbook again." });
    }

    const requestedOps = Array.isArray(req.body.operations) ? req.body.operations : session.operations;
    const { valid: operations, issues } = validateOperations(requestedOps);
    if (operations.length === 0) {
      return res.status(422).json({ error: "No valid operations to apply.", details: issues.map((i) => i.message) });
    }

    const engine = ExcelEngine.blank();
    engine.workbook.addWorksheet("Sheet1");
    const { finalBuffer, report } = await runPipeline(engine, operations);

    if (report.verifiedCount + report.unverifiedCount === 0) {
      return res.status(422).json({
        error: "The workbook could not be generated from that description. Please add more detail and try again.",
        report,
      });
    }

    const fileId = fileUtils.newInternalId();
    await fileUtils.writeBuffer(fileId, ".xlsx", finalBuffer);
    sessionStore.remove(sessionId);

    res.json({
      downloadUrl: `/api/download/${fileId}?name=${encodeURIComponent("workbook.xlsx")}`,
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
