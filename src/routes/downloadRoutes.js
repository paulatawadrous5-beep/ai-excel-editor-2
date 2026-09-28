"use strict";

const express = require("express");
const fileUtils = require("../utils/fileUtils");

const router = express.Router();
const UUID_RE = /^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i;

router.get("/:fileId", async (req, res, next) => {
  try {
    const { fileId } = req.params;
    if (!UUID_RE.test(fileId)) {
      return res.status(400).json({ error: "Invalid file id." });
    }

    let buffer;
    try {
      buffer = await fileUtils.readBuffer(fileId, ".xlsx");
    } catch {
      return res.status(404).json({ error: "This file is no longer available. It may have expired - please regenerate it." });
    }

    const displayName = fileUtils.sanitizeDisplayName(req.query.name || "workbook.xlsx");
    res.setHeader("Content-Type", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
    res.setHeader("Content-Disposition", `attachment; filename="${displayName}"`);
    res.send(buffer);
  } catch (err) {
    next(err);
  }
});

module.exports = router;
