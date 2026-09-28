"use strict";

const multer = require("multer");
const path = require("path");
const config = require("../config");

const ZIP_MAGIC = Buffer.from([0x50, 0x4b, 0x03, 0x04]); // "PK\x03\x04" - .xlsx is a zip archive

function fileFilter(req, file, cb) {
  const ext = path.extname(file.originalname || "").toLowerCase();
  if (!config.upload.allowedExtensions.includes(ext)) {
    const err = new multer.MulterError("LIMIT_UNEXPECTED_FILE");
    err.userMessage = "Only .xlsx files are supported.";
    return cb(err);
  }
  cb(null, true);
}

const upload = multer({
  storage: multer.memoryStorage(), // never written to disk with the client's filename
  limits: { fileSize: config.upload.maxSizeBytes, files: 1 },
  fileFilter,
});

/** Confirms the uploaded bytes are actually a zip (xlsx) container, not just named *.xlsx. */
function verifyMagicBytes(req, res, next) {
  if (!req.file) {
    return res.status(400).json({ error: "No file was uploaded." });
  }
  const buf = req.file.buffer;
  if (!buf || buf.length < 4 || !buf.subarray(0, 4).equals(ZIP_MAGIC)) {
    return res.status(400).json({ error: "We couldn't process this Excel file. Please make sure it is a valid .xlsx file." });
  }
  next();
}

module.exports = { upload, verifyMagicBytes };
