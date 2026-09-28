"use strict";

const fs = require("fs");
const fsp = fs.promises;
const path = require("path");
const { v4: uuidv4 } = require("uuid");
const config = require("../config");

/**
 * All uploaded/generated files live inside config.tempDir under a filename
 * we generate ourselves (a uuid). We NEVER use a user-supplied filename or
 * path segment to build a filesystem path, which is what prevents path
 * traversal regardless of what the client sends.
 */

function ensureTempDir() {
  if (!fs.existsSync(config.tempDir)) {
    fs.mkdirSync(config.tempDir, { recursive: true });
  }
  return config.tempDir;
}

/** Returns an absolute path inside the temp dir for a given internal id. */
function tempPathFor(id, suffix) {
  ensureTempDir();
  const safeId = String(id).replace(/[^a-zA-Z0-9-_]/g, "");
  if (!safeId) throw new Error("Invalid internal file id");
  const fileName = `${safeId}${suffix || ".xlsx"}`;
  const fullPath = path.join(config.tempDir, fileName);

  // Defense in depth: verify the resolved path is still inside tempDir.
  const resolved = path.resolve(fullPath);
  const resolvedDir = path.resolve(config.tempDir);
  if (!resolved.startsWith(resolvedDir + path.sep) && resolved !== resolvedDir) {
    throw new Error("Path traversal detected");
  }
  return resolved;
}

function newInternalId() {
  return uuidv4();
}

/** Sanitizes a user-supplied original filename for display purposes only. */
function sanitizeDisplayName(name) {
  if (!name) return "workbook.xlsx";
  const base = path.basename(String(name));
  const cleaned = base.replace(/[^a-zA-Z0-9 ._-]/g, "_").slice(0, 120);
  return cleaned || "workbook.xlsx";
}

async function writeBuffer(id, suffix, buffer) {
  const p = tempPathFor(id, suffix);
  await fsp.writeFile(p, buffer);
  return p;
}

async function readBuffer(id, suffix) {
  const p = tempPathFor(id, suffix);
  return fsp.readFile(p);
}

async function removeIfExists(id, suffix) {
  const p = tempPathFor(id, suffix);
  try {
    await fsp.unlink(p);
  } catch (err) {
    if (err.code !== "ENOENT") throw err;
  }
}

/** Deletes any temp files older than the configured TTL. Call periodically. */
async function sweepExpired() {
  ensureTempDir();
  const now = Date.now();
  let entries;
  try {
    entries = await fsp.readdir(config.tempDir);
  } catch {
    return;
  }
  await Promise.all(
    entries.map(async (name) => {
      const full = path.join(config.tempDir, name);
      try {
        const stat = await fsp.stat(full);
        if (now - stat.mtimeMs > config.fileTtlMs) {
          await fsp.unlink(full);
        }
      } catch {
        /* ignore races */
      }
    })
  );
}

module.exports = {
  ensureTempDir,
  tempPathFor,
  newInternalId,
  sanitizeDisplayName,
  writeBuffer,
  readBuffer,
  removeIfExists,
  sweepExpired,
};
