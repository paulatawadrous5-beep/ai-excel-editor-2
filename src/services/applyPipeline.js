"use strict";

const JSZip = require("jszip");
const { injectCharts } = require("./chartEngine");
const { verifyOperations, verifyCharts } = require("./verifier");
const { loadWorkbook } = require("./workbookAnalyzer");
const logger = require("../utils/logger");

/** Operations that are safe to re-apply once if verification did not find their effect. */
const RETRYABLE_TYPES = new Set([
  "set_cell", "set_range_values", "set_formula", "format_cells", "number_format",
  "set_row_height", "set_column_width", "freeze_panes", "add_hyperlink", "data_validation", "page_setup",
]);

async function countChartParts(buffer) {
  const zip = await JSZip.loadAsync(buffer);
  return Object.keys(zip.files).filter((f) => /^xl\/charts\/chart\d+\.xml$/.test(f)).length;
}

/**
 * Pure function: turns raw apply + verification outcomes into the final,
 * honest status. An operation only counts as "verified" if the re-opened
 * file proved it. Applied-but-failed-verification counts as FAILED.
 */
function computeFinalStatus({ results, verification, chartVerification }) {
  const verified = [...verification.verified, ...chartVerification.verified].map((v) => ({ type: v.op.type }));
  const unverified = verification.unchecked.map((u) => ({
    type: u.op.type,
    note: "applied, but this change type cannot be independently verified",
  }));
  const failed = [
    ...results.filter((r) => !r.ok).map((r) => ({ type: r.op.type, error: r.error })),
    ...verification.failed.filter((f) => results.find((r) => r.op === f.op && r.ok)).map((f) => ({ type: f.op.type, error: f.reason })),
    ...chartVerification.failed.map((f) => ({ type: f.op.type, error: f.reason })),
  ];
  let status = "complete";
  if (failed.length) status = verified.length + unverified.length === 0 ? "failed" : "partial";
  else if (unverified.length) status = "complete_with_unverified";
  return {
    status,
    verifiedCount: verified.length,
    failedCount: failed.length,
    unverifiedCount: unverified.length,
    verified, failed, unverified,
  };
}

/**
 * Applies operations to an ExcelEngine, injects charts, verifies against the
 * re-opened final bytes, retries verification failures once (idempotent ops
 * only), and returns the final buffer plus a strict status report.
 */
async function runPipeline(engine, operations) {
  const results = engine.applyAll(operations);
  const chartOps = operations.filter((o) => o.type === "add_chart");

  async function build() {
    const intermediate = await engine.toBuffer();
    const before = await countChartParts(intermediate);
    const finalBuffer = engine.pendingCharts.length
      ? await injectCharts(intermediate, engine.workbook, engine.pendingCharts)
      : intermediate;
    const after = await countChartParts(finalBuffer);
    const wb = await loadWorkbook(finalBuffer);
    return {
      finalBuffer,
      verification: verifyOperations(wb, results.filter((r) => r.op.type !== "add_chart")),
      chartVerification: verifyCharts(before, after, chartOps),
    };
  }

  let built = await build();
  const retryable = built.verification.failed.filter(
    (f) => RETRYABLE_TYPES.has(f.op.type) && results.find((r) => r.op === f.op && r.ok)
  );
  if (retryable.length) {
    for (const f of retryable) {
      try {
        engine.apply(f.op);
      } catch (err) {
        logger.warn(`Retry failed for ${f.op.type}`, err.message);
      }
    }
    built = await build();
  }

  return { finalBuffer: built.finalBuffer, report: computeFinalStatus({ results, ...built }) };
}

module.exports = { runPipeline, computeFinalStatus };
