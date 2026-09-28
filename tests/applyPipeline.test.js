"use strict";

const { computeFinalStatus } = require("../src/services/applyPipeline");

const op = (type) => ({ type });
const empty = { verified: [], failed: [], unchecked: [] };

describe("computeFinalStatus is strict", () => {
  test("applied-but-failed-verification is FAILED, not successful", () => {
    const a = op("format_cells"), b = op("set_cell");
    const report = computeFinalStatus({
      results: [{ op: a, ok: true }, { op: b, ok: true }],
      verification: { verified: [{ op: a }], failed: [{ op: b, reason: "expected change was not found after applying" }], unchecked: [] },
      chartVerification: { verified: [], failed: [] },
    });
    expect(report.status).toBe("partial");
    expect(report.verifiedCount).toBe(1);
    expect(report.failedCount).toBe(1);
  });

  test("hard failures, chart failures and unverifiable ops are reported separately", () => {
    const a = op("rename_sheet"), c = op("add_chart"), u = op("insert_rows"), v = op("set_cell");
    const report = computeFinalStatus({
      results: [{ op: a, ok: false, error: "nope" }, { op: u, ok: true }, { op: v, ok: true }],
      verification: { verified: [{ op: v }], failed: [{ op: a, reason: "nope" }], unchecked: [{ op: u }] },
      chartVerification: { verified: [], failed: [{ op: c, reason: "chart could not be embedded in the workbook" }] },
    });
    expect(report).toMatchObject({ status: "partial", verifiedCount: 1, failedCount: 2, unverifiedCount: 1 });
  });

  test("all verified => complete; only unverifiable => complete_with_unverified; nothing => failed", () => {
    const a = op("set_cell"), u = op("insert_rows");
    expect(computeFinalStatus({ results: [{ op: a, ok: true }], verification: { ...empty, verified: [{ op: a }] }, chartVerification: { verified: [], failed: [] } }).status).toBe("complete");
    expect(computeFinalStatus({ results: [{ op: u, ok: true }], verification: { ...empty, unchecked: [{ op: u }] }, chartVerification: { verified: [], failed: [] } }).status).toBe("complete_with_unverified");
    expect(computeFinalStatus({ results: [{ op: a, ok: false, error: "x" }], verification: { ...empty, failed: [{ op: a, reason: "x" }] }, chartVerification: { verified: [], failed: [] } }).status).toBe("failed");
  });
});
