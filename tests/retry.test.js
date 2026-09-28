"use strict";

// Verifies retry safety in applyPipeline without needing real workbooks:
// only idempotent operations may be re-applied; charts/rows/columns never.
jest.mock("jszip", () => ({ loadAsync: async () => ({ files: {} }) }));
jest.mock("../src/services/chartEngine", () => ({ injectCharts: jest.fn(async (buf) => buf) }));
jest.mock("../src/services/workbookAnalyzer", () => ({
  loadWorkbook: async () => ({
    getWorksheet: () => ({ getCell: () => ({ font: {}, value: null }), actualRowCount: 0, rowCount: 0 }),
  }),
}));

const { injectCharts } = require("../src/services/chartEngine");
const { runPipeline } = require("../src/services/applyPipeline");

function fakeEngine() {
  const applied = [];
  return {
    applied,
    pendingCharts: [{ type: "add_chart" }],
    workbook: {},
    applyAll: (ops) => ops.map((op) => ({ op, ok: true })),
    apply: (op) => applied.push(op.type),
    toBuffer: async () => Buffer.from("x"),
  };
}

test("retries only idempotent ops; never charts, rows, columns or totals", async () => {
  const engine = fakeEngine();
  const ops = [
    { type: "format_cells", sheet: "S", range: "A1:A1", bold: true }, // fails verification, retryable
    { type: "add_total_row", sheet: "S" },                            // fails verification, NOT retryable
    { type: "insert_rows", sheet: "S", at: 1 },                       // unchecked, never retried
    { type: "add_chart", sheet: "S", chartType: "pie", dataRange: "A1:B2" },
  ];
  const { report } = await runPipeline(engine, ops);
  expect(engine.applied).toEqual(["format_cells"]);          // exactly one retry, idempotent only
  expect(injectCharts).toHaveBeenCalledTimes(2);              // rebuilt from a fresh buffer, not stacked
  expect(report.status).toBe("partial");
  expect(report.failed.map((f) => f.type)).toEqual(expect.arrayContaining(["format_cells", "add_total_row", "add_chart"]));
  expect(report.verifiedCount).toBe(0);
});
