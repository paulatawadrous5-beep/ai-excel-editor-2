"use strict";

const ExcelJS = require("exceljs");
const { ExcelEngine } = require("../src/services/excelEngine");
const { verifyOperations, verifyCharts } = require("../src/services/verifier");

async function reload(engine) {
  const buffer = await engine.toBuffer();
  const wb = new ExcelJS.Workbook();
  await wb.xlsx.load(buffer);
  return wb;
}

describe("verifyOperations", () => {
  test("confirms a bold + fill format_cells operation actually applied", async () => {
    const engine = ExcelEngine.blank();
    engine.workbook.addWorksheet("Sheet1");
    const op = { type: "format_cells", sheet: "Sheet1", range: "A1:A1", bold: true, fillColor: "1F4E78" };
    const results = engine.applyAll([op]);
    const wb = await reload(engine);

    const { verified, failed } = verifyOperations(wb, results);
    expect(failed).toHaveLength(0);
    expect(verified).toHaveLength(1);
  });

  test("flags an operation whose effect cannot be found", async () => {
    const engine = ExcelEngine.blank();
    engine.workbook.addWorksheet("Sheet1");
    // Fabricate a "successful" apply result that doesn't match reality, to
    // exercise the failure path (simulates something silently not sticking).
    const fakeResults = [{ op: { type: "set_cell", sheet: "Sheet1", cell: "A1", value: "expected" }, ok: true }];
    const wb = await reload(engine); // A1 was never actually set
    const { failed } = verifyOperations(wb, fakeResults);
    expect(failed).toHaveLength(1);
  });

  test("reports hard apply failures as failed, not verified", async () => {
    const results = [{ op: { type: "rename_sheet", sheet: "Ghost", newName: "X" }, ok: false, error: "not found" }];
    const engine = ExcelEngine.blank();
    engine.workbook.addWorksheet("Sheet1");
    const wb = await reload(engine);
    const { failed, verified } = verifyOperations(wb, results);
    expect(failed).toHaveLength(1);
    expect(verified).toHaveLength(0);
  });

  test("marks structural ops with no simple invariant as unchecked, not failed", async () => {
    const engine = ExcelEngine.blank();
    engine.workbook.addWorksheet("Sheet1");
    const results = engine.applyAll([{ type: "insert_rows", sheet: "Sheet1", at: 1, count: 1 }]);
    const wb = await reload(engine);
    const { unchecked, failed } = verifyOperations(wb, results);
    expect(failed).toHaveLength(0);
    expect(unchecked).toHaveLength(1);
  });
});

describe("verifyCharts", () => {
  test("verifies when chart count increased as expected", () => {
    const chartOps = [{ type: "add_chart", sheet: "Sheet1", chartType: "column", dataRange: "A1:B4" }];
    const { verified, failed } = verifyCharts(0, 1, chartOps);
    expect(verified).toHaveLength(1);
    expect(failed).toHaveLength(0);
  });

  test("flags a chart that failed to embed", () => {
    const chartOps = [{ type: "add_chart", sheet: "Sheet1", chartType: "column", dataRange: "A1:B4" }];
    const { verified, failed } = verifyCharts(0, 0, chartOps);
    expect(verified).toHaveLength(0);
    expect(failed).toHaveLength(1);
  });
});
