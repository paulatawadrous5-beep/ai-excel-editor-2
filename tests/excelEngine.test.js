"use strict";

const { ExcelEngine } = require("../src/services/excelEngine");
const ExcelJS = require("exceljs");

async function reload(engine) {
  const buffer = await engine.toBuffer();
  const wb = new ExcelJS.Workbook();
  await wb.xlsx.load(buffer);
  return wb;
}

describe("ExcelEngine - workbook operations", () => {
  test("creates a blank workbook and adds sheets", () => {
    const engine = ExcelEngine.blank();
    engine.workbook.addWorksheet("Sheet1");
    engine.apply({ type: "create_sheet", sheet: "Data" });
    expect(engine.workbook.getWorksheet("Sheet1")).toBeTruthy();
    expect(engine.workbook.getWorksheet("Data")).toBeTruthy();
  });

  test("renames and deletes sheets", () => {
    const engine = ExcelEngine.blank();
    engine.workbook.addWorksheet("Sheet1");
    engine.apply({ type: "create_sheet", sheet: "Temp" });
    engine.apply({ type: "rename_sheet", sheet: "Temp", newName: "Renamed" });
    expect(engine.workbook.getWorksheet("Renamed")).toBeTruthy();
    expect(engine.workbook.getWorksheet("Temp")).toBeFalsy();
    engine.apply({ type: "delete_sheet", sheet: "Renamed" });
    expect(engine.workbook.getWorksheet("Renamed")).toBeFalsy();
  });

  test("refuses to delete the last remaining worksheet", () => {
    const engine = ExcelEngine.blank();
    engine.workbook.addWorksheet("OnlySheet");
    expect(() => engine.apply({ type: "delete_sheet", sheet: "OnlySheet" })).toThrow();
  });
});

describe("ExcelEngine - cells", () => {
  let engine;
  beforeEach(() => {
    engine = ExcelEngine.blank();
    engine.workbook.addWorksheet("Sheet1");
  });

  test("set_cell writes a plain value", async () => {
    engine.apply({ type: "set_cell", sheet: "Sheet1", cell: "A1", value: "Hello" });
    const wb = await reload(engine);
    expect(wb.getWorksheet("Sheet1").getCell("A1").value).toBe("Hello");
  });

  test("set_range_values writes a 2D block", async () => {
    engine.apply({
      type: "set_range_values",
      sheet: "Sheet1",
      startCell: "A1",
      values: [
        ["Month", "Revenue"],
        ["Jan", 100],
        ["Feb", 200],
      ],
    });
    const wb = await reload(engine);
    const ws = wb.getWorksheet("Sheet1");
    expect(ws.getCell("A1").value).toBe("Month");
    expect(ws.getCell("B3").value).toBe(200);
  });

  test("clear_cells empties a range", async () => {
    engine.apply({ type: "set_cell", sheet: "Sheet1", cell: "A1", value: "x" });
    engine.apply({ type: "clear_cells", sheet: "Sheet1", range: "A1:A1" });
    const wb = await reload(engine);
    expect(wb.getWorksheet("Sheet1").getCell("A1").value).toBeNull();
  });

  test("find_replace updates matching text", async () => {
    engine.apply({ type: "set_cell", sheet: "Sheet1", cell: "A1", value: "Hello World" });
    engine.apply({ type: "find_replace", sheet: "Sheet1", find: "World", replace: "There" });
    const wb = await reload(engine);
    expect(wb.getWorksheet("Sheet1").getCell("A1").value).toBe("Hello There");
  });
});

describe("ExcelEngine - rows and columns", () => {
  let engine;
  beforeEach(() => {
    engine = ExcelEngine.blank();
    engine.workbook.addWorksheet("Sheet1");
    engine.apply({
      type: "set_range_values",
      sheet: "Sheet1",
      startCell: "A1",
      values: [
        ["Month", "Revenue"],
        ["Jan", 100],
        ["", ""],
      ],
    });
  });

  test("delete_empty_rows removes trailing blank rows", async () => {
    engine.apply({ type: "delete_empty_rows", sheet: "Sheet1" });
    const wb = await reload(engine);
    const ws = wb.getWorksheet("Sheet1");
    expect(ws.actualRowCount).toBe(2);
  });

  test("set_row_height applies the requested height", async () => {
    engine.apply({ type: "set_row_height", sheet: "Sheet1", rows: "1", height: 30 });
    const wb = await reload(engine);
    expect(wb.getWorksheet("Sheet1").getRow(1).height).toBeCloseTo(30);
  });

  test("set_column_width applies the requested width", async () => {
    engine.apply({ type: "set_column_width", sheet: "Sheet1", columns: "A", width: 25 });
    const wb = await reload(engine);
    expect(wb.getWorksheet("Sheet1").getColumn(1).width).toBeCloseTo(25);
  });

  test("insert_rows and delete_rows change row count", async () => {
    engine.apply({ type: "insert_rows", sheet: "Sheet1", at: 2, count: 2 });
    let wb = await reload(engine);
    expect(wb.getWorksheet("Sheet1").getCell("A4").value).toBe("Jan");

    engine.apply({ type: "delete_rows", sheet: "Sheet1", rows: "2:3" });
    wb = await reload(engine);
    expect(wb.getWorksheet("Sheet1").getCell("A2").value).toBe("Jan");
  });
});

describe("ExcelEngine - formatting", () => {
  let engine;
  beforeEach(() => {
    engine = ExcelEngine.blank();
    engine.workbook.addWorksheet("Sheet1");
    engine.apply({ type: "set_cell", sheet: "Sheet1", cell: "A1", value: "Header" });
  });

  test("bold, fill and font color are applied", async () => {
    engine.apply({
      type: "format_cells",
      sheet: "Sheet1",
      range: "A1:A1",
      bold: true,
      fillColor: "1F4E78",
      fontColor: "FFFFFF",
    });
    const wb = await reload(engine);
    const cell = wb.getWorksheet("Sheet1").getCell("A1");
    expect(cell.font.bold).toBe(true);
    expect(cell.fill.fgColor.argb).toMatch(/1F4E78$/i);
    expect(cell.font.color.argb).toMatch(/FFFFFF$/i);
  });

  test("number_format applies currency formatting", async () => {
    engine.apply({ type: "set_cell", sheet: "Sheet1", cell: "B1", value: 1234.5 });
    engine.apply({ type: "number_format", sheet: "Sheet1", range: "B1:B1", format: "currency" });
    const wb = await reload(engine);
    expect(wb.getWorksheet("Sheet1").getCell("B1").numFmt).toContain("$");
  });

  test("merge_cells and unmerge_cells", async () => {
    engine.apply({ type: "merge_cells", sheet: "Sheet1", range: "A1:B1" });
    let wb = await reload(engine);
    expect(wb.getWorksheet("Sheet1").getCell("A1").isMerged).toBe(true);

    engine.apply({ type: "unmerge_cells", sheet: "Sheet1", range: "A1:B1" });
    wb = await reload(engine);
    expect(wb.getWorksheet("Sheet1").getCell("A1").isMerged).toBe(false);
  });

  test("freeze_panes freezes at the requested cell", async () => {
    engine.apply({ type: "freeze_panes", sheet: "Sheet1", cell: "B2" });
    const wb = await reload(engine);
    const views = wb.getWorksheet("Sheet1").views;
    expect(views[0].state).toBe("frozen");
  });
});

describe("ExcelEngine - formulas and totals", () => {
  let engine;
  beforeEach(() => {
    engine = ExcelEngine.blank();
    engine.workbook.addWorksheet("Sheet1");
    engine.apply({
      type: "set_range_values",
      sheet: "Sheet1",
      startCell: "A1",
      values: [
        ["Month", "Revenue"],
        ["Jan", 100],
        ["Feb", 200],
      ],
    });
  });

  test("set_formula writes SUM/AVERAGE/IF style formulas", async () => {
    engine.apply({ type: "set_formula", sheet: "Sheet1", cell: "B4", formula: "SUM(B2:B3)" });
    engine.apply({ type: "set_formula", sheet: "Sheet1", cell: "B5", formula: "AVERAGE(B2:B3)" });
    engine.apply({ type: "set_formula", sheet: "Sheet1", cell: "B6", formula: 'IF(B4>100,"High","Low")' });
    const wb = await reload(engine);
    const ws = wb.getWorksheet("Sheet1");
    expect(ws.getCell("B4").value.formula).toBe("SUM(B2:B3)");
    expect(ws.getCell("B5").value.formula).toBe("AVERAGE(B2:B3)");
    expect(ws.getCell("B6").value.formula).toContain("IF(");
  });

  test("add_total_column inserts a header and per-row SUM formulas", async () => {
    engine.apply({ type: "add_total_column", sheet: "Sheet1", header: "Total", formulaType: "sum" });
    const wb = await reload(engine);
    const ws = wb.getWorksheet("Sheet1");
    expect(ws.getCell(1, 3).value).toBe("Total");
    expect(typeof ws.getCell(2, 3).value).toBe("object");
    expect(ws.getCell(2, 3).value.formula).toMatch(/SUM/);
  });
});

describe("ExcelEngine - charts are queued, not executed inline", () => {
  test("add_chart pushes to pendingCharts without touching cells", () => {
    const engine = ExcelEngine.blank();
    engine.workbook.addWorksheet("Sheet1");
    engine.apply({ type: "add_chart", sheet: "Sheet1", chartType: "column", dataRange: "A1:B3", title: "Test" });
    expect(engine.pendingCharts).toHaveLength(1);
    expect(engine.pendingCharts[0].chartType).toBe("column");
  });
});

describe("ExcelEngine.applyAll", () => {
  test("continues past a failing operation and reports it", () => {
    const engine = ExcelEngine.blank();
    engine.workbook.addWorksheet("Sheet1");
    const results = engine.applyAll([
      { type: "set_cell", sheet: "Sheet1", cell: "A1", value: "ok" },
      { type: "rename_sheet", sheet: "DoesNotExist", newName: "X" },
      { type: "set_cell", sheet: "Sheet1", cell: "A2", value: "also ok" },
    ]);
    expect(results[0].ok).toBe(true);
    expect(results[1].ok).toBe(false);
    expect(results[2].ok).toBe(true);
  });
});
