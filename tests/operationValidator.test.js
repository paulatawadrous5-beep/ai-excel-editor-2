"use strict";

const { validateOperations } = require("../src/operations/operationValidator");

describe("operationValidator", () => {
  test("accepts a well-formed set of operations", () => {
    const { valid, issues } = validateOperations([
      { type: "format_cells", sheet: "Sheet1", range: "A1:E1", bold: true, fillColor: "1F4E78" },
      { type: "set_row_height", sheet: "Sheet1", rows: "1:20", height: 0.75 },
      { type: "add_total_column", sheet: "Sheet1", header: "Total", formulaType: "sum" },
      { type: "add_chart", sheet: "Sheet1", chartType: "column", dataRange: "A1:B13", title: "Sales" },
    ]);
    expect(issues).toHaveLength(0);
    expect(valid).toHaveLength(4);
  });

  test("rejects unknown operation types", () => {
    const { valid, issues } = validateOperations([{ type: "delete_everything", sheet: "Sheet1" }]);
    expect(valid).toHaveLength(0);
    expect(issues[0].message).toMatch(/unknown operation type/);
  });

  test("rejects missing required fields", () => {
    const { valid, issues } = validateOperations([{ type: "format_cells", sheet: "Sheet1" }]);
    expect(valid).toHaveLength(0);
    expect(issues[0].message).toMatch(/missing required field/);
  });

  test("rejects malformed cell/range references", () => {
    const { valid, issues } = validateOperations([{ type: "set_cell", sheet: "Sheet1", cell: "1A", value: 5 }]);
    expect(valid).toHaveLength(0);
    expect(issues[0].message).toMatch(/valid cell reference/);
  });

  test("rejects invalid hex colors", () => {
    const { valid } = validateOperations([
      { type: "format_cells", sheet: "Sheet1", range: "A1:A1", fillColor: "not-a-color" },
    ]);
    expect(valid).toHaveLength(0);
  });

  test("rejects an empty operations array", () => {
    const { valid, issues } = validateOperations([]);
    expect(valid).toHaveLength(0);
    expect(issues[0].message).toMatch(/no operations/);
  });

  test("strips unknown fields instead of failing the whole operation", () => {
    const { valid, issues } = validateOperations([
      { type: "set_row_height", sheet: "Sheet1", rows: "1", height: 20, extraField: "nope" },
    ]);
    expect(valid).toHaveLength(1);
    expect(valid[0].extraField).toBeUndefined();
    expect(issues[0].message).toMatch(/unsupported field/);
  });

  test("normalizes a leading '=' on formulas instead of rejecting", () => {
    const { valid } = validateOperations([{ type: "set_formula", sheet: "Sheet1", cell: "C1", formula: "=SUM(A1:A10)" }]);
    expect(valid).toHaveLength(1);
    expect(valid[0].formula).toBe("SUM(A1:A10)");
  });
});
