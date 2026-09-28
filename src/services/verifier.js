"use strict";

const { parseRange, parseColumnsSpec, parseRowsSpec, parseCell } = require("../utils/cellRef");
const { normalizeColor } = require("../operations/operationValidator");
const NUMBER_FORMAT_CODES = {
  currency: '"$"#,##0.00', percentage: "0.00%", date: "yyyy-mm-dd", decimal: "0.00", integer: "0", text: "@",
};

function cellValueOf(cell) {
  if (cell.value && typeof cell.value === "object") {
    if (cell.value.formula !== undefined) return { formula: cell.value.formula };
    if (cell.value.result !== undefined) return cell.value.result;
    if (cell.value.richText) return cell.value.richText.map((t) => t.text).join("");
    if (cell.value.hyperlink !== undefined) return { hyperlink: cell.value.hyperlink, text: cell.value.text };
  }
  return cell.value;
}

/**
 * Each checker returns true (verified), false (failed), or null (not
 * independently checkable - treated as "trust the engine" and not counted
 * as a failure, e.g. purely structural ops with no simple invariant).
 */
const CHECKERS = {
  create_sheet: (wb, op) => !!wb.getWorksheet(op.sheet),
  rename_sheet: (wb, op) => !!wb.getWorksheet(op.newName),
  delete_sheet: (wb, op) => !wb.getWorksheet(op.sheet),
  reorder_sheets: () => null,

  set_cell: (wb, op) => {
    const ws = wb.getWorksheet(op.sheet);
    if (!ws) return false;
    const cell = ws.getCell(op.cell);
    if (op.formula) {
      return typeof cell.value === "object" && cell.value !== null && String(cell.value.formula || "").replace(/\s/g, "") === String(op.formula).replace(/^=/, "").replace(/\s/g, "");
    }
    const val = cellValueOf(cell);
    if (op.value === null) return val === null || val === undefined;
    return String(val) === String(op.value);
  },
  set_range_values: (wb, op) => {
    const ws = wb.getWorksheet(op.sheet);
    if (!ws) return false;
    const start = parseCell(op.startCell);
    const first = op.values[0] && op.values[0][0];
    const cellVal = cellValueOf(ws.getCell(start.row, start.col));
    return first === undefined || String(cellVal) === String(first);
  },
  clear_cells: (wb, op) => {
    const ws = wb.getWorksheet(op.sheet);
    if (!ws) return false;
    const r = parseRange(op.range);
    const v = ws.getCell(r.startRow, r.startCol).value;
    return v === null || v === undefined;
  },
  copy_cells: () => null,
  move_cells: () => null,
  find_replace: () => null,

  insert_rows: () => null,
  delete_rows: () => null,
  delete_empty_rows: (wb, op) => {
    const ws = wb.getWorksheet(op.sheet);
    if (!ws) return false;
    const colCount = ws.actualColumnCount || ws.columnCount || 1;
    const lastRow = ws.actualRowCount || ws.rowCount || 0;
    if (lastRow === 0) return true;
    const row = ws.getRow(lastRow);
    for (let c = 1; c <= colCount; c++) {
      const v = row.getCell(c).value;
      if (v !== null && v !== undefined && String(v).trim() !== "") return true;
    }
    return false;
  },
  set_row_height: (wb, op) => {
    const ws = wb.getWorksheet(op.sheet);
    if (!ws) return false;
    const { startRow } = parseRowsSpec(op.rows);
    const h = ws.getRow(startRow).height;
    return Math.abs((h || 0) - op.height) < 0.5;
  },

  insert_columns: () => null,
  delete_columns: () => null,
  set_column_width: (wb, op) => {
    const ws = wb.getWorksheet(op.sheet);
    if (!ws) return false;
    const { startCol } = parseColumnsSpec(op.columns);
    const w = ws.getColumn(startCol).width;
    return Math.abs((w || 0) - op.width) < 0.5;
  },
  auto_fit_columns: () => null,

  format_cells: (wb, op) => {
    const ws = wb.getWorksheet(op.sheet);
    if (!ws) return false;
    const r = parseRange(op.range);
    const cell = ws.getCell(r.startRow, r.startCol);
    const font = cell.font || {};
    if (op.bold !== undefined && !!font.bold !== !!op.bold) return false;
    if (op.italic !== undefined && !!font.italic !== !!op.italic) return false;
    if (op.fontSize !== undefined && font.size !== op.fontSize) return false;
    if (op.fontColor !== undefined) {
      const expected = normalizeColor(op.fontColor);
      if (!font.color || !String(font.color.argb || "").endsWith(expected)) return false;
    }
    if (op.fillColor !== undefined) {
      const expected = normalizeColor(op.fillColor);
      const fg = cell.fill && cell.fill.fgColor;
      if (!fg || !String(fg.argb || "").endsWith(expected)) return false;
    }
    return true;
  },
  number_format: (wb, op) => {
    const ws = wb.getWorksheet(op.sheet);
    if (!ws) return false;
    const r = parseRange(op.range);
    const expected = op.format === "custom" ? op.customFormat : NUMBER_FORMAT_CODES[op.format];
    return !!expected && ws.getCell(r.startRow, r.startCol).numFmt === expected;
  },

  merge_cells: (wb, op) => {
    const ws = wb.getWorksheet(op.sheet);
    if (!ws) return false;
    const r = parseRange(op.range);
    return !!ws.getCell(r.startRow, r.startCol).isMerged;
  },
  unmerge_cells: (wb, op) => {
    const ws = wb.getWorksheet(op.sheet);
    if (!ws) return false;
    const r = parseRange(op.range);
    return !ws.getCell(r.startRow, r.startCol).isMerged;
  },
  freeze_panes: (wb, op) => {
    const ws = wb.getWorksheet(op.sheet);
    if (!ws) return false;
    return !!(ws.views && ws.views[0] && ws.views[0].state === "frozen");
  },
  create_table: (wb, op) => {
    const ws = wb.getWorksheet(op.sheet);
    if (!ws) return false;
    return !!(ws.tables && Object.keys(ws.tables).length > 0);
  },
  sort_range: () => null,

  set_formula: (wb, op) => {
    const ws = wb.getWorksheet(op.sheet);
    if (!ws) return false;
    const cell = ws.getCell(op.cell);
    return typeof cell.value === "object" && cell.value !== null && !!cell.value.formula;
  },
  add_total_column: (wb, op) => {
    const ws = wb.getWorksheet(op.sheet);
    if (!ws) return false;
    const headerCol = ws.actualColumnCount || ws.columnCount;
    for (let c = 1; c <= headerCol; c++) {
      if (String(ws.getCell(1, c).value) === String(op.header)) return true;
    }
    return false;
  },
  add_total_row: (wb, op) => {
    const ws = wb.getWorksheet(op.sheet);
    if (!ws) return false;
    const lastRow = ws.actualRowCount || ws.rowCount;
    return String(ws.getCell(lastRow, 1).value || "").length > 0;
  },

  add_chart: () => null, // verified separately against the raw zip (see verifyCharts)

  add_hyperlink: (wb, op) => {
    const ws = wb.getWorksheet(op.sheet);
    if (!ws) return false;
    const cell = ws.getCell(op.cell);
    return typeof cell.value === "object" && cell.value !== null && cell.value.hyperlink === op.url;
  },
  data_validation: (wb, op) => {
    const ws = wb.getWorksheet(op.sheet);
    if (!ws) return false;
    const r = parseRange(op.range);
    return !!ws.getCell(r.startRow, r.startCol).dataValidation;
  },
  conditional_format: () => null,
  page_setup: (wb, op) => {
    const ws = wb.getWorksheet(op.sheet);
    if (!ws) return false;
    return !op.orientation || ws.pageSetup.orientation === op.orientation;
  },
};

/**
 * Re-opens the generated workbook and checks that each successfully-applied
 * operation actually produced the expected state.
 * @returns {{ verified: object[], failed: object[], unchecked: object[] }}
 */
function verifyOperations(workbook, appliedResults) {
  const verified = [];
  const failed = [];
  const unchecked = [];

  for (const result of appliedResults) {
    if (!result.ok) {
      failed.push({ op: result.op, reason: result.error });
      continue;
    }
    const checker = CHECKERS[result.op.type];
    if (!checker) {
      unchecked.push({ op: result.op });
      continue;
    }
    let outcome;
    try {
      outcome = checker(workbook, result.op);
    } catch (err) {
      outcome = false;
    }
    if (outcome === null) {
      unchecked.push({ op: result.op });
    } else if (outcome) {
      verified.push({ op: result.op });
    } else {
      failed.push({ op: result.op, reason: "expected change was not found after applying" });
    }
  }

  return { verified, failed, unchecked };
}

/** Confirms chart parts were actually written into the final zip. */
function verifyCharts(chartCountBefore, chartCountAfter, chartOps) {
  const expected = chartOps.length;
  const added = chartCountAfter - chartCountBefore;
  if (added >= expected) {
    return { verified: chartOps.map((op) => ({ op })), failed: [] };
  }
  return {
    verified: chartOps.slice(0, Math.max(added, 0)).map((op) => ({ op })),
    failed: chartOps.slice(Math.max(added, 0)).map((op) => ({ op, reason: "chart could not be embedded in the workbook" })),
  };
}

module.exports = { verifyOperations, verifyCharts };
