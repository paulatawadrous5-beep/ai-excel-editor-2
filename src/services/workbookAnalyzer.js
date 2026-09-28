"use strict";

const ExcelJS = require("exceljs");

const MAX_SAMPLE_ROWS = 8;
const MAX_SAMPLE_COLS = 15;

function colLetter(n) {
  let s = "";
  while (n > 0) {
    const rem = (n - 1) % 26;
    s = String.fromCharCode(65 + rem) + s;
    n = Math.floor((n - 1) / 26);
  }
  return s;
}

function cellDisplayValue(cell) {
  if (cell.value === null || cell.value === undefined) return "";
  if (typeof cell.value === "object") {
    if (cell.value.richText) return cell.value.richText.map((t) => t.text).join("");
    if (cell.value.formula !== undefined) return `=${cell.value.formula}`;
    if (cell.value.result !== undefined) return cell.value.result;
    if (cell.value instanceof Date) return cell.value.toISOString().slice(0, 10);
  }
  return cell.value;
}

function isRowEffectivelyEmpty(row, colCount) {
  for (let c = 1; c <= colCount; c++) {
    const val = row.getCell(c).value;
    if (val !== null && val !== undefined && String(val).trim() !== "") return false;
  }
  return true;
}

/** Loads a workbook from a buffer. Throws a descriptive error on bad files. */
async function loadWorkbook(buffer) {
  const workbook = new ExcelJS.Workbook();
  try {
    await workbook.xlsx.load(buffer);
  } catch (err) {
    const wrapped = new Error("The uploaded file could not be read as a valid .xlsx workbook");
    wrapped.cause = err;
    wrapped.userFacing = true;
    throw wrapped;
  }
  if (!workbook.worksheets.length) {
    const err = new Error("The workbook has no worksheets");
    err.userFacing = true;
    throw err;
  }
  return workbook;
}

/**
 * Produces a compact JSON-friendly structural summary of a workbook:
 * sheet names, dimensions, header row guess, a few sample rows, merged
 * cells, freeze panes, and detected trailing empty rows. This is what gets
 * sent to the AI (instead of the whole file) and shown to developers.
 */
function analyzeWorkbook(workbook) {
  const sheets = workbook.worksheets.map((ws) => {
    const rowCount = ws.actualRowCount || ws.rowCount || 0;
    const colCount = Math.min(ws.actualColumnCount || ws.columnCount || 0, MAX_SAMPLE_COLS);

    const headerRow = [];
    if (rowCount > 0) {
      const row1 = ws.getRow(1);
      for (let c = 1; c <= colCount; c++) {
        headerRow.push(cellDisplayValue(row1.getCell(c)));
      }
    }

    const sampleRows = [];
    const lastSampleRow = Math.min(rowCount, MAX_SAMPLE_ROWS + 1);
    for (let r = 2; r <= lastSampleRow; r++) {
      const row = ws.getRow(r);
      const values = [];
      for (let c = 1; c <= colCount; c++) {
        values.push(cellDisplayValue(row.getCell(c)));
      }
      sampleRows.push({ rowNumber: r, values });
    }

    const trailingEmptyRows = [];
    for (let r = rowCount; r >= 1 && r > rowCount - 20 && r > 0; r--) {
      const row = ws.getRow(r);
      if (isRowEffectivelyEmpty(row, colCount || ws.columnCount || 1)) {
        trailingEmptyRows.push(r);
      } else {
        break;
      }
    }

    const mergedRanges = [];
    if (ws._merges) {
      Object.keys(ws._merges).forEach((key) => mergedRanges.push(key));
    }

    return {
      name: ws.name,
      rowCount,
      columnCount: ws.actualColumnCount || ws.columnCount || colCount,
      lastColumnLetter: colLetter(colCount || 1),
      headerRow,
      sampleRows,
      trailingEmptyRows: trailingEmptyRows.reverse(),
      mergedRanges,
      frozen: !!(ws.views && ws.views[0] && (ws.views[0].xSplit || ws.views[0].ySplit)),
      hasTable: (ws.tables && Object.keys(ws.tables).length > 0) || false,
    };
  });

  return {
    sheetNames: sheets.map((s) => s.name),
    sheets,
  };
}

module.exports = { loadWorkbook, analyzeWorkbook, colLetter };
