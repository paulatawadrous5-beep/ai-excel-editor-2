"use strict";

const ExcelJS = require("exceljs");
const { parseCell, parseRange, parseColumnsSpec, parseRowsSpec, numberToCol } = require("../utils/cellRef");
const { normalizeColor } = require("../operations/operationValidator");

const NUMBER_FORMAT_CODES = {
  currency: '"$"#,##0.00',
  percentage: "0.00%",
  date: "yyyy-mm-dd",
  decimal: "0.00",
  integer: "0",
  text: "@",
};

const BORDER_STYLE_MAP = {
  thin: "thin",
  medium: "medium",
  thick: "thick",
  dashed: "dashed",
  dotted: "dotted",
  double: "double",
};

class OperationError extends Error {
  constructor(message, opType) {
    super(message);
    this.name = "OperationError";
    this.opType = opType;
  }
}

class ExcelEngine {
  constructor(workbook) {
    this.workbook = workbook || new ExcelJS.Workbook();
    this.workbook.creator = this.workbook.creator || "AI Excel Editor";
    this.workbook.created = this.workbook.created || new Date();
    this.workbook.modified = new Date();
    /** Chart operations are collected here and injected as native charts
     *  post-write by chartEngine.js (ExcelJS itself cannot write charts). */
    this.pendingCharts = [];
  }

  static async fromBuffer(buffer) {
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.load(buffer);
    return new ExcelEngine(workbook);
  }

  static blank() {
    return new ExcelEngine(new ExcelJS.Workbook());
  }

  getSheet(name, { createIfMissing = false } = {}) {
    let ws = this.workbook.getWorksheet(name);
    if (!ws && createIfMissing) {
      ws = this.workbook.addWorksheet(name);
    }
    if (!ws) throw new OperationError(`Worksheet "${name}" does not exist`, "sheet_lookup");
    return ws;
  }

  /** Applies one structured operation. Returns a short human-readable result note. */
  apply(op) {
    const handler = this[`op_${op.type}`];
    if (!handler) throw new OperationError(`No handler implemented for "${op.type}"`, op.type);
    return handler.call(this, op);
  }

  applyAll(operations) {
    const results = [];
    for (const op of operations) {
      try {
        const note = this.apply(op);
        results.push({ op, ok: true, note });
      } catch (err) {
        results.push({ op, ok: false, error: err.message });
      }
    }
    return results;
  }

  // ---------------------------------------------------------------- Workbook

  op_create_sheet(op) {
    if (this.workbook.getWorksheet(op.sheet)) {
      throw new OperationError(`Worksheet "${op.sheet}" already exists`, op.type);
    }
    this.workbook.addWorksheet(op.sheet);
    return `Created sheet "${op.sheet}"`;
  }

  op_rename_sheet(op) {
    const ws = this.getSheet(op.sheet);
    ws.name = op.newName;
    return `Renamed "${op.sheet}" to "${op.newName}"`;
  }

  op_delete_sheet(op) {
    const ws = this.getSheet(op.sheet);
    if (this.workbook.worksheets.length <= 1) {
      throw new OperationError("Cannot delete the only remaining worksheet", op.type);
    }
    this.workbook.removeWorksheet(ws.id);
    return `Deleted sheet "${op.sheet}"`;
  }

  op_reorder_sheets(op) {
    // ExcelJS does not expose a public reorder API. We set orderNo, which
    // its writer uses as a secondary sort key for sheet tab order. This is
    // best-effort: if it has no effect in a given ExcelJS version, sheets
    // simply keep their original order (no data is lost or corrupted).
    op.order.forEach((name, idx) => {
      const ws = this.workbook.getWorksheet(name);
      if (ws) ws.orderNo = idx;
    });
    return `Requested sheet order: ${op.order.join(", ")}`;
  }

  // ------------------------------------------------------------------ Cells

  op_set_cell(op) {
    const ws = this.getSheet(op.sheet);
    const cell = ws.getCell(op.cell);
    if (op.formula) {
      cell.value = { formula: op.formula.replace(/^=/, "") };
    } else {
      cell.value = op.value;
    }
    return `Set ${op.sheet}!${op.cell}`;
  }

  op_set_range_values(op) {
    const ws = this.getSheet(op.sheet);
    const start = parseCell(op.startCell);
    op.values.forEach((rowValues, rIdx) => {
      rowValues.forEach((val, cIdx) => {
        ws.getCell(start.row + rIdx, start.col + cIdx).value = val;
      });
    });
    return `Wrote ${op.values.length} row(s) starting at ${op.sheet}!${op.startCell}`;
  }

  op_clear_cells(op) {
    const ws = this.getSheet(op.sheet);
    const r = parseRange(op.range);
    for (let row = r.startRow; row <= r.endRow; row++) {
      for (let col = r.startCol; col <= r.endCol; col++) {
        ws.getCell(row, col).value = null;
      }
    }
    return `Cleared ${op.sheet}!${op.range}`;
  }

  op_copy_cells(op) {
    const src = this.getSheet(op.sheet);
    const dst = op.targetSheet ? this.getSheet(op.targetSheet) : src;
    const r = parseRange(op.from);
    const dest = parseCell(op.to.includes(":") ? op.to.split(":")[0] : op.to);
    for (let row = r.startRow; row <= r.endRow; row++) {
      for (let col = r.startCol; col <= r.endCol; col++) {
        const srcCell = src.getCell(row, col);
        const dstCell = dst.getCell(dest.row + (row - r.startRow), dest.col + (col - r.startCol));
        dstCell.value = srcCell.value;
        dstCell.style = { ...srcCell.style };
      }
    }
    return `Copied ${op.sheet}!${op.from} to ${(op.targetSheet || op.sheet)}!${op.to}`;
  }

  op_move_cells(op) {
    const ws = this.getSheet(op.sheet);
    const r = parseRange(op.from);
    const dest = parseCell(op.to.includes(":") ? op.to.split(":")[0] : op.to);
    const snapshot = [];
    for (let row = r.startRow; row <= r.endRow; row++) {
      const rowVals = [];
      for (let col = r.startCol; col <= r.endCol; col++) {
        const c = ws.getCell(row, col);
        rowVals.push({ value: c.value, style: { ...c.style } });
        c.value = null;
      }
      snapshot.push(rowVals);
    }
    snapshot.forEach((rowVals, rIdx) => {
      rowVals.forEach((cellData, cIdx) => {
        const dstCell = ws.getCell(dest.row + rIdx, dest.col + cIdx);
        dstCell.value = cellData.value;
        dstCell.style = cellData.style;
      });
    });
    return `Moved ${op.sheet}!${op.from} to ${op.to}`;
  }

  op_find_replace(op) {
    const ws = this.getSheet(op.sheet);
    const find = op.matchCase ? op.find : String(op.find).toLowerCase();
    let count = 0;
    ws.eachRow({ includeEmpty: false }, (row) => {
      row.eachCell({ includeEmpty: false }, (cell) => {
        if (typeof cell.value === "string") {
          const haystack = op.matchCase ? cell.value : cell.value.toLowerCase();
          if (haystack.includes(find)) {
            const re = new RegExp(op.find.replace(/[.*+?^${}()|[\]\\]/g, "\\$&"), op.matchCase ? "g" : "gi");
            cell.value = cell.value.replace(re, op.replace);
            count++;
          }
        }
      });
    });
    return `Replaced text in ${count} cell(s)`;
  }

  // ------------------------------------------------------------------- Rows

  op_insert_rows(op) {
    const ws = this.getSheet(op.sheet);
    const count = op.count || 1;
    const blanks = Array.from({ length: count }, () => []);
    ws.spliceRows(op.at, 0, ...blanks);
    return `Inserted ${count} row(s) at row ${op.at} in ${op.sheet}`;
  }

  op_delete_rows(op) {
    const ws = this.getSheet(op.sheet);
    const { startRow, endRow } = parseRowsSpec(op.rows);
    ws.spliceRows(startRow, endRow - startRow + 1);
    return `Deleted row(s) ${op.rows} in ${op.sheet}`;
  }

  op_delete_empty_rows(op) {
    const ws = this.getSheet(op.sheet);
    const colCount = ws.actualColumnCount || ws.columnCount || 1;
    const range = op.withinRange ? parseRange(op.withinRange) : null;
    const rowsToDelete = [];
    const lastRow = range ? range.endRow : ws.actualRowCount || ws.rowCount;
    const firstRow = range ? range.startRow : 1;
    for (let r = firstRow; r <= lastRow; r++) {
      const row = ws.getRow(r);
      let empty = true;
      for (let c = 1; c <= colCount; c++) {
        const v = row.getCell(c).value;
        if (v !== null && v !== undefined && String(v).trim() !== "") {
          empty = false;
          break;
        }
      }
      if (empty) rowsToDelete.push(r);
    }
    // Delete from the bottom up so row numbers of not-yet-deleted rows stay valid.
    rowsToDelete.sort((a, b) => b - a).forEach((r) => ws.spliceRows(r, 1));
    return rowsToDelete.length
      ? `Deleted ${rowsToDelete.length} empty row(s): ${rowsToDelete.sort((a, b) => a - b).join(", ")}`
      : "No empty rows found";
  }

  op_set_row_height(op) {
    const ws = this.getSheet(op.sheet);
    const { startRow, endRow } = parseRowsSpec(op.rows);
    for (let r = startRow; r <= endRow; r++) {
      ws.getRow(r).height = op.height;
    }
    return `Set height ${op.height} for row(s) ${op.rows}`;
  }

  // --------------------------------------------------------------- Columns

  op_insert_columns(op) {
    const ws = this.getSheet(op.sheet);
    const atCol = typeof op.at === "number" ? op.at : parseColumnsSpec(op.at).startCol;
    const count = op.count || 1;
    ws.spliceColumns(atCol, 0, ...Array.from({ length: count }, () => []));
    return `Inserted ${count} column(s) at column ${numberToCol(atCol)}`;
  }

  op_delete_columns(op) {
    const ws = this.getSheet(op.sheet);
    const { startCol, endCol } = parseColumnsSpec(op.columns);
    ws.spliceColumns(startCol, endCol - startCol + 1);
    return `Deleted column(s) ${op.columns}`;
  }

  op_set_column_width(op) {
    const ws = this.getSheet(op.sheet);
    const { startCol, endCol } = parseColumnsSpec(op.columns);
    for (let c = startCol; c <= endCol; c++) {
      ws.getColumn(c).width = op.width;
    }
    return `Set width ${op.width} for column(s) ${op.columns}`;
  }

  op_auto_fit_columns(op) {
    const ws = this.getSheet(op.sheet);
    const { startCol, endCol } = parseColumnsSpec(op.columns);
    for (let c = startCol; c <= endCol; c++) {
      const column = ws.getColumn(c);
      let maxLen = 8;
      column.eachCell({ includeEmpty: false }, (cell) => {
        const len = String(cell.value !== null && cell.value !== undefined ? cell.value : "").length;
        if (len > maxLen) maxLen = len;
      });
      column.width = Math.min(maxLen + 2, 60);
    }
    return `Auto-fit column(s) ${op.columns}`;
  }

  // ------------------------------------------------------------- Formatting

  op_format_cells(op) {
    const ws = this.getSheet(op.sheet);
    const r = parseRange(op.range);
    for (let row = r.startRow; row <= r.endRow; row++) {
      for (let col = r.startCol; col <= r.endCol; col++) {
        const cell = ws.getCell(row, col);
        const font = { ...(cell.font || {}) };
        if (op.bold !== undefined) font.bold = !!op.bold;
        if (op.italic !== undefined) font.italic = !!op.italic;
        if (op.underline !== undefined) font.underline = !!op.underline;
        if (op.fontSize !== undefined) font.size = op.fontSize;
        if (op.fontFamily !== undefined) font.name = op.fontFamily;
        if (op.fontColor !== undefined) font.color = { argb: `FF${normalizeColor(op.fontColor)}` };
        if (Object.keys(font).length) cell.font = font;

        if (op.fillColor !== undefined) {
          cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: `FF${normalizeColor(op.fillColor)}` } };
        }
        if (op.border !== undefined && op.border !== "none") {
          const style = BORDER_STYLE_MAP[op.border] || "thin";
          const side = { style };
          cell.border = { top: side, left: side, bottom: side, right: side };
        } else if (op.border === "none") {
          cell.border = {};
        }
        if (op.align !== undefined || op.valign !== undefined || op.wrapText !== undefined) {
          cell.alignment = {
            ...(cell.alignment || {}),
            ...(op.align !== undefined ? { horizontal: op.align } : {}),
            ...(op.valign !== undefined ? { vertical: op.valign } : {}),
            ...(op.wrapText !== undefined ? { wrapText: !!op.wrapText } : {}),
          };
        }
      }
    }
    return `Formatted ${op.sheet}!${op.range}`;
  }

  op_number_format(op) {
    const ws = this.getSheet(op.sheet);
    const r = parseRange(op.range);
    const fmt = op.format === "custom" ? op.customFormat : NUMBER_FORMAT_CODES[op.format];
    if (!fmt) throw new OperationError(`Unknown number format "${op.format}"`, op.type);
    for (let row = r.startRow; row <= r.endRow; row++) {
      for (let col = r.startCol; col <= r.endCol; col++) {
        ws.getCell(row, col).numFmt = fmt;
      }
    }
    return `Applied "${op.format}" number format to ${op.sheet}!${op.range}`;
  }

  // --------------------------------------------------------------- Structure

  op_merge_cells(op) {
    const ws = this.getSheet(op.sheet);
    ws.mergeCells(op.range);
    return `Merged ${op.sheet}!${op.range}`;
  }

  op_unmerge_cells(op) {
    const ws = this.getSheet(op.sheet);
    ws.unMergeCells(op.range);
    return `Unmerged ${op.sheet}!${op.range}`;
  }

  op_freeze_panes(op) {
    const ws = this.getSheet(op.sheet);
    const { col, row } = parseCell(op.cell);
    ws.views = [{ state: "frozen", xSplit: col - 1, ySplit: row - 1, topLeftCell: op.cell, activePane: "bottomRight" }];
    return `Froze panes at ${op.sheet}!${op.cell}`;
  }

  op_create_table(op) {
    const ws = this.getSheet(op.sheet);
    const r = parseRange(op.range);
    const hasHeader = op.hasHeader !== false;
    const columns = [];
    for (let c = r.startCol; c <= r.endCol; c++) {
      const headerCell = ws.getCell(r.startRow, c);
      columns.push({ name: String(headerCell.value !== null && headerCell.value !== undefined ? headerCell.value : `Column${c}`) });
    }
    const rows = [];
    const dataStartRow = hasHeader ? r.startRow + 1 : r.startRow;
    for (let row = dataStartRow; row <= r.endRow; row++) {
      const rowData = [];
      for (let c = r.startCol; c <= r.endCol; c++) rowData.push(ws.getCell(row, c).value);
      rows.push(rowData);
    }
    ws.addTable({
      name: op.name.replace(/[^A-Za-z0-9_]/g, "_"),
      ref: `${numberToCol(r.startCol)}${r.startRow}`,
      headerRow: hasHeader,
      totalsRow: false,
      style: { theme: op.style || "TableStyleMedium2", showRowStripes: true },
      columns,
      rows,
    });
    return `Created table "${op.name}" at ${op.sheet}!${op.range}`;
  }

  op_sort_range(op) {
    const ws = this.getSheet(op.sheet);
    const r = parseRange(op.range);
    const hasHeader = op.hasHeader !== false;
    const dataStartRow = hasHeader ? r.startRow + 1 : r.startRow;
    const sortColIndex = typeof op.column === "number" ? op.column : parseColumnsSpec(op.column).startCol;

    const rows = [];
    for (let row = dataStartRow; row <= r.endRow; row++) {
      const rowData = [];
      for (let c = r.startCol; c <= r.endCol; c++) rowData.push(ws.getCell(row, c).value);
      rows.push(rowData);
    }
    const colOffset = sortColIndex - r.startCol;
    const dir = op.order === "desc" ? -1 : 1;
    rows.sort((a, b) => {
      const av = a[colOffset];
      const bv = b[colOffset];
      if (av === bv) return 0;
      if (av === null || av === undefined) return 1;
      if (bv === null || bv === undefined) return -1;
      if (typeof av === "number" && typeof bv === "number") return (av - bv) * dir;
      return String(av).localeCompare(String(bv)) * dir;
    });
    rows.forEach((rowData, rIdx) => {
      rowData.forEach((val, cIdx) => {
        ws.getCell(dataStartRow + rIdx, r.startCol + cIdx).value = val;
      });
    });
    return `Sorted ${op.sheet}!${op.range} by column ${numberToCol(r.startCol + colOffset)} (${op.order || "asc"})`;
  }

  // ---------------------------------------------------------------- Formulas

  op_set_formula(op) {
    const ws = this.getSheet(op.sheet);
    ws.getCell(op.cell).value = { formula: op.formula.replace(/^=/, "") };
    return `Set formula at ${op.sheet}!${op.cell}`;
  }

  op_add_total_column(op) {
    const ws = this.getSheet(op.sheet);
    const lastCol = (ws.actualColumnCount || ws.columnCount || 1) + 1;
    const targetCol = op.column ? parseColumnsSpec(op.column).startCol : lastCol;
    const lastRow = ws.actualRowCount || ws.rowCount || 1;
    const range = op.sourceRange ? parseRange(op.sourceRange) : { startRow: 2, endRow: lastRow, startCol: 1, endCol: targetCol - 1 };

    ws.getCell(1, targetCol).value = op.header;
    ws.getCell(1, targetCol).font = { bold: true };

    const formulaType = (op.formulaType || "sum").toLowerCase();
    for (let row = range.startRow; row <= range.endRow; row++) {
      const rangeRef = `${numberToCol(range.startCol)}${row}:${numberToCol(range.endCol)}${row}`;
      let formula;
      if (formulaType === "average") formula = `AVERAGE(${rangeRef})`;
      else if (formulaType === "count") formula = `COUNT(${rangeRef})`;
      else formula = `SUM(${rangeRef})`;
      ws.getCell(row, targetCol).value = { formula };
    }
    return `Added "${op.header}" column with ${formulaType.toUpperCase()} formulas`;
  }

  op_add_total_row(op) {
    const ws = this.getSheet(op.sheet);
    const lastRow = (ws.actualRowCount || ws.rowCount || 1) + 1;
    const lastCol = ws.actualColumnCount || ws.columnCount || 1;
    const targetCols = op.columns ? op.columns.map((c) => parseColumnsSpec(c).startCol) : Array.from({ length: lastCol }, (_, i) => i + 1);

    ws.getCell(lastRow, 1).value = op.label || "Total";
    ws.getCell(lastRow, 1).font = { bold: true };
    targetCols.forEach((col) => {
      if (col === 1 && !op.columns) return;
      const colLetter = numberToCol(col);
      ws.getCell(lastRow, col).value = { formula: `SUM(${colLetter}2:${colLetter}${lastRow - 1})` };
      ws.getCell(lastRow, col).font = { bold: true };
    });
    return `Added total row at row ${lastRow}`;
  }

  // ------------------------------------------------------------------ Charts

  op_add_chart(op) {
    // ExcelJS cannot itself write native chart XML, so we record the request
    // and let chartEngine.js inject it into the finished .xlsx as a
    // post-processing step (see server-side apply pipeline).
    this.pendingCharts.push(op);
    return `Queued ${op.chartType} chart "${op.title || ""}" for ${op.sheet}!${op.dataRange}`;
  }

  // ------------------------------------------------------------------- Misc

  op_add_hyperlink(op) {
    const ws = this.getSheet(op.sheet);
    const cell = ws.getCell(op.cell);
    cell.value = { text: op.text || op.url, hyperlink: op.url };
    cell.font = { ...(cell.font || {}), color: { argb: "FF0563C1" }, underline: true };
    return `Added hyperlink at ${op.sheet}!${op.cell}`;
  }

  op_data_validation(op) {
    const ws = this.getSheet(op.sheet);
    const r = parseRange(op.range);
    const validTypes = ["list", "whole", "decimal", "date", "textLength", "custom"];
    const type = validTypes.includes(op.validationType) ? op.validationType : "list";
    for (let row = r.startRow; row <= r.endRow; row++) {
      for (let col = r.startCol; col <= r.endCol; col++) {
        ws.getCell(row, col).dataValidation = {
          type,
          allowBlank: true,
          formulae: [op.formula1, op.formula2].filter(Boolean),
          showErrorMessage: true,
          prompt: op.prompt,
        };
      }
    }
    return `Added data validation to ${op.sheet}!${op.range}`;
  }

  op_conditional_format(op) {
    const ws = this.getSheet(op.sheet);
    const rule = op.rule || {};
    const type = rule.type || "cellIs";
    const style = {};
    if (rule.fillColor) style.fill = { type: "pattern", pattern: "solid", bgColor: { argb: `FF${normalizeColor(rule.fillColor) || "FFFF00"}` } };
    if (rule.fontColor) style.font = { color: { argb: `FF${normalizeColor(rule.fontColor) || "FF0000"}` } };

    ws.addConditionalFormatting({
      ref: op.range,
      rules: [
        {
          type,
          operator: rule.operator || "greaterThan",
          formulae: rule.formulae || [rule.value !== undefined ? String(rule.value) : "0"],
          style,
          priority: 1,
        },
      ],
    });
    return `Added conditional formatting to ${op.sheet}!${op.range}`;
  }

  op_page_setup(op) {
    const ws = this.getSheet(op.sheet);
    ws.pageSetup = {
      ...ws.pageSetup,
      orientation: op.orientation || ws.pageSetup.orientation || "portrait",
      fitToPage: !!(op.fitToWidth || op.fitToHeight),
      fitToWidth: op.fitToWidth || 1,
      fitToHeight: op.fitToHeight || 1,
      margins: op.margins || ws.pageSetup.margins,
    };
    return `Updated page setup for ${op.sheet}`;
  }

  async toBuffer() {
    return this.workbook.xlsx.writeBuffer();
  }
}

module.exports = { ExcelEngine, OperationError, NUMBER_FORMAT_CODES };
