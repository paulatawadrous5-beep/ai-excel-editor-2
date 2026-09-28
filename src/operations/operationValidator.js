"use strict";

const { OPERATION_TYPES, CHART_TYPES, NUMBER_FORMATS, BORDER_STYLES } = require("./schema");

const CELL_RE = /^[A-Za-z]{1,3}[1-9][0-9]*$/;
const RANGE_RE = /^[A-Za-z]{1,3}[1-9][0-9]*(:[A-Za-z]{1,3}[1-9][0-9]*)?$/;
const COL_RE = /^[A-Za-z]{1,3}$/;
const COL_RANGE_RE = /^[A-Za-z]{1,3}(:[A-Za-z]{1,3})?$/;
const ROWS_RE = /^[0-9]+(:[0-9]+)?$/;
const HEX_COLOR_RE = /^#?[0-9A-Fa-f]{6}$/;
const MAX_SHEET_NAME_LEN = 31;
const MAX_OPS_PER_REQUEST = 200;

class ValidationIssue {
  constructor(index, type, message) {
    this.index = index;
    this.type = type;
    this.message = message;
  }
}

function isNonEmptyString(v) {
  return typeof v === "string" && v.trim().length > 0;
}

function normalizeColor(c) {
  if (typeof c !== "string") return null;
  const hex = c.startsWith("#") ? c.slice(1) : c;
  return HEX_COLOR_RE.test(c) ? hex.toUpperCase() : null;
}

/**
 * Validates a single operation's field values beyond "does the key exist".
 * Returns an array of human-readable problems (empty = fine).
 */
function checkFieldShapes(op) {
  const problems = [];
  const t = op.type;

  if ("sheet" in op && !isNonEmptyString(op.sheet)) problems.push("`sheet` must be a non-empty string");
  if ("newName" in op && (!isNonEmptyString(op.newName) || op.newName.length > MAX_SHEET_NAME_LEN)) {
    problems.push(`\`newName\` must be 1-${MAX_SHEET_NAME_LEN} characters`);
  }
  if ("cell" in op && !CELL_RE.test(op.cell || "")) problems.push("`cell` is not a valid cell reference (e.g. B2)");
  if ("range" in op && !RANGE_RE.test(op.range || "")) problems.push("`range` is not a valid range (e.g. A1:D10)");
  if ("startCell" in op && !CELL_RE.test(op.startCell || "")) problems.push("`startCell` is invalid");
  if ("from" in op && !RANGE_RE.test(op.from || "")) problems.push("`from` is not a valid cell/range");
  if ("to" in op && !RANGE_RE.test(op.to || "") && !CELL_RE.test(op.to || "")) problems.push("`to` is not a valid cell/range");

  if (t === "set_row_height") {
    if (!ROWS_RE.test(String(op.rows))) problems.push("`rows` must look like 5 or 5:10");
    if (typeof op.height !== "number" || op.height <= 0 || op.height > 500) problems.push("`height` must be a positive number");
  }
  if (t === "delete_rows" && !ROWS_RE.test(String(op.rows))) problems.push("`rows` must look like 5 or 5:10");
  if (t === "insert_rows" && (!Number.isInteger(op.at) || op.at < 1)) problems.push("`at` must be a positive integer row number");
  if (t === "insert_columns" && !(COL_RE.test(op.at) || Number.isInteger(op.at))) problems.push("`at` must be a column letter or index");

  if ((t === "delete_columns" || t === "set_column_width" || t === "auto_fit_columns") && "columns" in op) {
    if (!COL_RANGE_RE.test(String(op.columns))) problems.push("`columns` must look like B or B:D");
  }
  if (t === "set_column_width" && (typeof op.width !== "number" || op.width <= 0 || op.width > 500)) {
    problems.push("`width` must be a positive number");
  }

  if (t === "format_cells") {
    if ("fontColor" in op && op.fontColor && !normalizeColor(op.fontColor)) problems.push("`fontColor` must be a hex color");
    if ("fillColor" in op && op.fillColor && !normalizeColor(op.fillColor)) problems.push("`fillColor` must be a hex color");
    if ("fontSize" in op && (typeof op.fontSize !== "number" || op.fontSize < 1 || op.fontSize > 400)) problems.push("`fontSize` out of range");
    if ("border" in op && op.border && !BORDER_STYLES.includes(op.border)) problems.push(`\`border\` must be one of ${BORDER_STYLES.join(", ")}`);
    if ("align" in op && op.align && !["left", "center", "right"].includes(op.align)) problems.push("`align` must be left/center/right");
    if ("valign" in op && op.valign && !["top", "middle", "bottom"].includes(op.valign)) problems.push("`valign` must be top/middle/bottom");
  }

  if (t === "number_format" && !NUMBER_FORMATS.includes(op.format) && !(op.format === "custom" && isNonEmptyString(op.customFormat))) {
    problems.push(`\`format\` must be one of ${NUMBER_FORMATS.join(", ")}`);
  }

  if (t === "add_chart") {
    if (!CHART_TYPES.includes(op.chartType)) problems.push(`\`chartType\` must be one of ${CHART_TYPES.join(", ")}`);
    if (!RANGE_RE.test(op.dataRange || "")) problems.push("`dataRange` is not a valid range");
    if ("anchorCell" in op && op.anchorCell && !CELL_RE.test(op.anchorCell)) problems.push("`anchorCell` is invalid");
  }

  if ((t === "set_formula" || t === "set_cell") && "formula" in op && op.formula !== undefined) {
    const formula = op.formula;
    if (typeof formula !== "string" || formula.trim().length === 0 || formula.length > 2000) {
      problems.push("`formula` must be a non-empty string");
    } else {
      // A leading "=" is harmless and common - normalize it away rather than
      // rejecting the operation, since the engine writes the formula body
      // to a cell's formula property (it is never eval'd server-side).
      op.formula = formula.replace(/^=+/, "");
    }
  }

  if (t === "reorder_sheets") {
    if (!Array.isArray(op.order) || op.order.some((s) => !isNonEmptyString(s))) {
      problems.push("`order` must be an array of sheet names");
    }
  }

  if (t === "add_hyperlink" && !/^https?:\/\//i.test(op.url || "")) {
    problems.push("`url` must start with http:// or https://");
  }

  return problems;
}

/**
 * Validates an array of candidate operations.
 * @returns {{ valid: object[], issues: ValidationIssue[] }}
 */
function validateOperations(operations) {
  const issues = [];
  const valid = [];

  if (!Array.isArray(operations)) {
    return { valid: [], issues: [new ValidationIssue(-1, "root", "operations must be an array")] };
  }
  if (operations.length === 0) {
    return { valid: [], issues: [new ValidationIssue(-1, "root", "no operations were produced")] };
  }
  if (operations.length > MAX_OPS_PER_REQUEST) {
    return { valid: [], issues: [new ValidationIssue(-1, "root", `too many operations (max ${MAX_OPS_PER_REQUEST})`)] };
  }

  operations.forEach((op, index) => {
    if (!op || typeof op !== "object" || Array.isArray(op)) {
      issues.push(new ValidationIssue(index, "unknown", "operation must be an object"));
      return;
    }
    const spec = OPERATION_TYPES[op.type];
    if (!spec) {
      issues.push(new ValidationIssue(index, op.type || "unknown", `unknown operation type "${op.type}"`));
      return;
    }
    const missing = spec.required.filter((f) => op[f] === undefined || op[f] === null || op[f] === "");
    if (missing.length) {
      issues.push(new ValidationIssue(index, op.type, `missing required field(s): ${missing.join(", ")}`));
      return;
    }
    const allowedFields = new Set(["type", ...spec.required, ...spec.optional]);
    const unknownFields = Object.keys(op).filter((f) => !allowedFields.has(f));
    if (unknownFields.length) {
      issues.push(new ValidationIssue(index, op.type, `unsupported field(s): ${unknownFields.join(", ")} (ignored)`));
      // Non-fatal: strip unknown fields and continue.
      unknownFields.forEach((f) => delete op[f]);
    }

    const shapeProblems = checkFieldShapes(op);
    if (shapeProblems.length) {
      issues.push(new ValidationIssue(index, op.type, shapeProblems.join("; ")));
      return;
    }

    valid.push(op);
  });

  return { valid, issues };
}

module.exports = { validateOperations, normalizeColor, ValidationIssue };
