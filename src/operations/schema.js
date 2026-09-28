"use strict";

/**
 * This is the single source of truth for what the AI is allowed to ask the
 * Excel engine to do. The AI never executes code directly - it can only
 * produce JSON objects shaped like these, which are validated here before
 * excelEngine.js is allowed to touch a workbook.
 *
 * Each entry: { required: [...fields], optional: [...fields], describe }
 * Field-level type/shape checking happens in operationValidator.js.
 */
const OPERATION_TYPES = {
  create_sheet: { required: ["sheet"], optional: ["index"] },
  rename_sheet: { required: ["sheet", "newName"], optional: [] },
  delete_sheet: { required: ["sheet"], optional: [] },
  reorder_sheets: { required: ["order"], optional: [] },

  set_cell: { required: ["sheet", "cell", "value"], optional: ["formula"] },
  set_range_values: { required: ["sheet", "startCell", "values"], optional: [] },
  clear_cells: { required: ["sheet", "range"], optional: [] },
  copy_cells: { required: ["sheet", "from", "to"], optional: ["targetSheet"] },
  move_cells: { required: ["sheet", "from", "to"], optional: [] },
  find_replace: { required: ["sheet", "find", "replace"], optional: ["matchCase"] },

  insert_rows: { required: ["sheet", "at"], optional: ["count"] },
  delete_rows: { required: ["sheet", "rows"], optional: [] },
  delete_empty_rows: { required: ["sheet"], optional: ["withinRange"] },
  set_row_height: { required: ["sheet", "rows", "height"], optional: [] },

  insert_columns: { required: ["sheet", "at"], optional: ["count"] },
  delete_columns: { required: ["sheet", "columns"], optional: [] },
  set_column_width: { required: ["sheet", "columns", "width"], optional: [] },
  auto_fit_columns: { required: ["sheet", "columns"], optional: [] },

  format_cells: {
    required: ["sheet", "range"],
    optional: [
      "bold", "italic", "underline", "fontSize", "fontFamily", "fontColor",
      "fillColor", "border", "align", "valign", "wrapText",
    ],
  },
  number_format: { required: ["sheet", "range", "format"], optional: [] },

  merge_cells: { required: ["sheet", "range"], optional: [] },
  unmerge_cells: { required: ["sheet", "range"], optional: [] },
  freeze_panes: { required: ["sheet", "cell"], optional: [] },
  create_table: { required: ["sheet", "range", "name"], optional: ["hasHeader", "style"] },
  sort_range: { required: ["sheet", "range", "column"], optional: ["order", "hasHeader"] },

  set_formula: { required: ["sheet", "cell", "formula"], optional: [] },
  add_total_column: {
    required: ["sheet", "header"],
    optional: ["sourceRange", "column", "formulaType"],
  },
  add_total_row: { required: ["sheet"], optional: ["label", "columns"] },

  add_chart: {
    required: ["sheet", "chartType", "dataRange"],
    optional: ["title", "categoryColumn", "valueColumns", "anchorCell", "width", "height"],
  },

  add_hyperlink: { required: ["sheet", "cell", "url"], optional: ["text"] },
  data_validation: { required: ["sheet", "range", "validationType"], optional: ["formula1", "formula2", "prompt"] },
  conditional_format: { required: ["sheet", "range", "rule"], optional: [] },
  page_setup: { required: ["sheet"], optional: ["orientation", "fitToWidth", "fitToHeight", "margins"] },
};

const CHART_TYPES = ["bar", "column", "line", "pie"];
const NUMBER_FORMATS = ["currency", "percentage", "date", "decimal", "integer", "text", "custom"];
const BORDER_STYLES = ["thin", "medium", "thick", "dashed", "dotted", "double", "none"];

module.exports = { OPERATION_TYPES, CHART_TYPES, NUMBER_FORMATS, BORDER_STYLES };
