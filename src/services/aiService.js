"use strict";

const config = require("../config");
const logger = require("../utils/logger");
const { OPERATION_TYPES, CHART_TYPES, NUMBER_FORMATS, BORDER_STYLES } = require("../operations/schema");

class AiServiceError extends Error {
  constructor(message, userFacing = true) {
    super(message);
    this.name = "AiServiceError";
    this.userFacing = userFacing;
  }
}

function buildSchemaReference() {
  const lines = Object.entries(OPERATION_TYPES).map(([type, spec]) => {
    const req = spec.required.join(", ");
    const opt = spec.optional.length ? ` | optional: ${spec.optional.join(", ")}` : "";
    return `- ${type}(${req}${opt})`;
  });
  return lines.join("\n");
}

const SYSTEM_PROMPT = `You are the planning layer of an Excel-editing application.
Your ONLY job is to translate a user's plain-English request into a JSON array of
structured "operations" drawn EXCLUSIVELY from the allowed operation list below.
You never produce code, scripts, or macros. You never invent operation types or
fields that are not listed. If the user's request is ambiguous, make the single
most reasonable assumption and proceed - do not ask questions, since there is no
follow-up turn available.

ALLOWED OPERATIONS (type(required fields | optional: optional fields)):
${buildSchemaReference()}

Notes:
- Cell references look like "B7"; ranges look like "A1:D10"; columns look like "B" or "B:D"; rows look like "5" or "5:10".
- chartType must be one of: ${CHART_TYPES.join(", ")}.
- number_format's "format" must be one of: ${NUMBER_FORMATS.join(", ")} (use "custom" + "customFormat" for anything else).
- format_cells' "border" must be one of: ${BORDER_STYLES.join(", ")}.
- Colors are 6-digit hex, e.g. "1F4E78".
- For "add a Total column", prefer add_total_column with formulaType "sum"|"average"|"count".
- For "add a chart", the dataRange's first row must be headers, first column categories, remaining columns are value series.
- For "delete the empty row(s)", prefer delete_empty_rows over guessing a row number, unless the user names a specific row.
- Formulas must NOT include a leading "=".
- Respond with ONLY a raw JSON object of the shape {"operations":[...]} - no markdown fences, no commentary, no explanation text before or after.`;

function stripCodeFences(text) {
  return text
    .trim()
    .replace(/^```(json)?/i, "")
    .replace(/```$/, "")
    .trim();
}

function extractJsonObject(text) {
  const cleaned = stripCodeFences(text);
  try {
    return JSON.parse(cleaned);
  } catch {
    const start = cleaned.indexOf("{");
    const end = cleaned.lastIndexOf("}");
    if (start >= 0 && end > start) {
      try {
        return JSON.parse(cleaned.slice(start, end + 1));
      } catch {
        /* fall through */
      }
    }
  }
  throw new AiServiceError("The AI response could not be parsed as JSON");
}

async function callAnthropic(userPrompt) {
  if (!config.ai.apiKey) {
    throw new AiServiceError(
      "AI is not configured on this server (ANTHROPIC_API_KEY is missing). Set it in your .env file.",
      true
    );
  }

  let response;
  try {
    response = await fetch(config.ai.apiBase, {
      method: "POST",
      headers: {
        "Content-Type": "application/json",
        "x-api-key": config.ai.apiKey,
        "anthropic-version": "2023-06-01",
      },
      body: JSON.stringify({
        model: config.ai.model,
        max_tokens: 4096,
        system: SYSTEM_PROMPT,
        messages: [{ role: "user", content: userPrompt }],
      }),
    });
  } catch (err) {
    logger.error("Network error calling AI provider", err);
    throw new AiServiceError("Could not reach the AI provider. Please try again.");
  }

  if (!response.ok) {
    const bodyText = await response.text().catch(() => "");
    logger.error(`AI provider returned HTTP ${response.status}`, bodyText);
    throw new AiServiceError("The AI provider rejected the request. Please try again in a moment.");
  }

  const data = await response.json();
  const textBlocks = (data.content || []).filter((b) => b.type === "text").map((b) => b.text);
  if (!textBlocks.length) {
    throw new AiServiceError("The AI returned an empty response.");
  }
  return textBlocks.join("\n");
}

function buildEditPrompt(instructions, workbookSummary) {
  return `WORKBOOK STRUCTURE:
${JSON.stringify(workbookSummary, null, 2)}

USER REQUEST:
"""${instructions}"""

Produce the JSON operations needed to satisfy the user's request against this workbook.`;
}

function buildCreatePrompt(instructions) {
  return `There is no existing workbook - you are creating a brand new one.
Start by assuming a single worksheet named "Sheet1" already exists (do not
issue a create_sheet for it); use create_sheet only for ADDITIONAL sheets.

USER REQUEST:
"""${instructions}"""

Produce the JSON operations needed to build this workbook from scratch,
including headers, sample structure, formatting, formulas/totals, and a
chart if one was requested.`;
}

async function planEdit(instructions, workbookSummary) {
  const raw = await callAnthropic(buildEditPrompt(instructions, workbookSummary));
  const parsed = extractJsonObject(raw);
  if (!parsed || !Array.isArray(parsed.operations)) {
    throw new AiServiceError("The AI response did not contain an operations array.");
  }
  return parsed.operations;
}

async function planCreate(instructions) {
  const raw = await callAnthropic(buildCreatePrompt(instructions));
  const parsed = extractJsonObject(raw);
  if (!parsed || !Array.isArray(parsed.operations)) {
    throw new AiServiceError("The AI response did not contain an operations array.");
  }
  return parsed.operations;
}

/** Produces a short, user-facing bullet-point summary of a plan (no AI call needed). */
function summarizeOperations(operations) {
  return operations.map((op) => {
    switch (op.type) {
      case "create_sheet": return `Create sheet "${op.sheet}"`;
      case "rename_sheet": return `Rename "${op.sheet}" to "${op.newName}"`;
      case "delete_sheet": return `Delete sheet "${op.sheet}"`;
      case "reorder_sheets": return `Reorder sheets: ${op.order.join(", ")}`;
      case "set_cell": return `Set ${op.sheet}!${op.cell}`;
      case "set_range_values": return `Write values starting at ${op.sheet}!${op.startCell}`;
      case "clear_cells": return `Clear ${op.sheet}!${op.range}`;
      case "copy_cells": return `Copy ${op.sheet}!${op.from} to ${op.to}`;
      case "move_cells": return `Move ${op.sheet}!${op.from} to ${op.to}`;
      case "find_replace": return `Find & replace "${op.find}" -> "${op.replace}"`;
      case "insert_rows": return `Insert ${op.count || 1} row(s) at row ${op.at}`;
      case "delete_rows": return `Delete row(s) ${op.rows}`;
      case "delete_empty_rows": return `Delete empty row(s) in "${op.sheet}"`;
      case "set_row_height": return `Set row height to ${op.height}" for row(s) ${op.rows}`;
      case "insert_columns": return `Insert ${op.count || 1} column(s) at ${op.at}`;
      case "delete_columns": return `Delete column(s) ${op.columns}`;
      case "set_column_width": return `Set column width to ${op.width} for ${op.columns}`;
      case "auto_fit_columns": return `Auto-size column(s) ${op.columns}`;
      case "format_cells": {
        const bits = [];
        if (op.bold) bits.push("bold");
        if (op.italic) bits.push("italic");
        if (op.fillColor) bits.push(`fill #${op.fillColor}`);
        if (op.fontColor) bits.push(`font color #${op.fontColor}`);
        return `Format ${op.sheet}!${op.range}${bits.length ? " (" + bits.join(", ") + ")" : ""}`;
      }
      case "number_format": return `Apply ${op.format} format to ${op.sheet}!${op.range}`;
      case "merge_cells": return `Merge ${op.sheet}!${op.range}`;
      case "unmerge_cells": return `Unmerge ${op.sheet}!${op.range}`;
      case "freeze_panes": return `Freeze panes at ${op.sheet}!${op.cell}`;
      case "create_table": return `Create table "${op.name}" at ${op.sheet}!${op.range}`;
      case "sort_range": return `Sort ${op.sheet}!${op.range}`;
      case "set_formula": return `Set formula at ${op.sheet}!${op.cell}`;
      case "add_total_column": return `Add "${op.header}" total column`;
      case "add_total_row": return `Add total row`;
      case "add_chart": return `Create ${op.chartType} chart${op.title ? ` "${op.title}"` : ""}`;
      case "add_hyperlink": return `Add hyperlink at ${op.sheet}!${op.cell}`;
      case "data_validation": return `Add data validation to ${op.sheet}!${op.range}`;
      case "conditional_format": return `Add conditional formatting to ${op.sheet}!${op.range}`;
      case "page_setup": return `Update page setup for "${op.sheet}"`;
      default: return `${op.type}`;
    }
  });
}

module.exports = { planEdit, planCreate, summarizeOperations, AiServiceError, extractJsonObject };
