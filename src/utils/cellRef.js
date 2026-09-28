"use strict";

/** Converts a 1-based column number to its letter form (1 -> A, 27 -> AA). */
function numberToCol(n) {
  let s = "";
  let num = n;
  while (num > 0) {
    const rem = (num - 1) % 26;
    s = String.fromCharCode(65 + rem) + s;
    num = Math.floor((num - 1) / 26);
  }
  return s;
}

/** Converts a column letter (case-insensitive) to its 1-based number. */
function colToNumber(letters) {
  const s = String(letters).toUpperCase();
  let n = 0;
  for (let i = 0; i < s.length; i++) {
    n = n * 26 + (s.charCodeAt(i) - 64);
  }
  return n;
}

/** Parses "B7" -> { col: 2, row: 7, colLetter: "B" } */
function parseCell(ref) {
  const m = String(ref).match(/^([A-Za-z]{1,3})([0-9]+)$/);
  if (!m) throw new Error(`Invalid cell reference: ${ref}`);
  return { col: colToNumber(m[1]), row: parseInt(m[2], 10), colLetter: m[1].toUpperCase() };
}

/** Parses "A1:C10" or a single cell "A1" -> { startCol, startRow, endCol, endRow } */
function parseRange(ref) {
  const parts = String(ref).split(":");
  const start = parseCell(parts[0]);
  const end = parts[1] ? parseCell(parts[1]) : start;
  return {
    startCol: Math.min(start.col, end.col),
    startRow: Math.min(start.row, end.row),
    endCol: Math.max(start.col, end.col),
    endRow: Math.max(start.row, end.row),
  };
}

/** Parses a columns spec: "B" or "B:D" -> { startCol, endCol } */
function parseColumnsSpec(spec) {
  const parts = String(spec).split(":");
  const startCol = colToNumber(parts[0]);
  const endCol = parts[1] ? colToNumber(parts[1]) : startCol;
  return { startCol: Math.min(startCol, endCol), endCol: Math.max(startCol, endCol) };
}

/** Parses a rows spec: "5" or "5:10" -> { startRow, endRow } */
function parseRowsSpec(spec) {
  const parts = String(spec).split(":").map((p) => parseInt(p, 10));
  const startRow = parts[0];
  const endRow = parts[1] !== undefined ? parts[1] : startRow;
  return { startRow: Math.min(startRow, endRow), endRow: Math.max(startRow, endRow) };
}

module.exports = { numberToCol, colToNumber, parseCell, parseRange, parseColumnsSpec, parseRowsSpec };
