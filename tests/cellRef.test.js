"use strict";

const { colToNumber, numberToCol, parseCell, parseRange, parseColumnsSpec, parseRowsSpec } = require("../src/utils/cellRef");

describe("cellRef utilities", () => {
  test("colToNumber / numberToCol round-trip", () => {
    expect(colToNumber("A")).toBe(1);
    expect(colToNumber("Z")).toBe(26);
    expect(colToNumber("AA")).toBe(27);
    expect(numberToCol(1)).toBe("A");
    expect(numberToCol(26)).toBe("Z");
    expect(numberToCol(27)).toBe("AA");
    for (const n of [1, 5, 26, 27, 52, 703]) {
      expect(colToNumber(numberToCol(n))).toBe(n);
    }
  });

  test("parseCell", () => {
    expect(parseCell("B7")).toEqual({ col: 2, row: 7, colLetter: "B" });
    expect(() => parseCell("bad")).toThrow();
  });

  test("parseRange", () => {
    expect(parseRange("A1:D10")).toEqual({ startCol: 1, startRow: 1, endCol: 4, endRow: 10 });
    expect(parseRange("D10:A1")).toEqual({ startCol: 1, startRow: 1, endCol: 4, endRow: 10 });
    expect(parseRange("B2")).toEqual({ startCol: 2, startRow: 2, endCol: 2, endRow: 2 });
  });

  test("parseColumnsSpec", () => {
    expect(parseColumnsSpec("B")).toEqual({ startCol: 2, endCol: 2 });
    expect(parseColumnsSpec("B:D")).toEqual({ startCol: 2, endCol: 4 });
  });

  test("parseRowsSpec", () => {
    expect(parseRowsSpec("5")).toEqual({ startRow: 5, endRow: 5 });
    expect(parseRowsSpec("5:10")).toEqual({ startRow: 5, endRow: 10 });
  });
});
