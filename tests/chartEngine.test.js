"use strict";

const JSZip = require("jszip");
const { ExcelEngine } = require("../src/services/excelEngine");
const { injectCharts } = require("../src/services/chartEngine");

async function buildWorkbookWithChart(chartType) {
  const engine = ExcelEngine.blank();
  engine.workbook.addWorksheet("Sheet1");
  engine.apply({
    type: "set_range_values",
    sheet: "Sheet1",
    startCell: "A1",
    values: [
      ["Month", "Revenue"],
      ["Jan", 100],
      ["Feb", 200],
      ["Mar", 300],
    ],
  });
  engine.apply({ type: "add_chart", sheet: "Sheet1", chartType, dataRange: "A1:B4", title: "Monthly Revenue" });

  const intermediate = await engine.toBuffer();
  const final = await injectCharts(intermediate, engine.workbook, engine.pendingCharts);
  return final;
}

describe("chartEngine.injectCharts", () => {
  test.each(["column", "bar", "line", "pie"])("produces a valid chart part for %s charts", async (chartType) => {
    const buffer = await buildWorkbookWithChart(chartType);
    const zip = await JSZip.loadAsync(buffer);

    const chartFiles = Object.keys(zip.files).filter((f) => /^xl\/charts\/chart\d+\.xml$/.test(f));
    expect(chartFiles.length).toBe(1);

    const chartXml = await zip.file(chartFiles[0]).async("string");
    expect(chartXml).toContain("<c:chartSpace");
    expect(chartXml).toContain("Monthly Revenue");
    expect(chartXml).toContain("Sheet1!$A$2:$A$4"); // category reference
    expect(chartXml).toContain("Sheet1!$B$2:$B$4"); // value reference

    if (chartType === "pie") {
      expect(chartXml).toContain("<c:pieChart>");
    } else if (chartType === "line") {
      expect(chartXml).toContain("<c:lineChart>");
    } else {
      expect(chartXml).toContain("<c:barChart>");
    }
  });

  test("wires a drawing into the worksheet and content types", async () => {
    const buffer = await buildWorkbookWithChart("column");
    const zip = await JSZip.loadAsync(buffer);

    const drawingFiles = Object.keys(zip.files).filter((f) => /^xl\/drawings\/drawing\d+\.xml$/.test(f));
    expect(drawingFiles.length).toBe(1);

    const sheetXml = await zip.file("xl/worksheets/sheet1.xml").async("string");
    expect(sheetXml).toMatch(/<drawing [^>]*r:id="rId\d+"\/>/);

    const contentTypes = await zip.file("[Content_Types].xml").async("string");
    expect(contentTypes).toContain("drawingml.chart+xml");
    expect(contentTypes).toContain("drawing+xml");
  });

  test("does nothing when there are no chart specs", async () => {
    const engine = ExcelEngine.blank();
    engine.workbook.addWorksheet("Sheet1");
    const buffer = await engine.toBuffer();
    const result = await injectCharts(buffer, engine.workbook, []);
    expect(result).toBe(buffer);
  });
});
