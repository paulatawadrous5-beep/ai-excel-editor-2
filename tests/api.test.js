"use strict";

// End-to-end workflow through the real HTTP layer with only the AI call mocked.
jest.mock("../src/services/aiService", () => {
  const actual = jest.requireActual("../src/services/aiService");
  return { ...actual, planEdit: jest.fn(), planCreate: jest.fn() };
});

const request = require("supertest");
const ExcelJS = require("exceljs");
const ai = require("../src/services/aiService");
const app = require("../server");

async function makeXlsx() {
  const wb = new ExcelJS.Workbook();
  const ws = wb.addWorksheet("Sheet1");
  ws.addRows([["Month", "Revenue"], ["Jan", 100], ["Feb", 200], ["", ""]]);
  return Buffer.from(await wb.xlsx.writeBuffer());
}

describe("Edit workflow: upload -> plan -> validate -> apply -> verify -> download", () => {
  test("full happy path produces a valid downloadable .xlsx", async () => {
    ai.planEdit.mockResolvedValue([
      { type: "format_cells", sheet: "Sheet1", range: "A1:B1", bold: true, fillColor: "1F4E78" },
      { type: "delete_empty_rows", sheet: "Sheet1" },
      { type: "add_total_column", sheet: "Sheet1", header: "Total" },
      { type: "add_chart", sheet: "Sheet1", chartType: "column", dataRange: "A1:B3", title: "Sales" },
      { type: "bogus", sheet: "Sheet1" },
    ]);
    const preview = await request(app)
      .post("/api/edit/preview")
      .field("instructions", "bold header, delete empty row, add total, chart")
      .attach("file", await makeXlsx(), "sales.xlsx");
    expect(preview.status).toBe(200);
    expect(preview.body.operations).toHaveLength(4);
    expect(preview.body.warnings.join(" ")).toMatch(/unknown operation type/);

    const apply = await request(app).post("/api/edit/apply").send({ sessionId: preview.body.sessionId });
    expect(apply.status).toBe(200);
    expect(apply.body.failedCount).toBe(0);
    expect(apply.body.verifiedCount + apply.body.unverifiedCount).toBe(4);
    expect(["complete", "complete_with_unverified"]).toContain(apply.body.status);

    const dl = await request(app).get(apply.body.downloadUrl).buffer(true).parse((res, cb) => {
      const chunks = [];
      res.on("data", (c) => chunks.push(c));
      res.on("end", () => cb(null, Buffer.concat(chunks)));
    });
    expect(dl.status).toBe(200);
    const out = new ExcelJS.Workbook();
    await out.xlsx.load(dl.body);
    const ws = out.getWorksheet("Sheet1");
    expect(ws.getCell("A1").font.bold).toBe(true);
    expect(ws.actualRowCount).toBe(3);
    expect(ws.getCell("C1").value).toBe("Total");
  });

  test("rejects non-xlsx uploads", async () => {
    const res = await request(app).post("/api/edit/preview").field("instructions", "x").attach("file", Buffer.from("hello"), "evil.txt");
    expect(res.status).toBe(400);
  });

  test("rejects a fake .xlsx that is not a zip", async () => {
    const res = await request(app).post("/api/edit/preview").field("instructions", "x").attach("file", Buffer.from("not a zip at all"), "fake.xlsx");
    expect(res.status).toBe(400);
  });

  test("requires instructions", async () => {
    const res = await request(app).post("/api/edit/preview").attach("file", await makeXlsx(), "a.xlsx");
    expect(res.status).toBe(400);
  });

  test("apply with unknown session returns 404 without leaking internals", async () => {
    const res = await request(app).post("/api/edit/apply").send({ sessionId: "nope" });
    expect(res.status).toBe(404);
    expect(JSON.stringify(res.body)).not.toMatch(/stack|at /);
  });

  test("download rejects path-traversal style ids", async () => {
    const res = await request(app).get("/api/download/..%2F..%2Fetc%2Fpasswd");
    expect([400, 404]).toContain(res.status);
  });
});

describe("Create workflow: describe -> plan -> validate -> build -> verify -> download", () => {
  test("builds a workbook from scratch", async () => {
    ai.planCreate.mockResolvedValue([
      { type: "set_range_values", sheet: "Sheet1", startCell: "A1", values: [["Date", "Category", "Amount"], ["2026-01-01", "Food", 12.5], ["2026-01-02", "Rent", 800]] },
      { type: "format_cells", sheet: "Sheet1", range: "A1:C1", bold: true },
      { type: "add_total_row", sheet: "Sheet1", label: "Total", columns: ["C"] },
      { type: "number_format", sheet: "Sheet1", range: "C2:C4", format: "currency" },
      { type: "add_chart", sheet: "Sheet1", chartType: "pie", dataRange: "B1:C3", title: "Spend" },
    ]);
    const preview = await request(app).post("/api/create/preview").send({ instructions: "expense tracker" });
    expect(preview.status).toBe(200);
    const apply = await request(app).post("/api/create/apply").send({ sessionId: preview.body.sessionId });
    expect(apply.status).toBe(200);
    expect(apply.body.failedCount).toBe(0);
    expect(apply.body.downloadUrl).toMatch(/^\/api\/download\//);
  });

  test("empty instructions are rejected", async () => {
    const res = await request(app).post("/api/create/preview").send({ instructions: "  " });
    expect(res.status).toBe(400);
  });
});

describe("health", () => {
  test("GET /api/health", async () => {
    const res = await request(app).get("/api/health");
    expect(res.body.ok).toBe(true);
  });
});
