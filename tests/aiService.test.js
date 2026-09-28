"use strict";

const { extractJsonObject, summarizeOperations, planEdit, AiServiceError } = require("../src/services/aiService");
const config = require("../src/config");

describe("aiService", () => {
  const originalFetch = global.fetch;
  afterEach(() => {
    global.fetch = originalFetch;
    config.ai.apiKey = "";
  });

  test("extractJsonObject handles code fences and surrounding prose", () => {
    expect(extractJsonObject('```json\n{"operations":[]}\n```')).toEqual({ operations: [] });
    expect(extractJsonObject('Here you go: {"operations":[{"type":"x"}]} thanks')).toEqual({ operations: [{ type: "x" }] });
    expect(() => extractJsonObject("not json")).toThrow(AiServiceError);
  });

  test("summarizeOperations produces readable lines", () => {
    const lines = summarizeOperations([
      { type: "format_cells", sheet: "S", range: "A1:B1", bold: true },
      { type: "add_chart", sheet: "S", chartType: "pie", dataRange: "A1:B3", title: "T" },
    ]);
    expect(lines[0]).toContain("bold");
    expect(lines[1]).toContain("pie chart");
  });

  test("planEdit fails clearly when no API key is configured", async () => {
    config.ai.apiKey = "";
    await expect(planEdit("make it bold", { sheets: [] })).rejects.toThrow(/ANTHROPIC_API_KEY/);
  });

  test("planEdit returns operations from a mocked provider response", async () => {
    config.ai.apiKey = "test-key";
    global.fetch = jest.fn().mockResolvedValue({
      ok: true,
      json: async () => ({ content: [{ type: "text", text: '{"operations":[{"type":"freeze_panes","sheet":"Sheet1","cell":"A2"}]}' }] }),
    });
    const ops = await planEdit("freeze header", { sheets: [] });
    expect(ops).toHaveLength(1);
    const [, init] = global.fetch.mock.calls[0];
    expect(JSON.parse(init.body).system).toContain("ALLOWED OPERATIONS");
    expect(init.headers["x-api-key"]).toBe("test-key");
  });

  test("planEdit surfaces provider HTTP errors as a friendly error", async () => {
    config.ai.apiKey = "test-key";
    global.fetch = jest.fn().mockResolvedValue({ ok: false, status: 500, text: async () => "boom" });
    await expect(planEdit("x", {})).rejects.toBeInstanceOf(AiServiceError);
  });
});
