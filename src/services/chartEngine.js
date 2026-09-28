"use strict";

const JSZip = require("jszip");
const path = require("path");
const { parseRange, numberToCol } = require("../utils/cellRef");

/**
 * ExcelJS can read charts but cannot write them, so native Excel charts are
 * added here as a post-processing step: we open the .xlsx (a zip) that
 * ExcelJS just produced, and add the extra OOXML parts a chart needs
 * (chart XML, drawing XML, and the relationships/content-type entries that
 * tie them to a worksheet). Excel opens the result as a normal, fully
 * editable native chart - nothing here is a pasted-in image.
 */

const EMU_PER_CELL_COL = 609600; // ~64px column, close enough for a default anchor
const EMU_PER_CELL_ROW = 190500; // ~15pt row

function xmlEscape(str) {
  return String(str)
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&apos;");
}

function safeSheetRefName(sheetName) {
  // Sheet names with spaces or special chars need single-quoting in formulas.
  return /^[A-Za-z_][A-Za-z0-9_.]*$/.test(sheetName) ? sheetName : `'${sheetName.replace(/'/g, "''")}'`;
}

/** Reads the values of a column range from the ExcelJS worksheet for chart caching. */
function readColumnValues(worksheet, col, startRow, endRow) {
  const values = [];
  for (let r = startRow; r <= endRow; r++) {
    const cell = worksheet.getCell(r, col);
    let v = cell.value;
    if (v && typeof v === "object") {
      if (v.result !== undefined) v = v.result;
      else if (v.richText) v = v.richText.map((t) => t.text).join("");
      else if (v instanceof Date) v = v.toISOString().slice(0, 10);
    }
    values.push(v === null || v === undefined ? "" : v);
  }
  return values;
}

function numCacheXml(values, formatCode) {
  const pts = values
    .map((v, i) => (v === "" || v === null || typeof v === "object" ? "" : `<c:pt idx="${i}"><c:v>${xmlEscape(v)}</c:v></c:pt>`))
    .join("");
  return `<c:numCache><c:formatCode>${xmlEscape(formatCode || "General")}</c:formatCode><c:ptCount val="${values.length}"/>${pts}</c:numCache>`;
}

function strCacheXml(values) {
  const pts = values.map((v, i) => `<c:pt idx="${i}"><c:v>${xmlEscape(v)}</c:v></c:pt>`).join("");
  return `<c:strCache><c:ptCount val="${values.length}"/>${pts}</c:strCache>`;
}

function buildSeriesXml({ idx, sheetRef, catRef, catValues, valRef, valValues, seriesName, seriesNameRef, isPie }) {
  const markerXml = isPie ? "" : `<c:marker><c:symbol val="none"/></c:marker>`;
  return `
    <c:ser>
      <c:idx val="${idx}"/>
      <c:order val="${idx}"/>
      <c:tx><c:strRef><c:f>${xmlEscape(seriesNameRef)}</c:f>${strCacheXml([seriesName])}</c:strRef></c:tx>
      ${markerXml}
      <c:cat><c:strRef><c:f>${xmlEscape(catRef)}</c:f>${strCacheXml(catValues)}</c:strRef></c:cat>
      <c:val><c:numRef><c:f>${xmlEscape(valRef)}</c:f>${numCacheXml(valValues)}</c:numRef></c:val>
    </c:ser>`;
}

function buildChartXml(spec) {
  const { chartType, title, series, catAxId, valAxId } = spec;
  const seriesXml = series.map((s, i) => buildSeriesXml({ ...s, idx: i, isPie: chartType === "pie" })).join("\n");

  let plotXml;
  if (chartType === "pie") {
    plotXml = `<c:pieChart><c:varyColors val="1"/>${seriesXml}</c:pieChart>`;
  } else if (chartType === "line") {
    plotXml = `<c:lineChart><c:grouping val="standard"/><c:varyColors val="0"/>${seriesXml}<c:marker val="1"/><c:axId val="${catAxId}"/><c:axId val="${valAxId}"/></c:lineChart>`;
  } else if (chartType === "bar") {
    plotXml = `<c:barChart><c:barDir val="bar"/><c:grouping val="clustered"/><c:varyColors val="0"/>${seriesXml}<c:axId val="${catAxId}"/><c:axId val="${valAxId}"/></c:barChart>`;
  } else {
    // "column" (default)
    plotXml = `<c:barChart><c:barDir val="col"/><c:grouping val="clustered"/><c:varyColors val="0"/>${seriesXml}<c:axId val="${catAxId}"/><c:axId val="${valAxId}"/></c:barChart>`;
  }

  const axesXml = chartType === "pie" ? "" : `
    <c:catAx>
      <c:axId val="${catAxId}"/>
      <c:scaling><c:orientation val="minMax"/></c:scaling>
      <c:delete val="0"/>
      <c:axPos val="${chartType === "bar" ? "l" : "b"}"/>
      <c:crossAx val="${valAxId}"/>
    </c:catAx>
    <c:valAx>
      <c:axId val="${valAxId}"/>
      <c:scaling><c:orientation val="minMax"/></c:scaling>
      <c:delete val="0"/>
      <c:axPos val="${chartType === "bar" ? "b" : "l"}"/>
      <c:crossAx val="${catAxId}"/>
    </c:valAx>`;

  const titleXml = title
    ? `<c:title><c:tx><c:rich><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>${xmlEscape(title)}</a:t></a:r></a:p></c:rich></c:tx><c:overlay val="0"/></c:title>`
    : `<c:autoTitleDeleted val="1"/>`;

  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <c:chart>
    ${titleXml}
    <c:plotArea>
      <c:layout/>
      ${plotXml}
      ${axesXml}
    </c:plotArea>
    <c:legend><c:legendPos val="b"/><c:overlay val="0"/></c:legend>
    <c:plotVisOnly val="1"/>
  </c:chart>
</c:chartSpace>`;
}

function buildDrawingXml({ fromCol, fromRow, toCol, toRow, chartRelId, name }) {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <xdr:twoCellAnchor>
    <xdr:from><xdr:col>${fromCol}</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>${fromRow}</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from>
    <xdr:to><xdr:col>${toCol}</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>${toRow}</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to>
    <xdr:graphicFrame macro="">
      <xdr:nvGraphicFramePr>
        <xdr:cNvPr id="2" name="${xmlEscape(name)}"/>
        <xdr:cNvGraphicFramePr/>
      </xdr:nvGraphicFramePr>
      <xdr:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/></xdr:xfrm>
      <a:graphic>
        <a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart">
          <c:chart xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:id="${chartRelId}"/>
        </a:graphicData>
      </a:graphic>
    </xdr:graphicFrame>
    <xdr:clientData/>
  </xdr:twoCellAnchor>
</xdr:wsDr>`;
}

async function loadXmlText(zip, path, fallback) {
  const file = zip.file(path);
  if (!file) return fallback;
  return file.async("string");
}

function nextIndexFromFiles(fileNames, prefix, ext) {
  let max = 0;
  const re = new RegExp(`^${prefix}(\\d+)\\.${ext}$`);
  fileNames.forEach((name) => {
    const base = name.split("/").pop();
    const m = base.match(re);
    if (m) max = Math.max(max, parseInt(m[1], 10));
  });
  return max + 1;
}

function maxRelIdFrom(relsXml) {
  if (!relsXml) return 0;
  let max = 0;
  const re = /Id="rId(\d+)"/g;
  let m;
  while ((m = re.exec(relsXml))) max = Math.max(max, parseInt(m[1], 10));
  return max;
}

/**
 * Injects one or more charts into a workbook buffer.
 * @param {Buffer} buffer - the .xlsx produced by ExcelEngine.toBuffer()
 * @param {ExcelJS.Workbook} workbook - the same workbook, for reading cached values
 * @param {object[]} chartSpecs - the validated `add_chart` operations
 */
async function injectCharts(buffer, workbook, chartSpecs) {
  if (!chartSpecs || !chartSpecs.length) return buffer;

  const zip = await JSZip.loadAsync(buffer);
  const allFiles = Object.keys(zip.files);

  let contentTypesXml = await loadXmlText(zip, "[Content_Types].xml", "");
  const workbookXml = await loadXmlText(zip, "xl/workbook.xml", "");
  const workbookRelsXml = await loadXmlText(zip, "xl/_rels/workbook.xml.rels", "");

  // Map sheet name -> worksheet part path (e.g. "Sales" -> "xl/worksheets/sheet2.xml")
  const sheetIdToRid = {};
  const sheetNameToRid = {};
  const sheetTagRe = /<sheet\b[^>]*\/>/g;
  (workbookXml.match(sheetTagRe) || []).forEach((tag) => {
    const nameM = tag.match(/name="([^"]*)"/);
    const ridM = tag.match(/r:id="([^"]*)"/);
    if (nameM && ridM) sheetNameToRid[nameM[1]] = ridM[1];
  });
  const relRe = /<Relationship\b[^>]*\/>/g;
  const ridToTarget = {};
  (workbookRelsXml.match(relRe) || []).forEach((tag) => {
    const idM = tag.match(/Id="([^"]*)"/);
    const targetM = tag.match(/Target="([^"]*)"/);
    if (idM && targetM) ridToTarget[idM[1]] = targetM[1];
  });

  function worksheetPathForSheet(name) {
    const rid = sheetNameToRid[name];
    if (!rid) return null;
    const target = ridToTarget[rid];
    if (!target) return null;
    return target.startsWith("/") ? target.slice(1) : `xl/${target.replace(/^\.?\//, "")}`;
  }

  let nextChartIdx = nextIndexFromFiles(allFiles, "chart", "xml");
  let nextDrawingIdx = nextIndexFromFiles(allFiles, "drawing", "xml");
  const contentTypeAdditions = [];

  for (const spec of chartSpecs) {
    const worksheet = workbook.getWorksheet(spec.sheet);
    if (!worksheet) continue; // already reported as a validation/apply failure elsewhere
    const sheetPath = worksheetPathForSheet(spec.sheet);
    if (!sheetPath) continue;

    const range = parseRange(spec.dataRange);
    const catCol = range.startCol;
    const headerRow = range.startRow;
    const dataStartRow = headerRow + 1;
    const dataEndRow = range.endRow;
    const valueCols = [];
    for (let c = range.startCol + 1; c <= range.endCol; c++) valueCols.push(c);

    const catValues = readColumnValues(worksheet, catCol, dataStartRow, dataEndRow).map((v) => (v === "" ? " " : v));
    const sheetRef = safeSheetRefName(spec.sheet);
    const catRef = `${sheetRef}!$${numberToCol(catCol)}$${dataStartRow}:$${numberToCol(catCol)}$${dataEndRow}`;

    const series = (valueCols.length ? valueCols : [catCol + 1]).map((col) => {
      const headerCell = worksheet.getCell(headerRow, col);
      const seriesName = headerCell.value !== null && headerCell.value !== undefined ? String(headerCell.value) : `Series ${col}`;
      const valValues = readColumnValues(worksheet, col, dataStartRow, dataEndRow).map((v) => (typeof v === "number" ? v : Number(v) || 0));
      return {
        sheetRef,
        catRef,
        catValues,
        valRef: `${sheetRef}!$${numberToCol(col)}$${dataStartRow}:$${numberToCol(col)}$${dataEndRow}`,
        valValues,
        seriesName,
        seriesNameRef: `${sheetRef}!$${numberToCol(col)}$${headerRow}`,
      };
    });

    // Pie charts only support a single series.
    const finalSeries = spec.chartType === "pie" ? [series[0]] : series;

    const catAxId = 100000000 + nextChartIdx * 2;
    const valAxId = catAxId + 1;
    const chartXml = buildChartXml({
      chartType: spec.chartType,
      title: spec.title,
      series: finalSeries,
      catAxId,
      valAxId,
    });

    const chartIdx = nextChartIdx++;
    const chartPath = `xl/charts/chart${chartIdx}.xml`;
    zip.file(chartPath, chartXml);
    contentTypeAdditions.push(
      `<Override PartName="/${chartPath}" ContentType="application/vnd.openxmlformats-officedocument.drawingml.chart+xml"/>`
    );

    // Anchor: default a few cells to the right of the data, or the requested anchorCell.
    let fromCol = range.endCol + 1;
    let fromRow = Math.max(headerRow - 1, 0);
    if (spec.anchorCell) {
      const anchor = parseRange(spec.anchorCell);
      fromCol = anchor.startCol - 1;
      fromRow = anchor.startRow - 1;
    }
    const widthCols = Math.round((spec.width || 480) / (EMU_PER_CELL_COL / 9525));
    const heightRows = Math.round((spec.height || 300) / (EMU_PER_CELL_ROW / 9525));
    const toCol = fromCol + Math.max(widthCols, 6);
    const toRow = fromRow + Math.max(heightRows, 12);

    const sheetXml = await loadXmlText(zip, sheetPath, "");
    const sheetRelsPath = sheetPath.replace(/^(.*)\/([^/]+)$/, "$1/_rels/$2.rels");
    const existingDrawingTagM = sheetXml.match(/<drawing\b[^>]*r:id="([^"]*)"[^>]*\/>/);

    if (existingDrawingTagM) {
      // The sheet already has a drawing (an image, or a chart from the
      // original file) - append our chart as an additional anchor inside
      // that SAME drawing part instead of creating a second, unreferenced
      // one, so existing content is fully preserved.
      const existingRid = existingDrawingTagM[1];
      const sheetRelsXml = await loadXmlText(zip, sheetRelsPath, "");
      const targetM = sheetRelsXml.match(new RegExp(`Id="${existingRid}"[^>]*Target="([^"]*)"`));
      const drawingTarget = targetM ? targetM[1] : null;
      // Sheet-rels targets are relative to xl/worksheets/, e.g. "../drawings/drawing1.xml".
      const drawingPath = drawingTarget
        ? path.posix.normalize(path.posix.join("xl/worksheets", drawingTarget)).replace(/^(\.\.\/)+/, "")
        : null;

      if (drawingPath && zip.file(drawingPath)) {
        const drawingRelsPath = drawingPath.replace(/^(.*)\/([^/]+)$/, "$1/_rels/$2.rels");
        let drawingRelsXml = await loadXmlText(zip, drawingRelsPath, null);
        let localRid;
        if (drawingRelsXml === null) {
          localRid = "rId1";
          drawingRelsXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="${localRid}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart" Target="../charts/chart${chartIdx}.xml"/>
</Relationships>`;
        } else {
          localRid = `rId${maxRelIdFrom(drawingRelsXml) + 1}`;
          drawingRelsXml = drawingRelsXml.replace(
            "</Relationships>",
            `<Relationship Id="${localRid}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart" Target="../charts/chart${chartIdx}.xml"/></Relationships>`
          );
        }
        zip.file(drawingRelsPath, drawingRelsXml);

        const anchorOnlyXml = buildDrawingXml({ fromCol, fromRow, toCol, toRow, chartRelId: localRid, name: spec.title || "Chart" })
          .replace(/^[\s\S]*?<xdr:twoCellAnchor>/, "<xdr:twoCellAnchor>")
          .replace(/<\/xdr:wsDr>\s*$/, "");
        const existingDrawingXml = await loadXmlText(zip, drawingPath, "");
        zip.file(drawingPath, existingDrawingXml.replace(/<\/xdr:wsDr>\s*$/, `${anchorOnlyXml}</xdr:wsDr>`));
      }
      // sheetXml/sheetRels already correctly reference the existing drawing - nothing more to do there.
      continue;
    }

    const drawingIdx = nextDrawingIdx++;
    const drawingPath = `xl/drawings/drawing${drawingIdx}.xml`;
    const drawingRelsPath = `xl/drawings/_rels/drawing${drawingIdx}.xml.rels`;
    zip.file(
      drawingRelsPath,
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart" Target="../charts/chart${chartIdx}.xml"/>
</Relationships>`
    );
    zip.file(drawingPath, buildDrawingXml({ fromCol, fromRow, toCol, toRow, chartRelId: "rId1", name: spec.title || "Chart" }));

    // Wire the new drawing into the worksheet: rels file + <drawing/> element.
    let sheetRelsXml = await loadXmlText(zip, sheetRelsPath, null);
    let rId;
    if (sheetRelsXml === null) {
      rId = "rId1";
      sheetRelsXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="${rId}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing" Target="../drawings/drawing${drawingIdx}.xml"/>
</Relationships>`;
    } else {
      const maxId = maxRelIdFrom(sheetRelsXml);
      rId = `rId${maxId + 1}`;
      sheetRelsXml = sheetRelsXml.replace(
        "</Relationships>",
        `<Relationship Id="${rId}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing" Target="../drawings/drawing${drawingIdx}.xml"/></Relationships>`
      );
    }
    zip.file(sheetRelsPath, sheetRelsXml);

    // Declare the relationships namespace on the element itself: not every
    // worksheet root declares xmlns:r, and an unbound prefix is invalid XML.
    const drawingTag = `<drawing xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:id="${rId}"/>`;
    let newSheetXml;
    if (/<tableParts\b/.test(sheetXml)) {
      newSheetXml = sheetXml.replace(/<tableParts\b/, `${drawingTag}<tableParts`);
    } else if (/<extLst>/.test(sheetXml)) {
      newSheetXml = sheetXml.replace(/<extLst>/, `${drawingTag}<extLst>`);
    } else {
      newSheetXml = sheetXml.replace(/<\/worksheet>\s*$/, `${drawingTag}</worksheet>`);
    }
    zip.file(sheetPath, newSheetXml);

    contentTypeAdditions.push(
      `<Override PartName="/${drawingPath}" ContentType="application/vnd.openxmlformats-officedocument.drawing+xml"/>`
    );
  }

  if (contentTypeAdditions.length) {
    contentTypesXml = contentTypesXml.replace("</Types>", `${contentTypeAdditions.join("")}</Types>`);
    zip.file("[Content_Types].xml", contentTypesXml);
  }

  return zip.generateAsync({ type: "nodebuffer" });
}

module.exports = { injectCharts };
