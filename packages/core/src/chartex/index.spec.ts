import { describe, expect, it } from "vite-plus/test";

import {
  CHART_EX_CONTENT_TYPE,
  CHART_EX_NAMESPACE,
  CHART_EX_RELATIONSHIP_TYPE,
  CHART_EX_PLACEHOLDER_LABEL,
  chartExExtentFromDrawingXml,
  chartExPartPath,
  chartExParts,
  chartExPlaceholderGeometry,
  chartExPlaceholderOfDrawingXml,
  isChartExDrawingXml,
  isChartExGraphicDataUri,
  isChartExPartPath,
  isChartExRelationshipType,
} from "./index";

describe("chartEx part surface", () => {
  it("exposes the URI, content type, and relationship type Word writes", () => {
    expect(CHART_EX_NAMESPACE).toBe("http://schemas.microsoft.com/office/drawing/2014/chartex");
    expect(CHART_EX_CONTENT_TYPE).toBe("application/vnd.ms-office.chartex+xml");
    expect(CHART_EX_RELATIONSHIP_TYPE).toBe(
      "http://schemas.microsoft.com/office/2014/relationships/chartEx",
    );
  });

  it("builds the part path Word allocates", () => {
    expect(chartExPartPath(1)).toBe("word/charts/chartEx1.xml");
    expect(chartExPartPath(12)).toBe("word/charts/chartEx12.xml");
    expect(chartExPartPath(2, "ppt/charts")).toBe("ppt/charts/chartEx2.xml");
  });

  it("recognizes part paths case-insensitively and with a leading slash", () => {
    expect(isChartExPartPath("word/charts/chartEx1.xml")).toBe(true);
    expect(isChartExPartPath("/word/charts/chartEx7.xml")).toBe(true);
    expect(isChartExPartPath("word/charts/chartex1.xml")).toBe(true);
    expect(isChartExPartPath("word/charts/chart1.xml")).toBe(false);
    expect(isChartExPartPath("word/charts/chartEx1.xml.rels")).toBe(false);
    expect(isChartExPartPath("word/charts/userShapes1.xml")).toBe(false);
  });

  it("matches the chartEx relationship type", () => {
    expect(
      isChartExRelationshipType("http://schemas.microsoft.com/office/2014/relationships/chartEx"),
    ).toBe(true);
    expect(
      isChartExRelationshipType(
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart",
      ),
    ).toBe(false);
  });

  it("accepts every chartEx graphicData URI revision, and nothing else", () => {
    expect(
      isChartExGraphicDataUri("http://schemas.microsoft.com/office/drawing/2014/chartex"),
    ).toBe(true);
    expect(
      isChartExGraphicDataUri("http://schemas.microsoft.com/office/drawing/2015/9/8/chartex"),
    ).toBe(true);
    expect(
      isChartExGraphicDataUri("http://schemas.microsoft.com/office/drawing/2016/5/14/chartex"),
    ).toBe(true);
    // classic charts and pictures are different payloads
    expect(isChartExGraphicDataUri("http://schemas.openxmlformats.org/drawingml/2006/chart")).toBe(
      false,
    );
    expect(
      isChartExGraphicDataUri("http://schemas.openxmlformats.org/drawingml/2006/picture"),
    ).toBe(false);
  });

  it("collects chartEx parts from a package listing", () => {
    const parts = chartExParts([
      "word/document.xml",
      "word/charts/chartEx1.xml",
      "word/charts/chartEx2.xml",
      "word/charts/chart1.xml",
      "_rels/.rels",
    ]);
    expect(parts).toEqual([
      { path: "word/charts/chartEx1.xml", contentType: CHART_EX_CONTENT_TYPE },
      { path: "word/charts/chartEx2.xml", contentType: CHART_EX_CONTENT_TYPE },
    ]);
  });
});

// ── renderer placeholder ──

const W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const WP_NS = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing";
const A_NS = "http://schemas.openxmlformats.org/drawingml/2006/main";
const R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

/** A preserved Word chartEx drawing, as it survives on a rawXml run. */
const CHART_EX_DRAWING =
  `<w:drawing xmlns:w="${W_NS}" xmlns:wp="${WP_NS}" xmlns:a="${A_NS}" xmlns:r="${R_NS}" ` +
  `xmlns:cx="${CHART_EX_NAMESPACE}">` +
  '<wp:inline distT="0" distB="0" distL="0" distR="0">' +
  '<wp:extent cx="5486400" cy="3200400"/>' +
  '<wp:docPr id="42" name="Chart 42"/>' +
  `<a:graphic><a:graphicData uri="${CHART_EX_NAMESPACE}">` +
  '<cx:chart r:id="rIdChartEx"/>' +
  "</a:graphicData></a:graphic>" +
  "</wp:inline></w:drawing>";

describe("chartEx renderer placeholder", () => {
  it("recognizes a preserved chartEx drawing and reads its wp:extent", () => {
    expect(isChartExDrawingXml(CHART_EX_DRAWING)).toBe(true);
    expect(chartExExtentFromDrawingXml(CHART_EX_DRAWING)).toEqual({
      width: 5486400,
      height: 3200400,
    });
  });

  it("does not claim classic chart, picture, or malformed drawings", () => {
    expect(
      isChartExDrawingXml(
        '<w:drawing><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:chart r:id="rId1"/></a:graphicData></w:drawing>',
      ),
    ).toBe(false);
    expect(isChartExDrawingXml("<w:p/>")).toBe(false);
    expect(chartExExtentFromDrawingXml("<w:drawing>")).toBeUndefined();
  });

  it("builds non-blank placeholder geometry for a Word-sized chart", () => {
    const placeholder = chartExPlaceholderGeometry({ width: 5486400, height: 3200400 });
    expect(placeholder.frame.width).toBe(5486400);
    expect(placeholder.frame.height).toBe(3200400);
    expect(placeholder.frame.width).toBeGreaterThan(0);
    expect(placeholder.frame.height).toBeGreaterThan(0);
    expect(placeholder.labelX).toBe(2743200);
    expect(placeholder.labelY).toBe(1600200);
    expect(placeholder.label).toBe(CHART_EX_PLACEHOLDER_LABEL);
    expect(placeholder.label).toBe("modern chart (not rendered)");
  });

  it("keeps a visible frame for a degenerate or missing extent", () => {
    for (const extent of [undefined, { width: 0, height: 0 }, { width: -5, height: Number.NaN }]) {
      const placeholder = chartExPlaceholderGeometry(extent);
      expect(placeholder.frame.width).toBeGreaterThan(0);
      expect(placeholder.frame.height).toBeGreaterThan(0);
      expect(placeholder.labelX).toBeGreaterThan(0);
      expect(placeholder.labelY).toBeGreaterThan(0);
    }
  });

  it("applies an inset without ever going blank", () => {
    const placeholder = chartExPlaceholderGeometry(
      { width: 1000, height: 600 },
      { inset: 50, minSize: 20 },
    );
    expect(placeholder.frame).toEqual({ x: 50, y: 50, width: 900, height: 500 });
    const tiny = chartExPlaceholderGeometry({ width: 10, height: 10 }, { inset: 50, minSize: 20 });
    expect(tiny.frame.width).toBe(20);
    expect(tiny.frame.height).toBe(20);
  });

  it("derives the placeholder from a preserved drawing in one step", () => {
    expect(chartExPlaceholderOfDrawingXml(CHART_EX_DRAWING)).toMatchObject({
      frame: { x: 0, y: 0, width: 5486400, height: 3200400 },
      label: "modern chart (not rendered)",
    });
    expect(chartExPlaceholderOfDrawingXml("<w:p/>")).toBeUndefined();
  });
});
