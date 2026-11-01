import { describe, expect, it } from "vite-plus/test";

import {
  CHART_EX_CONTENT_TYPE,
  CHART_EX_NAMESPACE,
  CHART_EX_RELATIONSHIP_TYPE,
  chartExPartPath,
  chartExParts,
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
